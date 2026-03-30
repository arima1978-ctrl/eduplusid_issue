"""
eduplus管理画面の自動操作モジュール
- 新規塾登録
- ID/パスワード取得
"""
import sys
import re
import time
import random
import string

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

LOGIN_URL = 'https://www.eduplus.jp/eduplus/idreg/master/master_login.aspx'
REGISTER_URL = 'https://www.eduplus.jp/eduplus/idreg/master/juku_register.aspx'
APPLY_LIST_URL = 'https://www.eduplus.jp/eduplus/idreg/master/apply_list.aspx'
APPLY_STATE_URL = 'https://www.eduplus.jp/eduplus/idreg/master/apply_state_new.aspx'

MASTER_ID = 'master1'
MASTER_PW = 'kanrisha1650'


def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--remote-debugging-port=9222')
    options.add_argument('--disable-software-rasterizer')
    # VPS(snap版Chromium)対応
    import shutil
    if shutil.which('chromium-browser'):
        options.binary_location = shutil.which('chromium-browser')
    elif shutil.which('chromium'):
        options.binary_location = shutil.which('chromium')
    return webdriver.Chrome(options=options)


def login(driver):
    driver.get(LOGIN_URL)
    time.sleep(3)
    driver.find_element(By.ID, 'ms_index').send_keys(MASTER_ID)
    driver.find_element(By.ID, 'ms_pass').send_keys(MASTER_PW)
    driver.find_element(By.ID, 'bms_login').click()
    time.sleep(3)


def generate_juku_id(juku_name):
    """塾名からローマ字3-5文字のIDを生成"""
    # 簡易的にランダム英数字5文字
    chars = string.ascii_lowercase + string.digits
    return ''.join(random.choice(chars) for _ in range(5))


def register_juku(juku_id, juku_name):
    """
    新規塾を登録してID・パスワードを取得
    Returns: dict with admin_id, admin_pw, sample_id, sample_pw or None on failure
    """
    driver = create_driver()
    try:
        login(driver)

        # 新規塾登録画面
        driver.get(REGISTER_URL)
        time.sleep(3)

        # 塾ID入力
        jid_field = driver.find_element(By.ID, 'jid')
        jid_field.clear()
        jid_field.send_keys(juku_id)

        # 塾名入力
        jname_field = driver.find_element(By.ID, 'jname')
        jname_field.clear()
        jname_field.send_keys(juku_name)

        # パッケージ: パターン2（フルパッケージ）
        pkg_select = Select(driver.find_element(By.ID, 'ddlPackageList'))
        pkg_select.select_by_value('2')
        time.sleep(1)

        # 登録区分: 体験塾
        entry_select = Select(driver.find_element(By.ID, 'selEntryType'))
        entry_select.select_by_value('1')

        # 登録ボタンを有効化してクリック
        driver.execute_script("document.getElementById('bj_register').disabled = false;")
        driver.execute_script("doJukuRegister();")
        time.sleep(5)

        # アラート処理（エラーメッセージを取得）
        alert_text = ''
        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            alert.accept()
            time.sleep(3)
        except:
            pass

        # 登録成功確認
        suc_mes = driver.find_element(By.ID, 'suc_mes_jr')
        if 'display: block' not in (suc_mes.get_attribute('style') or ''):
            # エラー内容を判定して返す
            error_type = 'unknown'
            if 'ID' in alert_text or 'id' in alert_text or '既に' in alert_text:
                error_type = 'id_duplicate'
            elif '塾名' in alert_text or '名前' in alert_text:
                error_type = 'name_duplicate'

            # ページ内のエラーメッセージも確認
            page_text = driver.find_element(By.TAG_NAME, 'body').text
            if '既に登録' in page_text or '重複' in page_text:
                if 'ID' in page_text:
                    error_type = 'id_duplicate'
                elif '塾名' in page_text:
                    error_type = 'name_duplicate'

            return {'error': True, 'error_type': error_type, 'error_message': alert_text or page_text[:200]}

        # 塾申込状況確認画面でindex番号を取得
        driver.get(APPLY_LIST_URL)
        time.sleep(3)

        # 塾IDで検索
        search_field = driver.find_element(By.ID, 'jid')
        search_field.clear()
        search_field.send_keys(juku_id)
        driver.find_element(By.ID, 'search_list').click()
        time.sleep(3)

        # 承認操作ボタンからindex番号を取得
        source = driver.page_source
        match = re.search(r'goApplicationFromNew\((\d+)\)', source)
        if not match:
            return None

        index_num = match.group(1)

        # apply_state_new.aspx でパスワードを取得
        driver.get(f'{APPLY_STATE_URL}?index={index_num}')
        time.sleep(3)

        source = driver.page_source

        # パスワードを抽出
        pw_match = re.search(
            r'【管理者ID/PW】(\S+)\s*/\s*(\S+)\s*【サンプルID/PW】(\S+)\s*/\s*([A-Za-z0-9]+)',
            source
        )

        if pw_match:
            return {
                'admin_id': pw_match.group(1),
                'admin_pw': pw_match.group(2),
                'sample_id': pw_match.group(3),
                'sample_pw': pw_match.group(4),
            }

        return None

    except Exception as e:
        print(f"eduplus登録エラー: {e}")
        return None
    finally:
        driver.quit()


if __name__ == '__main__':
    if sys.platform == 'win32':
        sys.stdout.reconfigure(encoding='utf-8')

    # テスト
    juku_id = 'at' + ''.join(random.choice(string.digits) for _ in range(3))
    print(f"テスト登録: ID={juku_id}, 名前=自動テスト塾")
    result = register_juku(juku_id, '自動テスト塾')
    if result:
        print(f"成功!")
        print(f"  管理者ID: {result['admin_id']}")
        print(f"  管理者PW: {result['admin_pw']}")
        print(f"  サンプルID: {result['sample_id']}")
        print(f"  サンプルPW: {result['sample_pw']}")
    else:
        print("失敗")

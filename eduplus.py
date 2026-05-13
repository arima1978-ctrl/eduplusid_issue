"""
eduplus管理画面の自動操作モジュール
- 新規塾登録
- ID/パスワード取得
"""
import logging
import sys
import re
import time
import random
import string
import os

from bs4 import BeautifulSoup
from dotenv import load_dotenv

logger = logging.getLogger(__name__)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

MIN_JUKU_ID_LENGTH = 3

load_dotenv()

LOGIN_URL = 'https://www.eduplus.jp/eduplus/idreg/master/master_login.aspx'
REGISTER_URL = 'https://www.eduplus.jp/eduplus/idreg/master/juku_register.aspx'
APPLY_LIST_URL = 'https://www.eduplus.jp/eduplus/idreg/master/apply_list.aspx'
APPLY_STATE_URL = 'https://www.eduplus.jp/eduplus/idreg/master/apply_state_new.aspx'

MASTER_ID = os.environ.get('EDUPLUS_MASTER_ID', 'master1')
MASTER_PW = os.environ['EDUPLUS_MASTER_PW']


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


def _find_index_for_juku_id(soup, juku_id):
    """apply_list.aspx の検索結果から、塾ID列が完全一致する行の index を返す。

    eduplus 側の jid 検索は部分一致のため、最初の goApplicationFromNew(N) を
    そのまま使うと別塾の index を拾うことがある（重複ID事象の温床）。
    行単位で BeautifulSoup でパースし、塾ID列の文字列が juku_id と完全一致する
    行のみを採用する。複数一致した場合は None を返す（安全側に倒す）。
    """
    matches = []
    for row in soup.select("table#apply_list tr"):
        cells = [td.get_text(strip=True) for td in row.find_all("td")]
        if juku_id not in cells:
            continue
        btn = row.find("input", {"value": "承認操作"})
        if not btn:
            continue
        m = re.search(r'goApplicationFromNew\((\d+)\)', btn.get("onclick", ""))
        if not m:
            continue
        matches.append(m.group(1))

    if len(matches) == 1:
        return matches[0]
    return None


def register_juku(juku_id, juku_name):
    """
    新規塾を登録してID・パスワードを取得
    Returns: dict with admin_id, admin_pw, sample_id, sample_pw or None on failure
    """
    if len(juku_id) < MIN_JUKU_ID_LENGTH:
        return {
            'error': True,
            'error_type': 'id_too_short',
            'error_message': f'塾IDは{MIN_JUKU_ID_LENGTH}文字以上必要です（指定: "{juku_id}"）',
        }

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

        # 行単位でパースし、塾ID列が完全一致する行の index を採用する。
        # （部分一致で別塾の goApplicationFromNew(N) を拾うのを防ぐ）
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        index_num = _find_index_for_juku_id(soup, juku_id)
        if not index_num:
            logger.error(
                "apply_list で塾ID完全一致の行が見つかりません: juku_id=%s, juku_name=%s",
                juku_id, juku_name,
            )
            return {
                'error': True,
                'error_type': 'juku_not_found',
                'error_message': (
                    f'登録は成功しましたが、塾ID "{juku_id}" の行が apply_list 一覧で '
                    '完全一致しません。eduplus管理画面で手動確認してください。'
                ),
            }

        # apply_state_new.aspx でパスワードを取得
        driver.get(f'{APPLY_STATE_URL}?index={index_num}')
        time.sleep(3)

        source = driver.page_source

        # 検証: 開いたページに juku_id と juku_name が両方含まれることを確認。
        # 含まれない場合は別塾のページを開いている可能性があるため、誤った
        # 管理者ID/PW を返さないよう中断する。
        if juku_id not in source or juku_name not in source:
            logger.error(
                "apply_state_new ページ検証失敗: index=%s, juku_id=%s, juku_name=%s, "
                "juku_id_in_page=%s, juku_name_in_page=%s",
                index_num, juku_id, juku_name,
                juku_id in source, juku_name in source,
            )
            return {
                'error': True,
                'error_type': 'page_verification_failed',
                'error_message': (
                    f'apply_state_new ページ（index={index_num}）が期待する塾と一致しません。'
                    '別塾の管理者ID/PWを誤って取得する可能性があるため中断しました。'
                ),
            }

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
        logger.error("eduplus登録エラー: %s", e, exc_info=True)
        return None
    finally:
        driver.quit()


def delete_juku(juku_id):
    """
    登録済みの塾を削除する
    Returns: True on success, False on failure
    """
    driver = create_driver()
    try:
        login(driver)

        # 塾申込状況確認画面で検索
        driver.get(APPLY_LIST_URL)
        time.sleep(3)

        search_field = driver.find_element(By.ID, 'jid')
        search_field.clear()
        search_field.send_keys(juku_id)
        driver.find_element(By.ID, 'search_list').click()
        time.sleep(3)

        # index番号を取得
        source = driver.page_source
        match = re.search(r'goApplicationFromNew\((\d+)\)', source)
        if not match:
            return False

        index_num = match.group(1)

        # apply_state_new.aspx で削除ボタンをクリック
        driver.get(f'{APPLY_STATE_URL}?index={index_num}')
        time.sleep(3)

        try:
            del_btn = driver.find_element(By.ID, 'bj_delete')
            driver.execute_script("arguments[0].disabled = false;", del_btn)
            del_btn.click()
            time.sleep(2)
        except:
            # 削除ボタンIDが異なる場合はJS実行
            driver.execute_script("doDelete();")
            time.sleep(2)

        # アラート（確認ダイアログ）を承認
        try:
            alert = driver.switch_to.alert
            alert.accept()
            time.sleep(3)
        except:
            pass

        return True

    except Exception as e:
        print(f"eduplus削除エラー: {e}")
        return False
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

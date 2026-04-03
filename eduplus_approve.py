import os
import requests
from bs4 import BeautifulSoup
import re
from datetime import datetime
from zoneinfo import ZoneInfo
from dotenv import load_dotenv

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(SCRIPT_DIR, ".env"))

TELEGRAM_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
EDUPLUS_CHAT_ID = int(os.environ.get("EDUPLUS_CHAT_ID", "-5126783705"))
BASE_URL = "https://www.eduplus.jp/eduplus/idreg/master/"
LOGIN_ID = os.environ.get("EDUPLUS_MASTER_ID", "master1")
LOGIN_PW = os.environ["EDUPLUS_MASTER_PW"]
JST = ZoneInfo("Asia/Tokyo")


def send_telegram(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    for i in range(0, len(text), 4000):
        resp = requests.post(url, json={"chat_id": EDUPLUS_CHAT_ID, "text": text[i:i+4000]})
        if not resp.ok:
            print(f"Telegram送信失敗: status={resp.status_code}, body={resp.text}")


def login(session):
    url = BASE_URL + "master_login.aspx"
    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")
    data = {
        "__VIEWSTATE": soup.find("input", {"name": "__VIEWSTATE"})["value"],
        "__VIEWSTATEGENERATOR": soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"],
        "__EVENTVALIDATION": soup.find("input", {"name": "__EVENTVALIDATION"})["value"],
        "__EVENTTARGET": "bms_login",
        "__EVENTARGUMENT": "",
        "ms_index": LOGIN_ID,
        "ms_pass": LOGIN_PW,
    }
    session.post(url, data=data, allow_redirects=True)


def get_unapproved_list(session):
    """未承認一覧を取得し、各塾のindex・名前を返す"""
    url = BASE_URL + "apply_list.aspx"
    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")
    search_data = {
        "__VIEWSTATE": soup.find("input", {"name": "__VIEWSTATE"})["value"],
        "__VIEWSTATEGENERATOR": soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"],
        "__EVENTVALIDATION": soup.find("input", {"name": "__EVENTVALIDATION"})["value"],
        "__EVENTTARGET": "search_list",
        "__EVENTARGUMENT": "",
        "jid": "",
        "confirmed": "confirmed2",
        "confirmed2": "confirmed2_3",
        "RadBuildAdmin": "RadBuildAdmin3",
        "SelSearchEntryType": "Z",
        "SelSetPrefe": "Z",
        "AgentSearchCombobox": "A",
    }
    r2 = session.post(url, data=search_data)
    soup2 = BeautifulSoup(r2.text, "html.parser")

    # ページング: 全ページ分の塾を取得
    entries = []
    page = 1
    while True:
        rows = soup2.select("table#apply_list tr")
        for row in rows:
            btn = row.find("input", {"value": "承認操作"})
            if not btn:
                continue
            m = re.search(r'goApplicationFromNew\((\d+)\)', btn.get("onclick", ""))
            if not m:
                continue
            idx = m.group(1)
            link = row.find("a", onclick=re.compile(r'goApplicationStateNew'))
            juku_name = link.get_text(strip=True) if link else f"index={idx}"
            entries.append({"index": idx, "juku_name": juku_name})

        # Check next page
        next_link = soup2.find("a", string=re.compile(r"次ページ"))
        if not next_link:
            break
        page += 1
        # Navigate to next page via postback
        vs = soup2.find("input", {"name": "__VIEWSTATE"})
        vsg = soup2.find("input", {"name": "__VIEWSTATEGENERATOR"})
        ev = soup2.find("input", {"name": "__EVENTVALIDATION"})
        if not vs:
            break
        page_data = {
            "__VIEWSTATE": vs["value"],
            "__VIEWSTATEGENERATOR": vsg["value"],
            "__EVENTVALIDATION": ev["value"],
            "__EVENTTARGET": "btnHiddenSetValue",
            "__EVENTARGUMENT": "",
            "curPg": str(page),
            "jid": "",
            "confirmed": "confirmed2",
            "confirmed2": "confirmed2_3",
            "RadBuildAdmin": "RadBuildAdmin3",
            "SelSearchEntryType": "Z",
            "SelSetPrefe": "Z",
            "AgentSearchCombobox": "A",
        }
        r3 = session.post(url, data=page_data)
        soup2 = BeautifulSoup(r3.text, "html.parser")

    return entries


def approve_juku(session, index):
    """指定塾の未承認を一括承認し、承認件数を返す。ページングに対応しループで全件処理。"""
    url = f"{BASE_URL}apply_manager_new.aspx?index={index}"
    total_count = 0
    max_rounds = 20  # 安全上限

    for _ in range(max_rounds):
        r = session.get(url)
        soup = BeautifulSoup(r.text, "html.parser")

        # btn_doApply* ボタン、または「申」ステータスの行で未承認件数を判定
        approve_btns = soup.find_all("input", {"value": "承認", "name": re.compile(r"btn_doApply\d+")})
        pending_rows = soup.select("table#juku_apply tr")
        pending_count = sum(1 for row in pending_rows
                           for td in row.find_all("td")
                           if td.get_text(strip=True) == "申")

        count = len(approve_btns) if approve_btns else pending_count
        if count == 0:
            break

        total_count += count

        post_data = {
            "__VIEWSTATE": soup.find("input", {"name": "__VIEWSTATE"})["value"],
            "__VIEWSTATEGENERATOR": soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"],
            "__EVENTVALIDATION": soup.find("input", {"name": "__EVENTVALIDATION"})["value"],
            "__EVENTTARGET": "btnHiddenApplyAll",
            "__EVENTARGUMENT": "",
            "index": str(index),
        }
        session.post(url, data=post_data)

    return total_count


def run_approve():
    """一括承認を実行し、結果メッセージを返す。"""
    now = datetime.now(JST).strftime("%Y/%m/%d %H:%M")
    session = requests.Session()

    try:
        login(session)

        results = {}
        total = 0
        max_passes = 10  # 安全上限

        for pass_num in range(max_passes):
            entries = get_unapproved_list(session)
            if not entries:
                break

            found_any = False
            for entry in entries:
                count = approve_juku(session, entry["index"])
                if count > 0:
                    name = entry["juku_name"]
                    results[name] = results.get(name, 0) + count
                    total += count
                    found_any = True

            if not found_any:
                break

        if total > 0:
            lines = [f"✅ [{now}] eduplus一括承認完了\n合計: {total}件\n"]
            for name, count in results.items():
                lines.append(f"• {name}: {count}件")
            return "\n".join(lines)
        else:
            return f"📋 [{now}] eduplus承認\n未承認なし"

    except Exception as e:
        return f"❌ [{now}] eduplus承認エラー\n{e}"


def main():
    msg = run_approve()
    send_telegram(msg)
    print(msg)


if __name__ == "__main__":
    main()

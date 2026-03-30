import os
import requests
from bs4 import BeautifulSoup
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from dotenv import load_dotenv

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(SCRIPT_DIR, ".env"))

TELEGRAM_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
EDUPLUS_CHAT_ID = int(os.environ.get("EDUPLUS_CHAT_ID", "-5126783705"))
BASE_MASTER = "https://www.eduplus.jp/eduplus/idreg/master/"
BASE_JUKU = "https://www.eduplus.jp/eduplus/"
LOGIN_ID = os.environ.get("EDUPLUS_MASTER_ID", "master1")
LOGIN_PW = os.environ["EDUPLUS_MASTER_PW"]
JST = ZoneInfo("Asia/Tokyo")
TRIAL_MONTHS = 3  # 体験期間の上限

# 除外する塾名（部分一致）
EXCLUDE_NAMES = [
    "アン進学ジム",
    "アン算国クラブ",
    "アン学習塾",
    "文理受験スクール",
    "星煌学院",
    "キタン塾",
    "学書",
    "進学塾サミット",
    "名大SKY",
    "名大ＳＫＹ",
    "ＭＤＰＳ",
    "MDPS",
    "小笠原",
    "明聖研究所",
    "誠伸社",
    "体験用",
    "サクセス",
    "学習塾SEEDS",
]


def send_telegram(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    for i in range(0, len(text), 4000):
        requests.post(url, json={"chat_id": EDUPLUS_CHAT_ID, "text": text[i:i+4000]})


def master_login():
    """運営者としてログインしたセッションを返す"""
    session = requests.Session()
    r = session.get(BASE_MASTER + "master_login.aspx")
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
    session.post(BASE_MASTER + "master_login.aspx", data=data, allow_redirects=True)
    return session


def get_trial_schools(session):
    """体験塾の一覧を全ページ取得（index, 塾名, 塾ID）"""
    url = BASE_MASTER + "apply_list.aspx"
    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")
    search_data = {
        "__VIEWSTATE": soup.find("input", {"name": "__VIEWSTATE"})["value"],
        "__VIEWSTATEGENERATOR": soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"],
        "__EVENTVALIDATION": soup.find("input", {"name": "__EVENTVALIDATION"})["value"],
        "__EVENTTARGET": "search_list",
        "__EVENTARGUMENT": "",
        "jid": "",
        "confirmed": "confirmed3",
        "confirmed2": "confirmed2_3",
        "RadBuildAdmin": "RadBuildAdmin3",
        "SelSearchEntryType": "1",  # 体験塾
        "SelSetPrefe": "Z",
        "AgentSearchCombobox": "A",
    }
    r2 = session.post(url, data=search_data)
    soup2 = BeautifulSoup(r2.text, "html.parser")

    entries = []
    page = 1
    while True:
        rows = soup2.select("table#apply_list tr")
        for row in rows:
            m = re.search(r'goApplicationStateNew\((\d+)\)', str(row))
            if not m:
                continue
            idx = m.group(1)
            link = row.find("a", onclick=re.compile(r'goApplicationStateNew'))
            juku_name = link.get_text(strip=True) if link else "?"
            entries.append({"index": idx, "juku_name": juku_name})

        # 次ページ確認
        next_link = soup2.find("a", string=re.compile(r"次ページ"))
        if not next_link:
            break
        page += 1
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
            "confirmed": "confirmed3",
            "confirmed2": "confirmed2_3",
            "RadBuildAdmin": "RadBuildAdmin3",
            "SelSearchEntryType": "1",
            "SelSetPrefe": "Z",
            "AgentSearchCombobox": "A",
        }
        r3 = session.post(url, data=page_data)
        soup2 = BeautifulSoup(r3.text, "html.parser")

    return entries


def get_admin_credentials(session, index):
    """apply_state_new.aspxから管理者ID/PWと承認日を取得"""
    r = session.get(f"{BASE_MASTER}apply_state_new.aspx?index={index}")
    text = r.text.replace("&emsp;", " ").replace("&nbsp;", " ")

    admin_id = admin_pw = None
    m = re.search(r'【管理者ID/PW】(\S+)\s*/\s*(\S+?)[\s<]', text)
    if m:
        admin_id = m.group(1)
        admin_pw = m.group(2)

    # 最初の承認日を取得（最も古い操作日 = ID発行日）
    soup = BeautifulSoup(r.text, "html.parser")
    dates = re.findall(r'(\d{4}/\d{2}/\d{2})', text)
    earliest_date = None
    for d in dates:
        try:
            dt = datetime.strptime(d, "%Y/%m/%d")
            if earliest_date is None or dt < earliest_date:
                earliest_date = dt
        except ValueError:
            pass

    return admin_id, admin_pw, earliest_date


def check_student_activity(admin_id, admin_pw):
    """塾管理者でログインし、生徒の最終操作日時を取得"""
    try:
        session = requests.Session()
        r = session.get(BASE_JUKU, timeout=15)
        soup = BeautifulSoup(r.text, "html.parser")
        login_data = {
            "__VIEWSTATE": soup.find("input", {"name": "__VIEWSTATE"})["value"],
            "__VIEWSTATEGENERATOR": soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"],
            "__EVENTVALIDATION": soup.find("input", {"name": "__EVENTVALIDATION"})["value"],
            "__EVENTTARGET": "btnLogin",
            "__EVENTARGUMENT": "",
            "index": admin_id,
            "pass": admin_pw,
        }
        r2 = session.post(BASE_JUKU, data=login_data, allow_redirects=True, timeout=15)

        if "welcome" not in r2.url:
            return None, 0  # ログイン失敗

        r3 = session.get(
            "https://www.eduplus.jp/eduplus/myadmin/now_action_list.aspx", timeout=15
        )
        soup3 = BeautifulSoup(r3.text, "html.parser")

        latest_dt = None
        student_count = 0
        tables = soup3.find_all("table")
        for table in tables:
            for row in table.find_all("tr"):
                cells = row.find_all("td")
                if len(cells) >= 2:
                    last_action = cells[1].get_text(strip=True)
                    m = re.match(r"(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2})", last_action)
                    if m:
                        student_count += 1
                        dt = datetime.strptime(m.group(1), "%Y/%m/%d %H:%M")
                        if latest_dt is None or dt > latest_dt:
                            latest_dt = dt

        return latest_dt, student_count
    except Exception:
        return None, 0


def main():
    now = datetime.now(JST)
    now_str = now.strftime("%Y/%m/%d %H:%M")
    cutoff = now - timedelta(days=TRIAL_MONTHS * 30)  # 3ヶ月前

    print(f"[{now_str}] 体験塾チェック開始...")

    try:
        session = master_login()
        schools = get_trial_schools(session)
        print(f"体験塾: {len(schools)}件")

        suspects = []

        for i, school in enumerate(schools):
            # 除外リストチェック
            if any(ex in school["juku_name"] for ex in EXCLUDE_NAMES):
                continue

            admin_id, admin_pw, reg_date = get_admin_credentials(session, school["index"])

            if not admin_id or not admin_pw:
                continue

            # 登録日が3ヶ月以内ならスキップ
            if reg_date and reg_date > cutoff.replace(tzinfo=None):
                continue

            latest_activity, student_count = check_student_activity(admin_id, admin_pw)

            if latest_activity is None or student_count == 0:
                continue

            # 直近30日以内にアクティビティがあれば無銭飲食の疑い
            now_naive = now.replace(tzinfo=None)
            days_since_activity = (now_naive - latest_activity).days
            if days_since_activity <= 30:
                days_since_reg = (now_naive - reg_date).days if reg_date else "?"
                suspects.append({
                    "name": school["juku_name"],
                    "reg_date": reg_date.strftime("%Y/%m/%d") if reg_date else "不明",
                    "days_since_reg": days_since_reg,
                    "latest_activity": latest_activity.strftime("%Y/%m/%d %H:%M"),
                    "days_since_activity": days_since_activity,
                    "student_count": student_count,
                })

            if (i + 1) % 50 == 0:
                print(f"  {i+1}/{len(schools)} 処理済...")

        # レポート作成
        if suspects:
            lines = [f"🚨 [{now_str}] 体験塾 無銭飲食チェック\n"]
            lines.append(f"該当: {len(suspects)}件\n")
            for s in suspects:
                lines.append(
                    f"• {s['name']}\n"
                    f"  登録: {s['reg_date']} ({s['days_since_reg']}日前)\n"
                    f"  最終操作: {s['latest_activity']} ({s['days_since_activity']}日前)\n"
                    f"  生徒数: {s['student_count']}名"
                )
            msg = "\n".join(lines)
        else:
            msg = f"✅ [{now_str}] 体験塾チェック完了\n該当なし"

        send_telegram(msg)
        print(msg)

    except Exception as e:
        err_msg = f"❌ [{now_str}] 体験塾チェックエラー\n{e}"
        send_telegram(err_msg)
        print(err_msg)


if __name__ == "__main__":
    main()

"""
eduplus 塾登録 Telegram Bot（Python + GAS API版）
"""

import json
import logging
import logging.handlers
import time
import requests
from datetime import datetime
from zoneinfo import ZoneInfo
from collections import deque
import re
import os
import sys
import threading
from dotenv import load_dotenv

load_dotenv(override=True)

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

# ====================================
# ログ設定
# ====================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(SCRIPT_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logger = logging.getLogger("eduplusbot")
logger.setLevel(logging.DEBUG)

# ファイルハンドラ（日次ローテーション、30日保持）
file_handler = logging.handlers.TimedRotatingFileHandler(
    os.path.join(LOG_DIR, "bot.log"),
    when="midnight",
    backupCount=30,
    encoding="utf-8",
)
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
))

# コンソールハンドラ
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)s] %(message)s"
))

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# ====================================
# 設定
# ====================================
CONFIG = {
    'TELEGRAM_BOT_TOKEN': os.environ['TELEGRAM_BOT_TOKEN'],
    'GAS_URL': os.environ['GAS_URL'],
    'EDUPLUS_CHAT_ID': int(os.environ.get('EDUPLUS_CHAT_ID', '-5126783705')),
}

JST = ZoneInfo("Asia/Tokyo")

# 自動承認スケジュール（JST時刻）
APPROVE_SCHEDULE_HOURS = [11]

# 対話型セッション管理（chat_id → セッション情報）
SESSIONS = {}

# 処理済み update_id を保持（同一プロセス内での重複防止）
PROCESSED_UPDATES = deque(maxlen=1000)

# 処理済み message_id を保持（chat_id:message_id で重複防止）
PROCESSED_MESSAGES = deque(maxlen=1000)

# 確認する項目の順番
CONFIRM_FIELDS = [
    ('塾名', '塾名'),
    ('法人名', '法人名'),
    ('代表者名', '代表者名'),
    ('郵便番号', '郵便番号'),
    ('住所', '住所'),
    ('電話番号', '電話番号'),
    ('メールアドレス', 'メールアドレス'),
    ('HP URL', 'HP'),
    ('営業担当者', '営業担当者'),
    ('塾ID', '塾ID'),
]

# ====================================
# Telegram API
# ====================================
def send_message(chat_id, text):
    url = f"https://api.telegram.org/bot{CONFIG['TELEGRAM_BOT_TOKEN']}/sendMessage"
    requests.post(url, json={'chat_id': chat_id, 'text': text}, timeout=10)

def get_updates(offset=None):
    params = {'timeout': 30, 'limit': 10}
    if offset:
        params['offset'] = offset
    url = f"https://api.telegram.org/bot{CONFIG['TELEGRAM_BOT_TOKEN']}/getUpdates"
    resp = requests.get(url, params=params, timeout=35)
    return resp.json()

# ====================================
# GAS API
# ====================================
def gas_api(data):
    resp = requests.post(CONFIG['GAS_URL'], json=data, allow_redirects=True, timeout=30)
    return resp.json()

def write_to_spreadsheet(juku_data):
    result = gas_api({'action': 'write', 'juku_data': juku_data})
    return result.get('row', 0)

def get_row_data(row_num):
    result = gas_api({'action': 'get_row', 'row_num': row_num})
    return result.get('data', {})

def update_cell(row_num, col_num, value):
    gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': col_num, 'value': value})

def build_email_body(juku_name, admin_id, admin_pw, sample_id, sample_pw, sales_person):
    return f"""{juku_name}御中

お世話になります。
名大ＳＫＹ未来教育事業部です。

早速ですが、eduplus+の管理者ＩＤと体験用ＩＤをご連絡させていただきます。

マニュアルをご覧いただきご不明な箇所等ございましたら、
担当者までお問い合せください。


■ログインＵＲＬ

https://www.eduplus.jp/eduplus/


  【管理者ID】{admin_id}

   【パスワード】  {admin_pw}

　【サンプルID】{sample_id}

   【パスワード】{sample_pw}


下記より書類をダウンロードしてご確認ください。

◆2026年度営業日カレンダー
https://www.eduplus.website/eduplus/2026_eigyo.pdf

◆生徒用マニュアル
http://www.eduplus.website/eduplus/2021student_manual.pdf

◆講師用マニュアル
http://www.eduplus.website/eduplus/2021teacher_manual.pdf
"""


def send_email_via_gas(to_email, juku_name, admin_id, admin_pw, sample_id, sample_pw, sales_person):
    from gmail_sender import send_email
    body = build_email_body(juku_name, admin_id, admin_pw, sample_id, sample_pw, sales_person)
    subject = f'eduplus+ 管理者ID・体験用IDのご連絡【{juku_name}】'
    logger.info("メール送信開始: to=%s, 塾名=%s", to_email, juku_name)
    success, error_msg = send_email(to_email, subject, body)
    if not success:
        logger.error("メール送信失敗: to=%s, 塾名=%s, error=%s", to_email, juku_name, error_msg)
    return success, error_msg

def save_to_drive_via_gas(file_url, file_name):
    result = gas_api({
        'action': 'save_drive',
        'file_url': file_url,
        'file_name': file_name,
    })
    return result.get('success', False)

# ====================================
# テキスト解析
# ====================================
def parse_juku_info(text):
    data = {}
    field_map = {
        '塾名': '塾名', '法人名': '法人名', '代表者名': '代表者名',
        '代表者': '代表者名', '郵便番号': '郵便番号', '〒': '郵便番号',
        '住所': '住所', '電話番号': '電話番号', '電話': '電話番号',
        'TEL': '電話番号', 'tel': '電話番号',
        'メールアドレス': 'メールアドレス', 'メール': 'メールアドレス',
        'email': 'メールアドレス', 'Email': 'メールアドレス',
        'HP': 'HP URL', 'HP URL': 'HP URL', 'URL': 'HP URL',
        'ホームページ': 'HP URL',
        '営業担当者': '営業担当者', '営業担当': '営業担当者',
        '担当者': '営業担当者', '担当': '営業担当者',
        '塾ID': '塾ID', 'ID': '塾ID', 'id': '塾ID',
    }

    for line in text.split('\n'):
        match = re.match(r'^(.+?)[：:]\s*(.+)$', line.strip())
        if match:
            key = match.group(1).strip()
            value = match.group(2).strip()
            if key in field_map:
                data[field_map[key]] = value

    if 'メールアドレス' not in data:
        email_match = re.search(r'[\w.-]+@[\w.-]+\.\w+', text)
        if email_match:
            data['メールアドレス'] = email_match.group()

    if 'HP URL' not in data:
        url_match = re.search(r'https?://[\w.-]+(?:/[\w./-]*)?', text)
        if url_match:
            data['HP URL'] = url_match.group()

    return data

# ====================================
# コマンド処理
# ====================================
def handle_message(update):
    if 'message' not in update:
        return

    message = update['message']
    chat_id = message['chat']['id']
    text = message.get('text', '')
    caption = message.get('caption', '')

    # message_id による重複処理防止（複数プロセス・再送対策）
    msg_id = message.get('message_id')
    if msg_id:
        dedup_key = f"{chat_id}:{msg_id}"
        if dedup_key in PROCESSED_MESSAGES:
            logger.debug("スキップ（message_id重複）: %s", dedup_key)
            return
        PROCESSED_MESSAGES.append(dedup_key)

    # Bot自身のメッセージは無視
    sender = message.get('from', {})
    if sender.get('is_bot', False):
        return

    # メッセージ受信ログ
    has_photo = 'photo' in message
    has_text = bool(message.get('text'))
    logger.info("受信: chat=%s, from=%s, photo=%s, text=%s, text_val=%s",
                chat_id, sender.get('first_name', '?'), has_photo, has_text,
                text[:30] if text else '')

    if 'photo' in message:
        handle_photo(message, chat_id)
        return

    if 'document' in message:
        handle_document(message, chat_id)
        return

    # 対話型セッション中の場合
    if chat_id in SESSIONS and text and not text.startswith('/'):
        session = SESSIONS[chat_id]
        if session.get('type') == 'mail':
            handle_mail_session_reply(chat_id, text)
        else:
            handle_session_reply(chat_id, text)
        return

    # /cancel でセッションキャンセル
    if text.startswith('/cancel'):
        if chat_id in SESSIONS:
            del SESSIONS[chat_id]
            send_message(chat_id, "❌ 登録をキャンセルしました。")
        else:
            send_message(chat_id, "キャンセルするセッションがありません。")
        return

    if text.startswith('/register') or text.startswith('/登録'):
        handle_registration(text, chat_id)
        return

    if text.startswith('/retry_name'):
        handle_retry_name(text, chat_id)
        return

    if text.startswith('/retry'):
        handle_retry(text, chat_id)
        return

    if text.startswith('/mail'):
        handle_mail_change(text, chat_id)
        return

    if text.startswith('/approve'):
        handle_approve(text, chat_id)
        return

    if text in ['/start', '/help']:
        send_help(chat_id)
        return

    if any(kw in text for kw in ['塾', '教室', 'スクール']):
        data = parse_juku_info(text)
        if data.get('塾名'):
            handle_registration('/登録\n' + text, chat_id)
        return

def send_help(chat_id):
    help_text = """📋 eduplus 塾登録Bot

【情報の送り方】
① 名刺の写真を送る
② テキストで情報を送る

【テキスト登録フォーマット】
/登録 の後に以下の形式で送信：

/登録
塾ID: abc12
塾名: ○○塾
法人名: 株式会社○○
代表者名: 山田太郎
郵便番号: 000-0000
住所: 東京都○○区...
電話番号: 03-0000-0000
メールアドレス: info@example.com
HP: https://example.com
営業担当者: 佐藤

※塾IDは英数字で指定（推奨3-5文字、任意の長さ可）
※登録後、自動でeduplus ID発行を実行します

【コマンド一覧】
/登録 - 新規塾登録 + ID自動発行
/approve 行番号 - メール送信先を確認
/approve 行番号 yes - メール送信実行
/mail 行番号 新アドレス - 送信先を変更
/help - このヘルプを表示"""
    send_message(chat_id, help_text)

def generate_juku_id_candidates(juku_name):
    """塾名から塾ID候補を生成"""
    from pykakasi import kakasi
    import random, string

    kks = kakasi()
    result = kks.convert(juku_name)

    # 頭文字パターン
    initials = ''.join([item['hepburn'][0] for item in result if item['hepburn'] and item['hepburn'][0].isalpha()])

    candidates = set()

    # 頭文字そのまま（3文字以上なら）
    if len(initials) >= 3:
        candidates.add(initials[:5])

    # 頭文字 + 数字
    for i in range(1, 10):
        candidates.add(initials[:4] + str(i))

    # 最初の単語のローマ字（短縮）
    if result:
        first_word = result[0]['hepburn'][:5]
        if len(first_word) >= 3:
            candidates.add(first_word)
            candidates.add(first_word + '1')

    # 5つに絞る
    candidates = sorted(list(candidates))[:5]
    return candidates


def parse_ocr_text(text):
    """OCRテキストから名刺情報を抽出"""
    data = {}

    lines = text.split('\n')
    full_text = text

    # 有効な行だけ抽出（空行、装飾行を除外）
    valid_lines = []
    for line in lines:
        line = line.strip()
        if line and not re.match(r'^[_\-=\s\.·*]+$', line) and len(line) > 1:
            valid_lines.append(line)

    # 電話番号（Tel/TEL優先、FAX/mobile除外）
    tel_match = re.search(r'[Tt][Ee][Ll][：:]\s*([\d\-（）()\s/]+\d)', full_text)
    if tel_match:
        phone = tel_match.group(1)
        # FAX部分を除去
        phone = phone.split('/')[0].strip()
        # ()（）をハイフンに変換
        phone = re.sub(r'[\s]', '', phone)
        phone = phone.replace('（', '-').replace('）', '-').replace('(', '-').replace(')', '-')
        phone = re.sub(r'-+', '-', phone)  # 連続ハイフンを1つに
        phone = phone.strip('-')
        data['電話番号'] = phone
    else:
        # Telがなければ最初の電話番号っぽいもの
        phones = re.findall(r'(\d{2,4}[-\-]\d{2,4}[-\-]\d{3,4})', full_text)
        if phones:
            data['電話番号'] = phones[0]

    # メールアドレス
    email_match = re.search(r'[\w.\-]+@[\w.\-]+\.\w+', full_text)
    if email_match:
        data['メールアドレス'] = email_match.group()

    # URL
    url_match = re.search(r'https?://[\w.\-/]+', full_text)
    if url_match:
        data['HP URL'] = url_match.group()
    else:
        www_match = re.search(r'www\.[\w.\-/]+', full_text)
        if www_match:
            data['HP URL'] = 'https://' + www_match.group()

    # 郵便番号
    zip_match = re.search(r'〒?\s*(\d{3}[-\-]\d{4})', full_text)
    if zip_match:
        data['郵便番号'] = zip_match.group(1)

    # 住所（都道府県を含む行）
    for line in valid_lines:
        if re.search(r'(東京都|北海道|大阪府|京都府|.{2,3}県)', line):
            # 郵便番号を除去
            addr = re.sub(r'〒?\s*\d{3}[-\-]\d{4}\s*', '', line).strip()
            if addr:
                data['住所'] = addr
            break

    # 法人名（株式会社、有限会社などを含む行）
    for line in valid_lines:
        if re.search(r'(株式会社|有限会社|合同会社|一般社団|NPO|ＮＰＯ)', line):
            data['法人名'] = line.strip()
            break

    # 代表者名（役職の次の行、または漢字2-4文字の名前っぽい行）
    for i, line in enumerate(valid_lines):
        if re.search(r'(代表|社長|取締役|塾長|教室長|室長|本部長|部長)', line):
            # この行に名前が含まれるか
            name_in_line = re.sub(r'(代表取締役|常務取締役|取締役|営業本部長|本部長|部長|社長|塾長|教室長|室長)', '', line).strip()
            if name_in_line and len(name_in_line) >= 2:
                data['代表者名'] = name_in_line
            elif i + 1 < len(valid_lines):
                next_line = valid_lines[i + 1].strip()
                if len(next_line) >= 2 and len(next_line) <= 10 and not re.search(r'[@\d\-/]', next_line):
                    data['代表者名'] = next_line
            break

    # 塾名候補（法人名以外で、最初の方にある短い日本語行）
    for line in valid_lines:
        line = line.strip()
        # 法人名、住所、電話等でない短い行
        if (line != data.get('法人名', '') and
            line != data.get('代表者名', '') and
            not re.search(r'[@\d〒]', line) and
            not re.search(r'(株式会社|有限会社|Tel|Fax|mobile|mail)', line, re.IGNORECASE) and
            2 <= len(line) <= 20 and
            re.search(r'[\u3000-\u9fff]', line)):  # 日本語を含む
            data['塾名候補'] = line
            break

    return data


def handle_photo(message, chat_id):
    photo = message['photo'][-1]
    file_id = photo['file_id']
    token = CONFIG['TELEGRAM_BOT_TOKEN']

    file_info = requests.get(f"https://api.telegram.org/bot{token}/getFile?file_id={file_id}").json()

    if not file_info.get('ok'):
        send_message(chat_id, "❌ 画像の取得に失敗しました。")
        return

    file_path = file_info['result']['file_path']
    file_url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    file_name = f"名刺_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"

    send_message(chat_id, "⏳ 名刺を読み取り中...")

    def do_ocr():
        try:
            # Googleドライブに保存
            save_to_drive_via_gas(file_url, file_name)

            # OCR実行
            ocr_result = gas_api({'action': 'ocr', 'file_url': file_url})

            if not ocr_result.get('success'):
                send_message(chat_id, f"❌ 名刺の読み取りに失敗しました: {ocr_result.get('error', '不明')}")
                return

            ocr_text = ocr_result['text']

            # OCRテキストから情報を抽出
            extracted = parse_ocr_text(ocr_text)

            # 塾名候補からID候補を生成
            guessed_name = extracted.get('塾名候補', '')
            candidates = []
            if guessed_name:
                candidates = generate_juku_id_candidates(guessed_name)
                extracted['塾名'] = guessed_name

            send_message(chat_id, "📋 名刺の読み取りが完了しました。項目ごとに確認していきます。\n\n変更がなければ「ok」、変更する場合は正しい値を入力してください。\n「やり直し」→ 最初からリセット\n/cancel でキャンセルできます。")

            # 対話型セッションを開始
            start_confirm_session(chat_id, extracted, candidates)

        except Exception as e:
            send_message(chat_id, f"❌ 名刺処理エラー: {str(e)}")

    thread = threading.Thread(target=do_ocr)
    thread.start()

def handle_document(message, chat_id):
    doc = message['document']
    file_id = doc['file_id']
    file_name = doc.get('file_name', f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    token = CONFIG['TELEGRAM_BOT_TOKEN']

    file_info = requests.get(f"https://api.telegram.org/bot{token}/getFile?file_id={file_id}").json()

    if file_info.get('ok'):
        file_path = file_info['result']['file_path']
        file_url = f"https://api.telegram.org/file/bot{token}/{file_path}"
        success = save_to_drive_via_gas(file_url, file_name)

        if success:
            send_message(chat_id, f"✅ ファイルをGoogleドライブに保存しました。\n📁 ファイル名: {file_name}")
        else:
            send_message(chat_id, "❌ Googleドライブへの保存に失敗しました。")
    else:
        send_message(chat_id, "❌ ファイルの取得に失敗しました。")

def generate_error_with_candidates(chat_id, row_num, juku_id, juku_name, result):
    """重複エラー時に候補を提示"""
    import random, string

    error_type = result.get('error_type', 'unknown') if result else 'unknown'

    if error_type == 'name_duplicate':
        # 塾名重複 → 塾名の候補を提示
        name_candidates = [
            f"{juku_name}教室",
            f"{juku_name}本校",
            f"{juku_name}本部",
            f"個別指導{juku_name}",
            f"{juku_name}進学教室",
        ]
        candidate_list = '\n'.join([f"  {n}" for n in name_candidates])
        send_message(chat_id,
            f"❌ 塾名「{juku_name}」は既に登録されています。\n\n"
            f"塾名の候補：\n{candidate_list}\n\n"
            f"塾名を変更して /retry_name {row_num} {juku_id} [新しい塾名] で再試行してください。\n"
            f"例: /retry_name {row_num} {juku_id} {juku_name}教室"
        )
    elif error_type == 'id_duplicate':
        # 塾ID重複 → IDの候補を提示
        base = juku_id[:3]
        candidates = set()
        while len(candidates) < 5:
            suffix = ''.join(random.choice(string.digits) for _ in range(2))
            candidates.add(base + suffix)
        candidate_list = '\n'.join([f"  /retry {row_num} {c}" for c in sorted(candidates)])
        send_message(chat_id,
            f"❌ 塾ID「{juku_id}」は既に使用されています。\n\n"
            f"候補から選んでください：\n{candidate_list}\n\n"
            f"または /retry {row_num} [任意のID] で指定"
        )
    else:
        # 不明なエラー
        error_msg = result.get('error_message', '') if result else ''
        base = juku_id[:3]
        candidates = set()
        while len(candidates) < 5:
            suffix = ''.join(random.choice(string.digits) for _ in range(2))
            candidates.add(base + suffix)
        candidate_list = '\n'.join([f"  /retry {row_num} {c}" for c in sorted(candidates)])
        send_message(chat_id,
            f"❌ 登録に失敗しました。\n"
            f"エラー: {error_msg[:100]}\n\n"
            f"塾IDを変更して再試行：\n{candidate_list}"
        )


def start_confirm_session(chat_id, extracted_data, id_candidates):
    """対話型確認セッションを開始"""
    SESSIONS[chat_id] = {
        'data': extracted_data,
        'step': 0,
        'id_candidates': id_candidates,
    }
    ask_next_field(chat_id)


def ask_next_field(chat_id):
    """次の項目を質問"""
    try:
        session = SESSIONS.get(chat_id)
        if not session:
            return

        step = session['step']

        # 同じステップを重複して質問しない（多重呼び出し防止）
        if session.get('_last_asked_step') == step:
            logger.debug('重複質問スキップ: chat=%s, step=%s', chat_id, step)
            return
        session['_last_asked_step'] = step

        if step >= len(CONFIRM_FIELDS):
            # 全項目確認完了 → 登録実行
            finalize_session(chat_id)
            return

        field_key, field_label = CONFIRM_FIELDS[step]
        current_value = session['data'].get(field_key, '')

        if field_key == '塾ID':
            # 塾ID候補を表示
            candidates = session.get('id_candidates', [])
            msg = f"【{step+1}/{len(CONFIRM_FIELDS)}】塾ID\n\n"
            if candidates:
                msg += "候補:\n"
                for c in candidates:
                    msg += f"  {c}\n"
                msg += f"\n候補の中から選ぶか、任意のIDを入力してください。"
            else:
                msg += "塾IDを入力してください（英数字）"
        elif current_value:
            msg = f"【{step+1}/{len(CONFIRM_FIELDS)}】{field_label}\n\n"
            msg += f"  {current_value}\n\n"
            msg += f"OKなら「はい」、変更する場合は正しい値を入力。"
        else:
            msg = f"【{step+1}/{len(CONFIRM_FIELDS)}】{field_label}\n\n"
            msg += f"値が読み取れませんでした。入力してください。\n"
            msg += f"スキップする場合は「skip」"

        send_message(chat_id, msg)
        logger.debug("質問送信: step=%s, field=%s, value=%s", step, field_key, current_value[:30] if current_value else 'なし')
    except Exception as e:
        logger.error("ask_next_field エラー: step=%s, error=%s", step if 'step' in dir() else '?', e, exc_info=True)
        import traceback
        traceback.print_exc()


def handle_session_reply(chat_id, text):
    """セッション中のユーザー返信を処理"""
    session = SESSIONS.get(chat_id)
    if not session:
        return

    step = session['step']
    if step >= len(CONFIRM_FIELDS):
        return

    field_key, field_label = CONFIRM_FIELDS[step]
    text = text.strip()

    try:
        # やり直し → セッションを完全リセット
        if text.strip() in ['やり直し', 'やりなおし', 'リセット', 'reset', 'restart']:
            del SESSIONS[chat_id]
            send_message(chat_id, "🔄 最初からやり直します。\n名刺の写真を送るか、/登録 コマンドで情報を入力してください。")
            return

        # 全角・半角・大小文字を正規化して判定
        normalized = text.replace('Ｏ', 'O').replace('Ｋ', 'K').replace('ｏ', 'o').replace('ｋ', 'k').upper().strip()
        ok_words = ['OK', 'Y', 'YES', 'はい', 'うん', '次', 'ＯＫ', 'おｋ', 'オッケー', 'おけ']
        if normalized in [w.upper() for w in ok_words] or text.strip() in ok_words:
            # 現在の値を採用
            pass
        elif text.lower() == 'skip' or text == 'スキップ':
            session['data'][field_key] = ''
        else:
            # ユーザーが新しい値を入力
            session['data'][field_key] = text

            # 塾名が変わったらID候補を再生成
            if field_key == '塾名':
                session['id_candidates'] = generate_juku_id_candidates(text)

        session['step'] += 1
        ask_next_field(chat_id)
    except Exception as e:
        logger.error("セッション返信エラー: step=%s, field=%s, error=%s", step, field_key, e, exc_info=True)
        send_message(chat_id, f"❌ エラーが発生しました: {str(e)}")


def finalize_session(chat_id):
    """全項目確認完了 → 登録・ID発行実行"""
    session = SESSIONS.get(chat_id)
    if not session:
        return

    data = session['data']
    del SESSIONS[chat_id]

    juku_id = data.get('塾ID', '')
    juku_name = data.get('塾名', '')

    if not juku_name:
        send_message(chat_id, "❌ 塾名が未入力です。最初からやり直してください。")
        return

    if not juku_id:
        send_message(chat_id, "❌ 塾IDが未入力です。最初からやり直してください。")
        return

    # 確認メッセージ
    msg = "📋 登録内容の最終確認:\n\n"
    for field_key, field_label in CONFIRM_FIELDS:
        val = data.get(field_key, '未入力')
        msg += f"{field_label}: {val}\n"
    msg += f"\n⏳ 登録・ID発行を開始します。1分ほどお待ちください..."
    send_message(chat_id, msg)

    # スプレッドシートに書き込み
    try:
        row = write_to_spreadsheet(data)
        logger.info("スプレッドシート書き込み完了: row=%s, 塾名=%s", row, juku_name)
    except Exception as e:
        logger.error("スプレッドシート書き込み失敗: 塾名=%s, error=%s", juku_name, e, exc_info=True)
        send_message(chat_id, f"❌ スプレッドシート書き込み失敗: {str(e)}")
        return

    # バックグラウンドでeduplus登録
    def do_issue():
        try:
            from eduplus import register_juku
            result = register_juku(juku_id, juku_name)

            if result and not result.get('error'):
                gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 11, 'value': result['admin_id']})
                gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 12, 'value': result['admin_pw']})
                gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 13, 'value': result['sample_id']})
                gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 14, 'value': result['sample_pw']})
                gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 15, 'value': 'ID発行済'})

                issue_msg = (
                    f"✅ {juku_name}のID発行が完了しました！\n\n"
                    f"管理者ID: {result['admin_id']}\n"
                    f"管理者PW: {result['admin_pw']}\n"
                    f"サンプルID: {result['sample_id']}\n"
                    f"サンプルPW: {result['sample_pw']}\n\n"
                    f"📧 メール送信:\n"
                    f"「確認」→ 送信内容を確認\n"
                    f"「送信」→ メール送信実行\n"
                    f"「変更 新アドレス」→ 送信先変更\n"
                    f"「削除して再登録」→ 誤ったIDを削除して再登録\n"
                    f"「キャンセル」→ 送信しない"
                )
                send_message(chat_id, issue_msg)

                # メール送信セッションを開始
                SESSIONS[chat_id] = {
                    'type': 'mail',
                    'row': row,
                    'juku_name': juku_name,
                    'juku_id': juku_id,
                }
            else:
                generate_error_with_candidates(chat_id, row, juku_id, juku_name, result)
        except Exception as e:
            logger.error("ID発行エラー: 塾ID=%s, 塾名=%s, error=%s", juku_id, juku_name, e, exc_info=True)
            send_message(chat_id, f"❌ ID発行エラー: {str(e)}")

    thread = threading.Thread(target=do_issue)
    thread.start()


def handle_mail_session_reply(chat_id, text):
    """メール送信セッション中の返信を処理"""
    session = SESSIONS.get(chat_id)
    if not session:
        return

    row_num = session['row']
    juku_name = session['juku_name']
    normalized = text.strip()

    # 「確認」系
    if normalized in ['確認', '送信先確認', '内容確認', '確認する']:
        row = get_row_data(row_num)
        email = row.get('送信先メール（変更時）') or row.get('メールアドレス')
        admin_id = row.get('管理者ID')
        sample_id = row.get('サンプルID')
        sales_person = row.get('営業担当者')
        msg = (
            f"📧 メール送信内容\n\n"
            f"送信先: {email}\n"
            f"塾名: {juku_name}\n"
            f"管理者ID: {admin_id}\n"
            f"サンプルID: {sample_id}\n"
            f"営業担当者: {sales_person}\n\n"
            f"「送信」→ 送信実行\n"
            f"「変更 新アドレス」→ 送信先変更\n"
            f"「キャンセル」→ 送信しない"
        )
        send_message(chat_id, msg)
        return

    # 「送信」系
    if normalized in ['送信', '送信実行', '送る', '送って', 'はい', 'yes']:
        row = get_row_data(row_num)
        email = row.get('送信先メール（変更時）') or row.get('メールアドレス')
        admin_id = row.get('管理者ID')
        admin_pw = row.get('パスワード')
        sample_id = row.get('サンプルID')
        sample_pw = row.get('サンプルパスワード')
        sales_person = row.get('営業担当者')

        if not email:
            send_message(chat_id, "⚠️ メールアドレスが登録されていません。")
            return
        if not admin_id:
            send_message(chat_id, "⚠️ 管理者IDが未設定です。")
            return

        success, error_msg = send_email_via_gas(email, juku_name, admin_id, admin_pw, sample_id, sample_pw, sales_person)
        if success:
            update_cell(row_num, 15, '送信済')
            send_message(chat_id, f"✅ {juku_name}（{email}）へメールを送信しました。")
        else:
            send_message(chat_id, f"❌ メール送信に失敗しました。\n原因: {error_msg}")

        del SESSIONS[chat_id]
        return

    # 「変更」系
    if normalized.startswith('変更') or normalized.startswith('アドレス変更'):
        parts = normalized.split()
        if len(parts) >= 2:
            new_email = parts[-1]
            if '@' in new_email and '.' in new_email:
                update_cell(row_num, 16, new_email)
                send_message(chat_id,
                    f"✅ 送信先を {new_email} に変更しました。\n\n"
                    f"「送信」→ メール送信実行\n"
                    f"「確認」→ 内容確認"
                )
                return
        send_message(chat_id, "⚠️ 「変更 新しいメールアドレス」の形式で入力してください。\n例: 変更 new@example.com")
        return

    # 「削除して再登録」系
    if normalized in ['削除して再登録', '削除', '再登録', 'やり直し', 'やりなおし', 'リセット']:
        row_num = session['row']
        juku_name = session['juku_name']
        row = get_row_data(row_num)
        juku_id = row.get('塾ID', '')
        if not juku_id:
            send_message(chat_id, "⚠️ 塾IDが取得できません。スプレッドシートを確認してください。")
            return
        send_message(chat_id,
            f"⚠️ 本当に「{juku_name}」（塾ID: {juku_id}）を削除して再登録しますか？\n\n"
            f"「削除実行」→ 削除して新しいIDを入力\n"
            f"「キャンセル」→ 戻る"
        )
        session['pending_delete'] = True
        session['juku_id'] = juku_id
        return

    # 「削除実行」確認後
    if normalized in ['削除実行'] and session.get('pending_delete'):
        juku_id = session.get('juku_id', '')
        juku_name = session['juku_name']
        row_num = session['row']
        del SESSIONS[chat_id]
        send_message(chat_id, f"⏳ 「{juku_name}」（塾ID: {juku_id}）を削除しています...")

        def do_delete_and_reset():
            try:
                from eduplus import delete_juku
                ok = delete_juku(juku_id)
                if ok:
                    send_message(chat_id,
                        f"✅ 削除しました。\n\n"
                        f"新しい塾IDで再登録するには:\n"
                        f"/retry {row_num} 新しい塾ID\n\n"
                        f"例: /retry {row_num} abc12"
                    )
                else:
                    send_message(chat_id, f"❌ 削除に失敗しました。手動で確認してください。\n塾ID: {juku_id}")
            except Exception as e:
                send_message(chat_id, f"❌ 削除エラー: {str(e)}")

        threading.Thread(target=do_delete_and_reset).start()
        return

    # 「キャンセル」系
    if normalized in ['キャンセル', 'やめる', '送信しない', 'cancel']:
        session.pop('pending_delete', None)
        del SESSIONS[chat_id]
        send_message(chat_id, "📧 メール送信をキャンセルしました。\n後から送信する場合は /approve コマンドを使ってください。")
        return

    # 認識できない入力
    send_message(chat_id,
        f"以下のいずれかで返信してください：\n"
        f"「確認」→ 送信内容を確認\n"
        f"「送信」→ メール送信実行\n"
        f"「変更 新アドレス」→ 送信先変更\n"
        f"「削除して再登録」→ 誤ったIDを削除して再登録\n"
        f"「キャンセル」→ 送信しない"
    )


def handle_registration(text, chat_id):
    data = parse_juku_info(text)

    if not data.get('塾名'):
        send_message(chat_id, "⚠️ 塾名が見つかりません。/help でフォーマットを確認してください。")
        return

    if not data.get('塾ID'):
        send_message(chat_id, "⚠️ 塾IDが見つかりません。「塾ID: xxx」（3-5文字）を追加してください。")
        return

    juku_id = data['塾ID']

    try:
        row = write_to_spreadsheet(data)
        msg = (
            f"✅ 塾情報を登録しました（行: {row}）\n\n"
            f"塾名: {data.get('塾名', '未入力')}\n"
            f"法人名: {data.get('法人名', '未入力')}\n"
            f"代表者名: {data.get('代表者名', '未入力')}\n"
            f"メールアドレス: {data.get('メールアドレス', '未入力')}\n"
            f"営業担当者: {data.get('営業担当者', '未入力')}\n\n"
            f"⏳ eduplus ID発行を開始します（塾ID: {juku_id}）..."
        )
        send_message(chat_id, msg)

        # バックグラウンドでeduplus登録
        def do_issue():
            try:
                from eduplus import register_juku
                result = register_juku(juku_id, data['塾名'])

                if result and not result.get('error'):
                    # スプレッドシートにID・パスワードを書き込み
                    gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 11, 'value': result['admin_id']})
                    gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 12, 'value': result['admin_pw']})
                    gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 13, 'value': result['sample_id']})
                    gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 14, 'value': result['sample_pw']})
                    gas_api({'action': 'update_cell', 'row_num': row, 'col_num': 15, 'value': 'ID発行済'})

                    issue_msg = (
                        f"✅ {data['塾名']}のID発行が完了しました！\n\n"
                        f"管理者ID: {result['admin_id']}\n"
                        f"管理者PW: {result['admin_pw']}\n"
                        f"サンプルID: {result['sample_id']}\n"
                        f"サンプルPW: {result['sample_pw']}\n\n"
                        f"📧 メール送信:\n"
                        f"/approve {row} → 送信先確認\n"
                        f"/approve {row} yes → 送信実行\n"
                        f"/mail {row} 別アドレス → 送信先変更"
                    )
                    send_message(chat_id, issue_msg)
                else:
                    generate_error_with_candidates(chat_id, row, juku_id, data['塾名'], result)
            except Exception as e:
                logger.error("ID発行エラー(写真): 塾ID=%s, error=%s", juku_id, e, exc_info=True)
                send_message(chat_id, f"❌ ID発行エラー: {str(e)}")

        thread = threading.Thread(target=do_issue)
        thread.start()

    except Exception as e:
        send_message(chat_id, f"❌ 登録に失敗しました: {str(e)}")

def handle_approve(text, chat_id):
    parts = text.strip().split()
    if len(parts) < 2:
        send_message(chat_id, "⚠️ 行番号を指定してください。\n例: /approve 2")
        return

    try:
        row_num = int(parts[1])
    except ValueError:
        send_message(chat_id, "⚠️ 正しい行番号を入力してください。")
        return

    row = get_row_data(row_num)

    email = row.get('送信先メール（変更時）') or row.get('メールアドレス')
    admin_id = row.get('管理者ID')
    juku_name = row.get('塾名')
    admin_pw = row.get('パスワード')
    sample_id = row.get('サンプルID')
    sample_pw = row.get('サンプルパスワード')
    sales_person = row.get('営業担当者')

    if not email:
        send_message(chat_id, "⚠️ メールアドレスが登録されていません。")
        return

    if not admin_id:
        send_message(chat_id, "⚠️ 管理者IDが未入力です。スプレッドシートにID情報を入力してから承認してください。")
        return

    if len(parts) > 2 and parts[2] == 'yes':
        success, error_msg = send_email_via_gas(email, juku_name, admin_id, admin_pw, sample_id, sample_pw, sales_person)
        if success:
            update_cell(row_num, 15, '送信済')
            send_message(chat_id, f"✅ {juku_name}（{email}）へメールを送信しました。")
        else:
            send_message(chat_id, f"❌ メール送信に失敗しました。\n原因: {error_msg}")
    else:
        msg = (
            f"📧 メール送信確認\n\n"
            f"送信先: {email}\n"
            f"塾名: {juku_name}\n"
            f"管理者ID: {admin_id}\n"
            f"サンプルID: {sample_id}\n"
            f"営業担当者: {sales_person}\n\n"
            f"この内容で送信しますか？\n"
            f"/approve {row_num} yes → 送信実行\n"
            f"/mail {row_num} 別アドレス → 送信先変更"
        )
        send_message(chat_id, msg)

def handle_issue(text, chat_id):
    """ID発行: /issue 行番号 塾ID"""
    parts = text.strip().split()
    if len(parts) < 3:
        send_message(chat_id, "⚠️ 形式: /issue [行番号] [塾ID(3-5文字)]\n例: /issue 2 abc12")
        return

    try:
        row_num = int(parts[1])
    except ValueError:
        send_message(chat_id, "⚠️ 正しい行番号を入力してください。")
        return

    juku_id = parts[2]
    if not juku_id:
        send_message(chat_id, "⚠️ 塾IDを指定してください。")
        return

    row = get_row_data(row_num)
    juku_name = row.get('塾名', '')
    if not juku_name:
        send_message(chat_id, "⚠️ 指定した行にデータがありません。")
        return

    send_message(chat_id, f"⏳ {juku_name}（ID: {juku_id}）のeduplus登録を開始します...")

    def do_issue():
        try:
            from eduplus import register_juku
            result = register_juku(juku_id, juku_name)

            if result:
                # スプレッドシートにID・パスワードを書き込み
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 11, 'value': result['admin_id']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 12, 'value': result['admin_pw']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 13, 'value': result['sample_id']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 14, 'value': result['sample_pw']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 15, 'value': 'ID発行済'})

                msg = (
                    f"✅ {juku_name}のID発行が完了しました！\n\n"
                    f"管理者ID: {result['admin_id']}\n"
                    f"管理者PW: {result['admin_pw']}\n"
                    f"サンプルID: {result['sample_id']}\n"
                    f"サンプルPW: {result['sample_pw']}\n\n"
                    f"メール送信:\n"
                    f"/approve {row_num} → 送信先確認\n"
                    f"/mail {row_num} 別アドレス → 送信先変更"
                )
                send_message(chat_id, msg)
            else:
                send_message(chat_id, f"❌ {juku_name}のeduplus登録に失敗しました。塾IDが重複しているか確認してください。")
        except Exception as e:
            logger.error("ID発行エラー(issue): error=%s", e, exc_info=True)
            send_message(chat_id, f"❌ ID発行エラー: {str(e)}")

    # バックグラウンドで実行（ブラウザ操作に時間がかかるため）
    thread = threading.Thread(target=do_issue)
    thread.start()


def handle_retry_name(text, chat_id):
    """/retry_name 行番号 塾ID 新しい塾名 で塾名変更して再試行"""
    parts = text.strip().split(None, 3)
    if len(parts) < 4:
        send_message(chat_id, "⚠️ 形式: /retry_name [行番号] [塾ID] [新しい塾名]")
        return

    try:
        row_num = int(parts[1])
    except ValueError:
        send_message(chat_id, "⚠️ 正しい行番号を入力してください。")
        return

    juku_id = parts[2]
    new_name = parts[3]

    send_message(chat_id, f"⏳ 塾名「{new_name}」（ID: {juku_id}）でID発行を再試行します...")

    def do_retry_name():
        try:
            from eduplus import register_juku
            result = register_juku(juku_id, new_name)

            if result and not result.get('error'):
                # 塾名もスプレッドシートで更新
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 2, 'value': new_name})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 11, 'value': result['admin_id']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 12, 'value': result['admin_pw']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 13, 'value': result['sample_id']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 14, 'value': result['sample_pw']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 15, 'value': 'ID発行済'})

                issue_msg = (
                    f"✅ {new_name}のID発行が完了しました！\n\n"
                    f"管理者ID: {result['admin_id']}\n"
                    f"管理者PW: {result['admin_pw']}\n"
                    f"サンプルID: {result['sample_id']}\n"
                    f"サンプルPW: {result['sample_pw']}\n\n"
                    f"📧 メール送信:\n"
                    f"/approve {row_num} → 送信先確認\n"
                    f"/approve {row_num} yes → 送信実行\n"
                    f"/mail {row_num} 別アドレス → 送信先変更"
                )
                send_message(chat_id, issue_msg)
            else:
                generate_error_with_candidates(chat_id, row_num, juku_id, new_name, result)
        except Exception as e:
            logger.error("ID発行エラー(retry_name): error=%s", e, exc_info=True)
            send_message(chat_id, f"❌ ID発行エラー: {str(e)}")

    thread = threading.Thread(target=do_retry_name)
    thread.start()


def handle_retry(text, chat_id):
    """/retry 行番号 新塾ID でID発行を再試行"""
    parts = text.strip().split()
    if len(parts) < 3:
        send_message(chat_id, "⚠️ 形式: /retry [行番号] [新しい塾ID]\n例: /retry 2 abc99")
        return

    try:
        row_num = int(parts[1])
    except ValueError:
        send_message(chat_id, "⚠️ 正しい行番号を入力してください。")
        return

    juku_id = parts[2]
    if not juku_id:
        send_message(chat_id, "⚠️ 塾IDを指定してください。")
        return

    row = get_row_data(row_num)
    juku_name = row.get('塾名', '')
    if not juku_name:
        send_message(chat_id, "⚠️ 指定した行にデータがありません。")
        return

    send_message(chat_id, f"⏳ {juku_name}（ID: {juku_id}）でID発行を再試行します...")

    def do_retry():
        try:
            from eduplus import register_juku
            result = register_juku(juku_id, juku_name)

            if result and not result.get('error'):
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 11, 'value': result['admin_id']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 12, 'value': result['admin_pw']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 13, 'value': result['sample_id']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 14, 'value': result['sample_pw']})
                gas_api({'action': 'update_cell', 'row_num': row_num, 'col_num': 15, 'value': 'ID発行済'})

                issue_msg = (
                    f"✅ {juku_name}のID発行が完了しました！\n\n"
                    f"管理者ID: {result['admin_id']}\n"
                    f"管理者PW: {result['admin_pw']}\n"
                    f"サンプルID: {result['sample_id']}\n"
                    f"サンプルPW: {result['sample_pw']}\n\n"
                    f"📧 メール送信:\n"
                    f"/approve {row_num} → 送信先確認\n"
                    f"/approve {row_num} yes → 送信実行\n"
                    f"/mail {row_num} 別アドレス → 送信先変更"
                )
                send_message(chat_id, issue_msg)
            else:
                generate_error_with_candidates(chat_id, row_num, juku_id, juku_name, result)
        except Exception as e:
            logger.error("ID発行エラー(retry): error=%s", e, exc_info=True)
            send_message(chat_id, f"❌ ID発行エラー: {str(e)}")

    thread = threading.Thread(target=do_retry)
    thread.start()


def handle_mail_change(text, chat_id):
    parts = text.strip().split()
    if len(parts) < 3:
        send_message(chat_id, "⚠️ 形式: /mail [行番号] [新しいメールアドレス]\n例: /mail 2 new@example.com")
        return

    try:
        row_num = int(parts[1])
    except ValueError:
        send_message(chat_id, "⚠️ 正しい行番号を入力してください。")
        return

    new_email = parts[2]
    if '@' not in new_email or '.' not in new_email:
        send_message(chat_id, "⚠️ 正しいメールアドレスを入力してください。")
        return

    update_cell(row_num, 16, new_email)
    row = get_row_data(row_num)
    juku_name = row.get('塾名', '')
    send_message(chat_id, f"✅ {juku_name}の送信先を変更しました。\n新しい送信先: {new_email}\n\n/approve {row_num} → メール送信を承認")

# ====================================
# 多重起動防止
# ====================================
def acquire_lock():
    """ファイルロックで多重起動を防止。ロック取得できなければ終了。"""
    lock_path = os.path.join(SCRIPT_DIR, ".bot.lock")
    lock_file = open(lock_path, 'w')
    try:
        if sys.platform == 'win32':
            import msvcrt
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except (OSError, IOError):
        logger.error("別のbot.pyプロセスが既に実行中です。終了します。")
        sys.exit(1)
    # lock_file を返して参照を保持（GCでクローズされないように）
    return lock_file

# ====================================
# 自動承認スケジューラ
# ====================================
def scheduled_approve():
    """毎日指定時刻にeduplus一括承認を実行するバックグラウンドスレッド"""
    from eduplus_approve import run_approve

    executed_today = set()

    while True:
        try:
            now = datetime.now(JST)
            today = now.strftime("%Y-%m-%d")
            hour = now.hour

            if hour in APPROVE_SCHEDULE_HOURS and today not in executed_today:
                logger.info("自動承認開始（スケジュール JST %s:00）", hour)
                msg = run_approve()
                chat_id = CONFIG['EDUPLUS_CHAT_ID']
                send_message(chat_id, msg)
                logger.info("自動承認完了: %s", msg[:80])
                executed_today.add(today)

                # 古い日付を削除（メモリリーク防止）
                stale = [d for d in executed_today if d < today]
                for d in stale:
                    executed_today.discard(d)

        except Exception as e:
            logger.error("自動承認エラー: %s", e, exc_info=True)

        time.sleep(30)


# ====================================
# メインループ
# ====================================
def main():
    lock_file = acquire_lock()
    logger.info("eduplus塾登録Bot 起動...")

    token = CONFIG['TELEGRAM_BOT_TOKEN']
    requests.get(f"https://api.telegram.org/bot{token}/deleteWebhook", params={"drop_pending_updates": True})
    logger.info("Webhook削除完了")

    # 自動承認スケジューラをバックグラウンドで起動
    approve_thread = threading.Thread(target=scheduled_approve, daemon=True)
    approve_thread.start()
    logger.info("自動承認スケジューラ起動（毎日 JST %s時）", APPROVE_SCHEDULE_HOURS)

    offset = None
    consecutive_errors = 0
    MAX_CONSECUTIVE_ERRORS = 20
    logger.info("ポーリング開始 token=%s...", CONFIG['TELEGRAM_BOT_TOKEN'][:15])

    while True:
        try:
            updates = get_updates(offset)

            if updates.get('ok') and updates['result']:
                for update in updates['result']:
                    update_id = update['update_id']

                    # 重複処理を防止
                    if update_id in PROCESSED_UPDATES:
                        logger.debug("スキップ（処理済み）: update_id=%s", update_id)
                        offset = update_id + 1
                        continue

                    PROCESSED_UPDATES.append(update_id)

                    try:
                        handle_message(update)
                    except Exception as e:
                        logger.error("メッセージ処理エラー: %s", e, exc_info=True)
                    offset = update_id + 1
            consecutive_errors = 0

        except requests.exceptions.Timeout:
            continue
        except Exception as e:
            consecutive_errors += 1
            logger.error(
                "ポーリングエラー (%s/%s): %s",
                consecutive_errors, MAX_CONSECUTIVE_ERRORS, e, exc_info=True,
            )
            if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                logger.critical("連続エラー上限到達。systemd に再起動を委ねます。")
                sys.exit(1)
            time.sleep(5)

if __name__ == '__main__':
    main()

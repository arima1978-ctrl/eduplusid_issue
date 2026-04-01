"""
Gmail API経由でメール送信するモジュール
"""
import logging
import os
import base64
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
SENDER_EMAIL = os.environ.get("GMAIL_SENDER", "名大ＳＫＹ未来教育事業部 <eduplus@meidaisky.jp>")


def _get_gmail_service():
    if not os.path.exists(TOKEN_PATH):
        raise FileNotFoundError(f"token.json が見つかりません: {TOKEN_PATH}")
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if creds.expired and creds.refresh_token:
        logger.info("OAuthトークンを更新中...")
        creds.refresh(Request())
        with open(TOKEN_PATH, "w") as f:
            f.write(creds.to_json())
        logger.info("OAuthトークン更新完了")
    elif creds.expired:
        raise RuntimeError("OAuthトークンが期限切れで、refresh_tokenがありません。再認証が必要です。")
    return build("gmail", "v1", credentials=creds)


def send_email(to_email: str, subject: str, body: str) -> tuple[bool, str]:
    """Gmail APIでメールを送信

    Returns:
        (成功フラグ, エラーメッセージ or 空文字)
    """
    try:
        service = _get_gmail_service()
        message = MIMEText(body, "plain", "utf-8")
        message["from"] = SENDER_EMAIL
        message["to"] = to_email
        message["subject"] = subject
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        service.users().messages().send(
            userId="me", body={"raw": raw}
        ).execute()
        logger.info("メール送信成功: to=%s, subject=%s", to_email, subject)
        return True, ""
    except FileNotFoundError as e:
        error_msg = f"認証ファイルエラー: {e}"
        logger.error(error_msg)
        return False, error_msg
    except RuntimeError as e:
        error_msg = str(e)
        logger.error("Gmail認証エラー: %s", error_msg)
        return False, error_msg
    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        logger.error("Gmail送信エラー: %s", error_msg)
        return False, error_msg

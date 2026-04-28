"""
SMTP (Gmail App Password) 経由でメール送信するモジュール

SMTP ログインは GMAIL_SMTP_USER (App Password の所有アカウント) で行い、
From ヘッダは GMAIL_FROM で指定する。GMAIL_FROM が SMTP_USER と異なる場合、
そのアドレスは Gmail の「他のアドレスから送信」で承認済みである必要がある。
"""
import logging
import os
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr

logger = logging.getLogger(__name__)

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465
SMTP_USER = os.environ.get("GMAIL_SMTP_USER", "eduplus@meidaisky.jp")
SMTP_PASSWORD = (os.environ.get("GMAIL_APP_PASSWORD") or "").replace(" ", "")
FROM_EMAIL = os.environ.get("GMAIL_FROM", SMTP_USER)
SENDER_NAME = os.environ.get("GMAIL_SENDER_NAME", "名大ＳＫＹ未来教育事業部")


def send_email(to_email: str, subject: str, body: str) -> tuple[bool, str]:
    """Gmail SMTP (App Password) でメールを送信

    Returns:
        (成功フラグ, エラーメッセージ or 空文字)
    """
    if not SMTP_PASSWORD:
        error_msg = "GMAIL_APP_PASSWORD が .env に設定されていません"
        logger.error(error_msg)
        return False, error_msg

    try:
        message = MIMEText(body, "plain", "utf-8")
        message["From"] = formataddr((SENDER_NAME, FROM_EMAIL))
        message["To"] = to_email
        message["Subject"] = subject

        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(message, from_addr=FROM_EMAIL, to_addrs=[to_email])

        logger.info("メール送信成功: smtp_user=%s, from=%s, to=%s, subject=%s",
                    SMTP_USER, FROM_EMAIL, to_email, subject)
        return True, ""
    except smtplib.SMTPAuthenticationError as e:
        error_msg = f"SMTP認証エラー (App Password要確認): {e}"
        logger.error(error_msg)
        return False, error_msg
    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        logger.error("SMTP送信エラー: %s", error_msg)
        return False, error_msg

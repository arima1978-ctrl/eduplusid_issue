"""
Gmail API OAuth 初期認証 / 再認証スクリプト

使い方:
    1. credentials.json を本スクリプトと同じディレクトリに配置
    2. python auth_setup.py を実行
    3. ブラウザが起動 → Googleアカウントでログイン＆許可
    4. token.json が生成される

SSH越しでブラウザが開けない場合:
    --console オプションで手動コピペ方式になります
        python auth_setup.py --console
"""
import argparse
import os
import sys

from google_auth_oauthlib.flow import InstalledAppFlow

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")
CREDENTIALS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


def main() -> int:
    parser = argparse.ArgumentParser(description="Gmail API OAuth 認証")
    parser.add_argument(
        "--console",
        action="store_true",
        help="ブラウザを開かず、URL手動コピペ方式で認証する（SSHなどCLI環境向け）",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=0,
        help="run_local_server で使うローカルポート（0=自動）",
    )
    args = parser.parse_args()

    if not os.path.exists(CREDENTIALS_PATH):
        print(f"[ERROR] credentials.json が見つかりません: {CREDENTIALS_PATH}", file=sys.stderr)
        print("GCPコンソールでOAuthクライアント(デスクトップ)のJSONをダウンロードして配置してください。", file=sys.stderr)
        return 1

    if os.path.exists(TOKEN_PATH):
        ans = input(f"既存の token.json を上書きしますか？ [y/N]: ").strip().lower()
        if ans != "y":
            print("中断しました。")
            return 0

    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)

    if args.console:
        # ブラウザが開けない環境向け: URLを表示 → ブラウザで認可 → コードを貼り付け
        flow.redirect_uri = "urn:ietf:wg:oauth:2.0:oob"
        auth_url, _ = flow.authorization_url(
            access_type="offline",
            prompt="consent",
            include_granted_scopes="true",
        )
        print("\n以下のURLをブラウザで開いて認可してください:")
        print(auth_url)
        code = input("\n表示された認可コードを貼り付けてください: ").strip()
        flow.fetch_token(code=code)
        creds = flow.credentials
    else:
        # ローカルPCで実行: ブラウザ自動起動
        creds = flow.run_local_server(
            port=args.port,
            access_type="offline",
            prompt="consent",
            open_browser=True,
        )

    with open(TOKEN_PATH, "w", encoding="utf-8") as f:
        f.write(creds.to_json())
    # token.json は秘匿情報なので権限を絞る（Unix系のみ有効、Windowsは無視される）
    try:
        os.chmod(TOKEN_PATH, 0o600)
    except OSError:
        pass

    print(f"\n[OK] token.json を生成しました: {TOKEN_PATH}")
    if not creds.refresh_token:
        print("[WARN] refresh_token が含まれていません。access_type=offline と prompt=consent を確認してください。")
    return 0


if __name__ == "__main__":
    sys.exit(main())

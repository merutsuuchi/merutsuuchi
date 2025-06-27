import os
import json
import imaplib
import email
import time
from email.header import decode_header
import requests
from linebot import LineBotApi
from linebot.models import TextSendMessage

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)

# OAuth2用のクライアントID・クライアントシークレットはここに固定で書くか、
# 環境変数や別ファイルから読み込む形にしてください。

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users.json")


NOTIFY_LIMIT = 30
COUNT_FILE = "notify_counts.json"

def is_user_ready(user):
    required_keys = ["LINE_USER_ID", "EMAIL_ADDRESS", "access_token", "refresh_token"]
    for key in required_keys:
        if not user.get(key):
            print(f"[{user.get('LINE_USER_ID', '不明')}] ⚠️ 必須キー {key} が未設定のためスキップ")
            return False
    return True

def generate_oauth2_string(email_address, access_token):
    return f"user={email_address}\1auth=Bearer {access_token}\1\1".encode()

def refresh_access_token(refresh_token):
    token_url = "https://oauth2.googleapis.com/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
    }
    res = requests.post(token_url, data=data)
    if res.status_code == 200:
        return res.json()["access_token"]
    else:
        print("トークン更新失敗:", res.text)
        return None
    
def decode_mime_words(s):
    if not s:
        return ""
    decoded_fragments = decode_header(s)
    decoded_string = ""
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            decoded_string += fragment.decode(encoding or 'utf-8', errors='ignore')
        else:
            decoded_string += fragment
    return decoded_string

def load_notify_counts():
    try:
        with open(COUNT_FILE, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        print("⚠️ 通知回数ファイルが存在しません。初期化します。")
        return {}

def save_notify_counts(counts):
    with open(COUNT_FILE, "w") as f:
        json.dump(counts, f, indent=2, ensure_ascii=False)

def check_email(user, users, counts):
    if not is_user_ready(user):
        return

    line_user_id = user["LINE_USER_ID"]
    if counts.get(line_user_id, 0) >= NOTIFY_LIMIT:
        print(f"[{line_user_id}] ⚠️ 通知上限（{NOTIFY_LIMIT}回）に達しています。通知をスキップします。")
        return

    email_address = user["EMAIL_ADDRESS"]
    access_token = user["access_token"]
    refresh_token = user["refresh_token"]
    imap_server = user.get("IMAP_SERVER", "imap.gmail.com")
    imap_port = user.get("IMAP_PORT", 993)

    try:
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        mail.authenticate("XOAUTH2", lambda x: generate_oauth2_string(email_address, access_token))
        mail.select("inbox")
    except imaplib.IMAP4.error:
        print(f"[{line_user_id}] IMAP認証失敗。リフレッシュを試みます")
        new_token = refresh_access_token(refresh_token)
        if new_token:
            user["access_token"] = new_token
            save_users(users)
            try:
                mail = imaplib.IMAP4_SSL(imap_server, imap_port)
                mail.authenticate("XOAUTH2", lambda x: generate_oauth2_string(email_address, new_token))
                mail.select("inbox")
            except imaplib.IMAP4.error as e2:
                print(f"[{line_user_id}] リフレッシュ後も認証失敗:", e2)
                return
        else:
            return

    status, messages = mail.search(None, "UNSEEN")
    if status != "OK":
        print(f"[{line_user_id}] メール検索失敗")
        mail.logout()
        return

    email_ids = messages[0].split()
    if not email_ids:
        print(f"[{line_user_id}] 未読メールなし")
        mail.logout()
        return

    subjects = []
    for num in email_ids[:5]:  # 最大5件
        status, data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            continue
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject = decode_mime_words(msg["Subject"])
                raw_from = msg.get("From", "不明")
                from_ = decode_mime_words(raw_from)  # ←ここを追加してデコードする
                subjects.append(f"{from_} / {subject}")

    others = len(email_ids) - 5
    subject_text = "\n".join(subjects)
    if others > 0:
        subject_text += f"\n他 {others} 件の未読メールあり"

    for num in email_ids:
        mail.store(num, '+FLAGS', '\\Seen')
    print("✅ 未読メールを既読にしました。")
        

    notify_count = counts.get(line_user_id, 0) + 1
    counts[line_user_id] = notify_count
    save_notify_counts(counts)

    tail_text = (
        f"\n-----\n"
        f"【通知回数】{notify_count}/{NOTIFY_LIMIT}回\n"
        "通知の継続をご希望の場合は、LINEで「メル通知」までご連絡ください。\n"
        "将来的にはプレミアムプランの導入も検討中です。\n\n"
        "▼ご支援はこちら（PayPay）\n"
        "https://qr.paypay.ne.jp/p2p01_NiHbdLbDfyqQRRa0"
    )

    message = f"📩 新着メール一覧:\n\n{subject_text}{tail_text}"
    line_bot_api.push_message(line_user_id, TextSendMessage(text=message))
    print(f"[{line_user_id}] メール通知＋PayPay支援文送信完了")

    mail.logout()


def save_users(users):
    with open("users.json", "w") as f:
        json.dump(users, f, indent=2, ensure_ascii=False)

def main():
    try:
        with open("users.json", "r") as f:
            users = json.load(f)
    except FileNotFoundError:
        print("users.json が見つかりません")
        return

    counts = load_notify_counts()
    for user in users:
        check_email(user, users, counts)

if __name__ == "__main__":
    while True:
        print("メールチェック開始")
        main()
        time.sleep(90)  # 10分
import os
import json
import uuid
import datetime
import requests
import imaplib
import email
from email.header import decode_header
from flask import Flask, request, redirect
from apscheduler.schedulers.background import BackgroundScheduler
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage


LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
REDIRECT_URI = os.environ.get("REDIRECT_URI")
USERS_FILE = "./persistent/users.json"
COUNT_FILE = "./persistent/notify_counts.json"
NOTIFY_LIMIT = 30

app = Flask(__name__)
line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)


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
    directory = os.path.dirname(COUNT_FILE)  
    if not os.path.exists(directory):        
        os.makedirs(directory, exist_ok=True) 
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

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_users(users):
    directory = os.path.dirname(USERS_FILE) 
    if not os.path.exists(directory):      
        os.makedirs(directory, exist_ok=True)  
    with open(USERS_FILE, 'w') as f:       
        json.dump(users, f, indent=2)       


def find_user_by_line_id(line_user_id):
    users = load_users()
    for user in users:
        if user.get("LINE_USER_ID") == line_user_id:
            return user
    return None

def find_user_by_state(state):
    users = load_users()
    for user in users:
        if user.get("state") == state:
            return user
    return None

def update_user_tokens(state, access_token, refresh_token, token_expiry, email_address):
    users = load_users()
    for user in users:
        if user.get("state") == state:
            user["access_token"] = access_token
            user["refresh_token"] = refresh_token
            user["token_expiry"] = token_expiry
            user["EMAIL_ADDRESS"] = email_address
            user["IMAP_SERVER"] = "imap.gmail.com"
            user["IMAP_PORT"] = 993
            break
    save_users(users)


# === LINE Webhook ===
@app.route("/line-callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except Exception as e:
        print(f"LINE callback error: {e}")
        return "Error", 500
    return "OK"

def main():
    users = load_users()
    counts = load_notify_counts()
    for user in users:
        check_email(user, users, counts)


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    line_user_id = event.source.user_id
    user = find_user_by_line_id(line_user_id)

    if not user:
        # 新規ユーザーなら state を生成して保存
        state = str(uuid.uuid4())
        users = load_users()
        users.append({
            "LINE_USER_ID": line_user_id,
            "state": state,
            "EMAIL_ADDRESS": "",
            "IMAP_SERVER": "",
            "IMAP_PORT": "",
            "access_token": "",
            "refresh_token": "",
            "token_expiry": ""
        })
        save_users(users)
    else:
        state = user["state"]

    # 認証URLを返信
    auth_url = (
        f"https://accounts.google.com/o/oauth2/v2/auth"
        f"?client_id={CLIENT_ID}"
        f"&redirect_uri={REDIRECT_URI}"
        f"&response_type=code"
        f"&scope=https://mail.google.com/ https://www.googleapis.com/auth/userinfo.email"
        f"&access_type=offline"
        f"&prompt=consent"
        f"&state={state}"
    )

    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=f"Google認証をお願いします：\n{auth_url}")
    )

# === Google OAuth コールバック ===
@app.route('/callback')
def oauth2callback():
    code = request.args.get('code')
    state = request.args.get('state')

    # トークン取得
    token_url = 'https://oauth2.googleapis.com/token'
    data = {
        'code': code,
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'redirect_uri': REDIRECT_URI,
        'grant_type': 'authorization_code'
    }
    r = requests.post(token_url, data=data)
    token_response = r.json()

    access_token = token_response.get("access_token")
    refresh_token = token_response.get("refresh_token")
    expires_in = token_response.get("expires_in")

    if not access_token or not refresh_token:
        return "認証に失敗しました。"

    expiry_time = (datetime.datetime.utcnow() + datetime.timedelta(seconds=expires_in)).isoformat()

    # ユーザー情報取得
    userinfo_response = requests.get(
        "https://www.googleapis.com/oauth2/v2/userinfo",
        headers={"Authorization": f"Bearer {access_token}"}
    )
    userinfo = userinfo_response.json()
    email_address = userinfo.get("email")

    if not email_address:
        return "メールアドレスの取得に失敗しました。"

    # ユーザー情報更新
    update_user_tokens(state, access_token, refresh_token, expiry_time, email_address)

    return "Google認証が完了しました！LINEで通知が届きます。"


# === ルート（/）にアクセスしたときの表示 ===
@app.route('/')
def home():
    return 'Merutsuuchi は正常に動作中です！'

def main():
    print("===== main() started =====")
    users = load_users()
    counts = load_notify_counts()
    for user in users:
        try:
            check_email(user, users, counts)
        except Exception as e:
            print(f"[{user.get('LINE_USER_ID', '不明')}] ❌ check_emailで例外発生: {e}")

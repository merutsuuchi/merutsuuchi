import json
import os
import uuid
import datetime
import requests
from flask import Flask, request, redirect
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage

app = Flask(__name__)

# ① LINE Bot設定（↓ここを自分の情報に変更）
LINE_CHANNEL_ACCESS_TOKEN = 'dummy_token'
LINE_CHANNEL_SECRET = 'dummy_secret'

CLIENT_ID = 'dummy_client_id'
CLIENT_SECRET = 'dummy_client_secret'
REDIRECT_URI = 'https://example.com/callback'

USERS_FILE = "users.json"

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_users(users):
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

def update_user_tokens(state, access_token, refresh_token, token_expiry):
    users = load_users()
    for user in users:
        if user.get("state") == state:
            user["access_token"] = access_token
            user["refresh_token"] = refresh_token
            user["token_expiry"] = token_expiry
            break
    save_users(users)

# === LINE Webhook ===
@app.route("/line-callback", methods=["POST"])
def callback():
    signature = request.headers["X-Line-Signature"]
    body = request.get_data(as_text=True)
    handler.handle(body, signature)
    return "OK"

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
        f"&scope=https://mail.google.com/"
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

    update_user_tokens(state, access_token, refresh_token, expiry_time)

    return "Google認証が完了しました！LINEで通知が届きます。"

# === ルート（/）にアクセスしたときの表示 ===
@app.route('/')
def home():
    return 'Merutsuuchi は正常に動作中です！'

# === メイン実行 ===
if __name__ == "__main__":
    app.run(debug=True)

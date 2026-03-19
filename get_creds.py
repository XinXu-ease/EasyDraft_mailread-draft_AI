# get_creds.py
import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json

# 你需要的 scopes，根据需要调整
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/calendar.events"
]

CLIENT_SECRET_FILE = "credentials.json"  # 你的 OAuth 客户端文件
TOKEN_FILE = "token.json"                  # 保存 creds 的文件

def get_and_save_creds():
    # 如果已有 token.json，则尝试直接加载并刷新（更友好）
    creds = None
    if os.path.exists(TOKEN_FILE):
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print("刷新凭证失败，将重新授权:", e)
                creds = None
        if not creds:
            # 发起本地 server 授权（会自动打开浏览器）
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            try:
                creds = flow.run_local_server(port=0)
            except Exception as e:
                print("自动打开浏览器失败，使用控制台模式，请复制输出的 URL 到浏览器完成授权。错误：", e)
                creds = flow.run_console()

        # 保存到 token.json 以便下次直接使用
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
        print(f"✅ 授权成功，凭证已保存到 {TOKEN_FILE}")

    else:
        print("已有有效凭证：", TOKEN_FILE)

    return creds

if __name__ == "__main__":
    creds = get_and_save_creds()
    # 简单展示 access token（仅调试用）
    try:
        print("access_token (前40字符):", creds.token[:40])
    except Exception:
        pass

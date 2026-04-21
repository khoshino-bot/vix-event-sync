#!/usr/bin/env python3
"""
担当者交代時のトークン再取得スクリプト。
ブラウザでGoogleログイン → 新しいリフレッシュトークンを表示する。
表示された値を GitHub Secrets に登録すればOK。
"""

from google_auth_oauthlib.flow import InstalledAppFlow
import json, sys

CLIENT_SECRETS = {
    "installed": {
        "client_id":     input("Google OAuth クライアントID: ").strip(),
        "client_secret": input("Google OAuth クライアントシークレット: ").strip(),
        "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
        "token_uri":     "https://oauth2.googleapis.com/token",
        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob", "http://localhost"],
    }
}

GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
]
SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

print("\n--- Gmail 認証 ---")
flow = InstalledAppFlow.from_client_config(CLIENT_SECRETS, GMAIL_SCOPES)
creds = flow.run_local_server(port=0)
print(f"\nGMAIL_REFRESH_TOKEN={creds.refresh_token}")

print("\n--- Sheets 認証 ---")
flow2 = InstalledAppFlow.from_client_config(CLIENT_SECRETS, SHEETS_SCOPES)
creds2 = flow2.run_local_server(port=0)
print(f"\nSHEETS_REFRESH_TOKEN={creds2.refresh_token}")

print("\n上記の値を GitHub Secrets に登録してください。")

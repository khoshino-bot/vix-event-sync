#!/usr/bin/env python3
"""
月次シート自動作成スクリプト

「テンプレート」シートを複製して当月の YYMM（例: 2605）シートを作成する。
シートが既に存在する場合はスキップ（冪等）。

使用方法:
  python3 create_monthly_sheet.py

必要な環境変数:
  SHEETS_CLIENT_ID, SHEETS_CLIENT_SECRET, SHEETS_REFRESH_TOKEN, SPREADSHEET_ID
  TEMPLATE_SHEET_NAME (省略時: "テンプレート")
"""

import os
from datetime import datetime
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SPREADSHEET_ID      = os.environ.get("SPREADSHEET_ID", "")
TEMPLATE_SHEET_NAME = os.environ.get("TEMPLATE_SHEET_NAME", "テンプレート")


def build_creds():
    creds = Credentials(
        token=None,
        refresh_token=os.environ["SHEETS_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["SHEETS_CLIENT_ID"],
        client_secret=os.environ["SHEETS_CLIENT_SECRET"],
    )
    creds.refresh(Request())
    return creds


def main():
    if not SPREADSHEET_ID:
        raise SystemExit("[ERROR] 環境変数 SPREADSHEET_ID が設定されていません")

    now    = datetime.now()
    target = f"{str(now.year)[2:]}{str(now.month).zfill(2)}"   # 例: "2605"
    print(f"=== 月次シート作成 {now.strftime('%Y-%m-%d %H:%M:%S')} ===")
    print(f"対象シート: {target}")

    svc    = build("sheets", "v4", credentials=build_creds())
    meta   = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = meta["sheets"]
    titles = {s["properties"]["title"]: s for s in sheets}

    # 既に存在する場合はスキップ
    if target in titles:
        print(f"シート '{target}' は既に存在します。スキップ。")
        return

    # テンプレートシートの確認
    if TEMPLATE_SHEET_NAME not in titles:
        raise SystemExit(
            f"[ERROR] テンプレートシート '{TEMPLATE_SHEET_NAME}' が見つかりません。\n"
            f"スプレッドシート内に '{TEMPLATE_SHEET_NAME}' シートを作成してください。"
        )

    template_id = titles[TEMPLATE_SHEET_NAME]["properties"]["sheetId"]

    # YYMM シート群の末尾の次の位置に挿入（4桁数字のシートを時系列順に並べる）
    yymm_sheets = sorted(
        [s for s in sheets
         if len(s["properties"]["title"]) == 4 and s["properties"]["title"].isdigit()],
        key=lambda s: s["properties"]["title"]
    )
    insert_index = (yymm_sheets[-1]["properties"]["index"] + 1) if yymm_sheets else 1

    res = svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"duplicateSheet": {
            "sourceSheetId":    template_id,
            "insertSheetIndex": insert_index,
            "newSheetName":     target,
        }}]}
    ).execute()

    new_id = res["replies"][0]["duplicateSheet"]["properties"]["sheetId"]
    print(f"シート '{target}' を作成しました (sheetId={new_id})")
    print("=== 完了 ===")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
イベント起案・振り返り自動チェックスクリプト
GitHub Actions で動作。環境変数から認証情報を取得する。

Gmail検索: vix.co.jp ドメインからの起案/振り返りメールを検知
Sheets更新: 店舗名+実施日で行を照合して B列(起案) / C列(振り返り) に「〇」を入力
"""

import os
import re
import json
from datetime import datetime, date
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ===== 設定 =====
SPREADSHEET_ID   = os.environ.get("SPREADSHEET_ID", "1ia4DThYgqZ3WdyeoQcBpKj5c_t8bJdYgFZIY7Ge2ubE")
PROCESSED_FILE   = "processed_ids.json"  # Actions のワーキングディレクトリに保存
SENDER_DOMAIN    = os.environ.get("SENDER_DOMAIN", "vix.co.jp")

# 既知の店舗名（シートのD列表記 → マッチキーワードリスト）
STORE_MAP = {
    "経堂":           ["経堂"],
    "日暮里":         ["日暮里"],
    "町田":           ["町田"],
    "ゆめが丘":       ["ゆめが丘", "ゆめがおか"],
    "立川":           ["立川"],
    "三軒茶屋":       ["三軒茶屋"],
    "木の葉モール橋本": ["木の葉", "橋本"],
}

# ===== 認証（環境変数から構築）=====
def build_creds_from_env(prefix):
    """
    環境変数 {prefix}_CLIENT_ID / CLIENT_SECRET / REFRESH_TOKEN から Credentials を生成。
    GitHub Secrets として保存された認証情報を使う。
    """
    return Credentials(
        token=None,
        refresh_token=os.environ[f"{prefix}_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ[f"{prefix}_CLIENT_ID"],
        client_secret=os.environ[f"{prefix}_CLIENT_SECRET"],
    )

# ===== 処理済みID管理 =====
def load_processed():
    if os.path.exists(PROCESSED_FILE):
        with open(PROCESSED_FILE) as f:
            return set(json.load(f))
    return set()

def save_processed(ids):
    with open(PROCESSED_FILE, "w") as f:
        json.dump(list(ids), f, indent=2)

# ===== 件名から種別・店舗・日付を抽出 =====
def parse_subject(subject):
    if "起案" in subject:
        kind = "起案"
    elif "振り返り" in subject:
        kind = "振り返り"
    else:
        return None

    matched_store = None
    for store, keywords in STORE_MAP.items():
        if any(kw in subject for kw in keywords):
            matched_store = store
            break
    if not matched_store:
        print(f"  [WARN] 店舗名を特定できません: {subject}")
        return None

    dates = extract_dates(subject)
    if not dates:
        print(f"  [WARN] 日付を抽出できません: {subject}")
        return None

    return kind, [(matched_store, d) for d in dates]

def extract_dates(text):
    found = []
    now = datetime.now()
    year = now.year

    # YYYYMMDD
    for m in re.finditer(r'(20\d{2})(\d{2})(\d{2})', text):
        try:
            found.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass

    # YYYY年M月D日
    for m in re.finditer(r'(20\d{2})年(\d{1,2})月(\d{1,2})日', text):
        try:
            found.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass

    # M/D（年はスクリプト実行年を使用）
    for m in re.finditer(r'(\d{1,2})[/月](\d{1,2})', text):
        try:
            month, day = int(m.group(1)), int(m.group(2))
            if 1 <= month <= 12 and 1 <= day <= 31:
                found.append(date(year, month, day))
        except ValueError:
            pass

    return list(dict.fromkeys(found))

# ===== シート操作 =====
def get_current_sheet_name():
    now = datetime.now()
    return f"{str(now.year)[2:]}{str(now.month).zfill(2)}"

def load_sheet_data(service, sheet_name):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A1:Z200"
    ).execute()
    rows = result.get("values", [])

    row_map = {}
    for i, row in enumerate(rows):
        if i < 3:
            continue
        store   = row[3].strip() if len(row) > 3 else ""
        date_str = row[4].strip() if len(row) > 4 else ""
        if not store or not date_str:
            continue
        try:
            d = datetime.strptime(date_str, "%Y/%m/%d").date()
            row_map[(store, d)] = i + 1
        except ValueError:
            pass

    return row_map

def update_cell(service, sheet_name, row_num, col):
    col_letter = "B" if col == "起案" else "C"
    cell = f"'{sheet_name}'!{col_letter}{row_num}"

    current = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=cell
    ).execute()
    val = current.get("values", [[""]])[0][0] if current.get("values") else ""
    if val == "〇":
        return False

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=cell,
        valueInputOption="RAW",
        body={"values": [["〇"]]}
    ).execute()
    return True

# ===== メイン =====
def main():
    print(f"=== 同期開始 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    gmail_creds  = build_creds_from_env("GMAIL")
    sheets_creds = build_creds_from_env("SHEETS")

    # トークン自動更新
    gmail_creds.refresh(Request())
    sheets_creds.refresh(Request())

    gmail_svc  = build("gmail",  "v1", credentials=gmail_creds)
    sheets_svc = build("sheets", "v4", credentials=sheets_creds)

    processed  = load_processed()
    sheet_name = get_current_sheet_name()
    print(f"対象シート: {sheet_name}")

    try:
        row_map = load_sheet_data(sheets_svc, sheet_name)
    except Exception as e:
        print(f"[ERROR] シート '{sheet_name}' を読めません: {e}")
        return

    print(f"シート行数: {len(row_map)}件")

    updated_count = 0
    # cc:me ではなく送信元ドメインで検索（特定個人に依存しない）
    for keyword in ["起案", "振り返り"]:
        query = f"from:@{SENDER_DOMAIN} {keyword}"
        result = gmail_svc.users().messages().list(
            userId="me", q=query, maxResults=50
        ).execute()
        messages = result.get("messages", [])
        print(f"\nGmail '{query}': {len(messages)}件")

        for m in messages:
            if m["id"] in processed:
                continue

            msg = gmail_svc.users().messages().get(
                userId="me", id=m["id"], format="metadata",
                metadataHeaders=["Subject", "Date"]
            ).execute()
            headers  = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
            subject  = headers.get("Subject", "")

            parsed = parse_subject(subject)
            if not parsed:
                processed.add(m["id"])
                continue

            kind, pairs = parsed
            print(f"  [{kind}] {subject[:70]}")

            matched_any = False
            for store, d in pairs:
                key = (store, d)
                if key not in row_map:
                    print(f"    → 行なし: {store} {d}")
                    continue
                matched_any = True
                row_num = row_map[key]
                updated = update_cell(sheets_svc, sheet_name, row_num, kind)
                status  = "✓ 更新" if updated else "- スキップ(既存)"
                print(f"    → {store} {d} 行{row_num} {kind} {status}")
                if updated:
                    updated_count += 1

            if matched_any:
                processed.add(m["id"])

    save_processed(processed)
    print(f"\n=== 完了: {updated_count}件更新 ===")

if __name__ == "__main__":
    main()

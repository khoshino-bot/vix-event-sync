#!/usr/bin/env python3
"""
イベント起案・振り返り自動チェックスクリプト

処理フロー:
  起案メール  → pptx/xlsx添付を読む → 正規表現 + Claude API で二重検証
              → 一致: シートに行挿入 + 起案〇
              → 不一致: 送信者に通知メール
  振り返りメール → xlsx添付を読む → 同様に二重検証
              → 一致: 既存行の振り返り〇
              → 不一致: 送信者に通知メール
"""

import os, re, json, base64, tempfile
from datetime import datetime, date
from email.mime.text import MIMEText
import anthropic
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from pptx import Presentation
import openpyxl

# ===== 設定 =====
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1ia4DThYgqZ3WdyeoQcBpKj5c_t8bJdYgFZIY7Ge2ubE")
PROCESSED_FILE = "processed_ids.json"
SENDER_DOMAIN  = os.environ.get("SENDER_DOMAIN", "vix.co.jp")

# D列ドロップダウン選択肢と照合キーワード
STORE_MAP = {
    "経堂":     ["経堂"],
    "日暮里":   ["日暮里"],
    "町田":     ["町田", "コビルナ"],
    "ゆめが丘": ["ゆめが丘", "ゆめがおか", "ゆめが丘ソラトス"],
    "立川":     ["立川"],
    "三軒茶屋": ["三軒茶屋"],
    "竹ノ塚":   ["竹ノ塚"],
    "木の葉":   ["木の葉", "橋本"],
}

# ===== 認証 =====
def build_creds(prefix):
    creds = Credentials(
        token=None,
        refresh_token=os.environ[f"{prefix}_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ[f"{prefix}_CLIENT_ID"],
        client_secret=os.environ[f"{prefix}_CLIENT_SECRET"],
    )
    creds.refresh(Request())
    return creds

# ===== 処理済みID =====
def load_processed():
    return set(json.load(open(PROCESSED_FILE))) if os.path.exists(PROCESSED_FILE) else set()

def save_processed(ids):
    json.dump(list(ids), open(PROCESSED_FILE, "w"), indent=2)

# ===== テキスト抽出 =====
def extract_text_from_pptx(path):
    prs = Presentation(path)
    parts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    t = para.text.strip()
                    if t:
                        parts.append(t)
            if shape.has_table:
                for row in shape.table.rows:
                    cells = [c.text.strip() for c in row.cells if c.text.strip()]
                    if cells:
                        parts.append(" | ".join(cells))
    return "\n".join(parts)

def extract_text_from_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    parts = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows(values_only=True):
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cells:
                parts.append(" | ".join(cells))
    return "\n".join(parts)

def download_attachment(gmail_svc, msg_id, part):
    """添付ファイルをダウンロードして一時ファイルパスを返す"""
    fname = part.get("filename", "")
    att_id = part["body"].get("attachmentId", "")
    if not att_id:
        return None, fname
    att = gmail_svc.users().messages().attachments().get(
        userId="me", messageId=msg_id, id=att_id
    ).execute()
    data = base64.urlsafe_b64decode(att["data"])
    ext = os.path.splitext(fname)[1].lower()
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    tmp.write(data)
    tmp.close()
    return tmp.name, fname

def get_attachment_text(gmail_svc, msg_id, payload):
    """メッセージの添付ファイル(pptx/xlsx)を全て読んでテキストを返す"""
    texts = []
    def walk(parts):
        for part in parts:
            fname = part.get("filename", "")
            if fname.endswith(".pptx"):
                path, _ = download_attachment(gmail_svc, msg_id, part)
                if path:
                    texts.append(extract_text_from_pptx(path))
                    os.unlink(path)
            elif fname.endswith((".xlsx", ".xls")):
                path, _ = download_attachment(gmail_svc, msg_id, part)
                if path:
                    texts.append(extract_text_from_xlsx(path))
                    os.unlink(path)
            if "parts" in part:
                walk(part["parts"])
    walk(payload.get("parts", []))
    return "\n".join(texts)

# ===== 情報抽出（正規表現）=====
def normalize_store(text):
    for store, keywords in STORE_MAP.items():
        if any(kw in text for kw in keywords):
            return store
    return None

def extract_dates(text):
    found = []
    year = datetime.now().year

    for m in re.finditer(r'(20\d{2})(\d{2})(\d{2})', text):
        try:
            found.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass

    for m in re.finditer(r'(20\d{2})年(\d{1,2})月(\d{1,2})日', text):
        try:
            found.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass

    for m in re.finditer(r'(\d{1,2})[/月](\d{1,2})', text):
        try:
            mo, dy = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12 and 1 <= dy <= 31:
                found.append(date(year, mo, dy))
        except ValueError:
            pass

    return list(dict.fromkeys(found))

def extract_info_regex(text):
    """テキストから (store, [dates]) を抽出"""
    store = normalize_store(text)
    dates = extract_dates(text)
    return store, dates

# ===== 情報抽出（Claude API）=====
def extract_info_claude(text, kind):
    """Claude Haikuで店舗名と日付を抽出して {"store": "...", "dates": ["YYYY-MM-DD", ...]} を返す"""
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    store_list = list(STORE_MAP.keys())
    prompt = f"""以下は楽天モバイルイベントの{kind}書類から抽出したテキストです。
店舗名と実施日付を全て読み取り、JSON形式で返してください。

店舗名は必ず以下のいずれかにしてください: {store_list}
日付はYYYY-MM-DD形式で全ての日付をリストで返してください。

テキスト:
{text[:3000]}

JSON形式で答えてください（他のテキストは不要）:
{{"store": "店舗名", "dates": ["YYYY-MM-DD", ...]}}"""

    resp = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=256,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = resp.content[0].text.strip()
    # JSONを抽出
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    if not m:
        return None, []
    data = json.loads(m.group())
    store = data.get("store")
    dates = []
    for ds in data.get("dates", []):
        try:
            dates.append(datetime.strptime(ds, "%Y-%m-%d").date())
        except ValueError:
            pass
    return store, dates

# ===== 二重検証 =====
def double_verify(text, kind):
    """
    正規表現 + Claude API で検証。
    Returns: (ok, store, dates, diff_msg)
    """
    store_r, dates_r = extract_info_regex(text)
    store_c, dates_c = extract_info_claude(text, kind)

    print(f"    [正規表現] store={store_r} dates={dates_r}")
    print(f"    [Claude]   store={store_c} dates={dates_c}")

    if not store_r and not store_c:
        return False, None, [], "店舗名を特定できませんでした"
    if not dates_r and not dates_c:
        return False, None, [], "日付を特定できませんでした"

    store_ok = store_r == store_c
    dates_ok = set(dates_r) == set(dates_c)

    if store_ok and dates_ok:
        return True, store_r or store_c, sorted(set(dates_r) | set(dates_c)), ""

    diff = []
    if not store_ok:
        diff.append(f"店舗: 正規表現={store_r} / Claude={store_c}")
    if not dates_ok:
        only_r = sorted(set(dates_r) - set(dates_c))
        only_c = sorted(set(dates_c) - set(dates_r))
        if only_r:
            diff.append(f"正規表現のみ: {only_r}")
        if only_c:
            diff.append(f"Claudeのみ: {only_c}")

    return False, store_r or store_c, sorted(set(dates_r) | set(dates_c)), "\n".join(diff)

# ===== 通知メール送信 =====
def send_discrepancy_email(gmail_svc, original_msg, subject_store, subject_dates, attachment_diff):
    headers = {h["name"]: h["value"] for h in original_msg["payload"]["headers"]}
    to      = headers.get("From", "")
    subject = "Re: " + headers.get("Subject", "")

    body = f"""件名と添付ファイルの内容に相違が見つかりました。
確認・修正の上、再送してください。

【件名から読み取った情報】
店舗: {subject_store}
日付: {subject_dates}

【添付ファイルとの相違点】
{attachment_diff}

---
このメールは自動送信です。
"""
    msg_obj = MIMEText(body, "plain", "utf-8")
    msg_obj["To"] = to
    msg_obj["Subject"] = subject

    encoded = base64.urlsafe_b64encode(msg_obj.as_bytes()).decode()
    gmail_svc.users().messages().send(
        userId="me", body={"raw": encoded}
    ).execute()
    print(f"    → 通知メール送信: {to}")

# ===== シート操作 =====
def get_current_sheet_name():
    now = datetime.now()
    return f"{str(now.year)[2:]}{str(now.month).zfill(2)}"

def load_sheet_rows(sheets_svc, sheet_name):
    """(store, date) → 行番号 のマップと、全行の日付リスト(ソート済み)を返す"""
    result = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A1:Z300"
    ).execute()
    rows = result.get("values", [])

    row_map   = {}   # (store, date) -> row_num
    date_rows = []   # [(date, row_num)] ソート用

    for i, row in enumerate(rows):
        if i < 3:
            continue
        store    = row[3].strip() if len(row) > 3 else ""
        date_str = row[4].strip() if len(row) > 4 else ""
        if not store or not date_str:
            continue
        try:
            d = datetime.strptime(date_str, "%Y/%m/%d").date()
            row_map[(store, d)] = i + 1
            date_rows.append((d, i + 1))
        except ValueError:
            pass

    return row_map, sorted(date_rows)

def find_insert_position(date_rows, new_date):
    """新しい日付を日付順に挿入する行番号を返す（1-indexed）"""
    for d, row_num in date_rows:
        if new_date < d:
            return row_num
    if date_rows:
        return date_rows[-1][1] + 1
    return 4  # データ開始行

def insert_event_row(sheets_svc, sheet_name, store, d):
    """日付順の正しい位置に行を挿入し D・E・B 列を書き込む"""
    row_map, date_rows = load_sheet_rows(sheets_svc, sheet_name)
    insert_at = find_insert_position(date_rows, d)

    # 行を挿入（insertDimension）
    sheet_meta = sheets_svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id   = next(s["properties"]["sheetId"] for s in sheet_meta["sheets"]
                      if s["properties"]["title"] == sheet_name)

    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{
            "insertDimension": {
                "range": {
                    "sheetId":    sheet_id,
                    "dimension":  "ROWS",
                    "startIndex": insert_at - 1,  # 0-indexed
                    "endIndex":   insert_at,
                },
                "inheritFromBefore": True
            }
        }]}
    ).execute()

    # D・E・B 列に値をセット
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!B{insert_at}:E{insert_at}",
        valueInputOption="RAW",
        body={"values": [["〇", "", store, d.strftime("%Y/%m/%d")]]}
    ).execute()

    print(f"    → 行{insert_at}に挿入: {store} {d} 起案〇")

def update_cell(sheets_svc, sheet_name, row_num, col):
    col_letter = "B" if col == "起案" else "C"
    cell = f"'{sheet_name}'!{col_letter}{row_num}"
    cur  = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=cell
    ).execute()
    if cur.get("values", [[""]])[0][0] == "〇":
        return False
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=cell,
        valueInputOption="RAW", body={"values": [["〇"]]}
    ).execute()
    return True

# ===== メイン =====
def main():
    print(f"=== 同期開始 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    gmail_creds  = build_creds("GMAIL")
    sheets_creds = build_creds("SHEETS")
    gmail_svc    = build("gmail",  "v1", credentials=gmail_creds)
    sheets_svc   = build("sheets", "v4", credentials=sheets_creds)

    processed  = load_processed()
    sheet_name = get_current_sheet_name()
    print(f"対象シート: {sheet_name}")

    try:
        row_map, _ = load_sheet_rows(sheets_svc, sheet_name)
        print(f"既存行数: {len(row_map)}")
    except Exception as e:
        print(f"[ERROR] シート '{sheet_name}' を読めません: {e}")
        return

    updated = 0

    for kind, query in [("起案", f"from:@{SENDER_DOMAIN} 起案"),
                        ("振り返り", f"from:@{SENDER_DOMAIN} 振り返り")]:
        result = gmail_svc.users().messages().list(
            userId="me", q=query, maxResults=50
        ).execute()
        messages = result.get("messages", [])
        print(f"\nGmail '{query}': {len(messages)}件")

        for m in messages:
            if m["id"] in processed:
                continue

            msg = gmail_svc.users().messages().get(
                userId="me", id=m["id"], format="full"
            ).execute()
            headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
            subject = headers.get("Subject", "")
            print(f"\n  [{kind}] {subject[:70]}")

            # 件名から暫定抽出
            subj_store = normalize_store(subject)
            subj_dates = extract_dates(subject)

            # 添付ファイルを読む
            att_text = get_attachment_text(gmail_svc, m["id"], msg["payload"])

            if att_text.strip():
                # 二重検証
                ok, store, dates, diff = double_verify(att_text, kind)
            else:
                # 添付なし → 件名のみで進める
                print("    [添付なし] 件名のみで処理")
                ok, store, dates, diff = True, subj_store, subj_dates, ""

            if not ok:
                # 相違あり → 通知メール
                print(f"    [不一致] {diff}")
                send_discrepancy_email(
                    gmail_svc, msg,
                    f"{subj_store} / {subj_dates}",
                    subj_dates, diff
                )
                processed.add(m["id"])
                continue

            if not store or not dates:
                print(f"    [WARN] 店舗または日付を特定できません")
                processed.add(m["id"])
                continue

            # シートを更新
            if kind == "起案":
                for d in dates:
                    insert_event_row(sheets_svc, sheet_name, store, d)
                    updated += 1
            else:  # 振り返り
                row_map, _ = load_sheet_rows(sheets_svc, sheet_name)
                matched = False
                for d in dates:
                    key = (store, d)
                    if key in row_map:
                        u = update_cell(sheets_svc, sheet_name, row_map[key], "振り返り")
                        print(f"    → {store} {d} 行{row_map[key]} 振り返り {'✓' if u else '(既存)'}")
                        if u:
                            updated += 1
                        matched = True
                if not matched:
                    print(f"    [行なし] {store} {dates} → 次回再試行")
                    continue  # 処理済みにしない

            processed.add(m["id"])

    save_processed(processed)
    print(f"\n=== 完了: {updated}件更新 ===")

if __name__ == "__main__":
    main()

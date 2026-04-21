#!/usr/bin/env python3
"""
イベント起案・振り返り自動チェックスクリプト（無料版）

検証ロジック:
  ① 件名から店舗・日付を正規表現で抽出
  ② 添付ファイル(pptx/xlsx)から店舗・日付を正規表現で抽出
  ① == ② → シート書き込み
  ① != ② → 送信者に通知メール（シートは更新しない）
  添付なし → 件名のみで処理
"""

import os, re, json, base64, tempfile
from datetime import datetime, date
from email.mime.text import MIMEText
import email.utils as email_utils
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from pptx import Presentation
import openpyxl

# ===== 設定 =====
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1ia4DThYgqZ3WdyeoQcBpKj5c_t8bJdYgFZIY7Ge2ubE")
PROCESSED_FILE = "processed_ids.json"
SENDER_DOMAIN  = os.environ.get("SENDER_DOMAIN", "vix.co.jp")

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

# ===== 情報抽出（正規表現）=====
def normalize_store(text):
    for store, keywords in STORE_MAP.items():
        if any(kw in text for kw in keywords):
            return store
    return None

def extract_dates(text):
    found = []
    year = datetime.now().year

    # YYYYMMDD 形式 (例: 20260419)
    for m in re.finditer(r'(20\d{2})(\d{2})(\d{2})', text):
        try:
            found.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass

    # YYYY年M月D日 形式
    for m in re.finditer(r'(20\d{2})年(\d{1,2})月(\d{1,2})日', text):
        try:
            found.append(date(int(m.group(1)), int(m.group(2)), int(m.group(3))))
        except ValueError:
            pass

    # M月D,D,D日 形式 (例: 4月3,6,9,22,23日)
    for m in re.finditer(r'(\d{1,2})月([\d,、・\s]+)日', text):
        mo = int(m.group(1))
        if not (1 <= mo <= 12):
            continue
        for day_str in re.split(r'[,、・\s]+', m.group(2)):
            day_str = day_str.strip()
            if not day_str:
                continue
            try:
                dy = int(day_str)
                if 1 <= dy <= 31:
                    found.append(date(year, mo, dy))
            except ValueError:
                pass

    # M/D 形式 (例: 4/5, 4/17)
    for m in re.finditer(r'(\d{1,2})/(\d{1,2})', text):
        try:
            mo, dy = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12 and 1 <= dy <= 31:
                found.append(date(year, mo, dy))
        except ValueError:
            pass

    return sorted(set(found))

# ===== 添付ファイル読み取り =====
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

def get_attachment_text(gmail_svc, msg_id, payload):
    texts = []
    def walk(parts):
        for part in parts:
            fname = part.get("filename", "")
            att_id = part["body"].get("attachmentId", "")
            if not att_id:
                if "parts" in part:
                    walk(part["parts"])
                continue
            att = gmail_svc.users().messages().attachments().get(
                userId="me", messageId=msg_id, id=att_id
            ).execute()
            data = base64.urlsafe_b64decode(att["data"])
            ext = os.path.splitext(fname)[1].lower()
            tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
            tmp.write(data)
            tmp.close()
            try:
                if fname.endswith(".pptx"):
                    texts.append(extract_text_from_pptx(tmp.name))
                elif fname.endswith((".xlsx", ".xls")):
                    texts.append(extract_text_from_xlsx(tmp.name))
            finally:
                os.unlink(tmp.name)
            if "parts" in part:
                walk(part["parts"])
    walk(payload.get("parts", []))
    return "\n".join(texts)

# ===== 二重検証（件名 vs 添付）=====
def verify(subject, att_text):
    """
    件名と添付ファイルのテキストを正規表現で比較。
    Returns: (ok, store, dates, diff_msg)
    """
    s_store = normalize_store(subject)
    s_dates = extract_dates(subject)
    a_store = normalize_store(att_text)
    a_dates = extract_dates(att_text)

    print(f"    [件名]   store={s_store} dates={s_dates}")
    print(f"    [添付]   store={a_store} dates={a_dates}")

    # 添付に情報がなければ件名を信頼
    if not a_store and not a_dates:
        return True, s_store, s_dates, ""

    diffs = []
    if s_store and a_store and s_store != a_store:
        diffs.append(f"店舗: 件名={s_store} / 添付={a_store}")

    s_set, a_set = set(s_dates), set(a_dates)
    only_subject = sorted(s_set - a_set)
    only_attach  = sorted(a_set - s_set)
    if only_subject:
        diffs.append(f"件名のみの日付: {only_subject}")
    if only_attach:
        diffs.append(f"添付のみの日付: {only_attach}")

    if diffs:
        return False, a_store or s_store, sorted(a_set | s_set), "\n".join(diffs)

    # 一致 → 添付の情報を正とする
    return True, a_store or s_store, sorted(a_set | s_set), ""

# ===== 通知メール =====
def send_discrepancy_email(gmail_svc, original_msg, subject, diff_msg):
    headers = {h["name"]: h["value"] for h in original_msg["payload"]["headers"]}
    raw_from = headers.get("From", "")
    # 表示名付き "名前 <email@domain>" からメールアドレスのみ抽出
    _, addr = email_utils.parseaddr(raw_from)
    to   = addr if addr else raw_from
    subj = "Re: " + headers.get("Subject", "")
    body    = f"""件名と添付ファイルの内容に相違が見つかりました。
確認・修正の上、再送してください。

【相違点】
{diff_msg}

---
このメールは自動送信です。
"""
    msg_obj = MIMEText(body, "plain", "utf-8")
    msg_obj["To"]      = to
    msg_obj["Subject"] = subj
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
    result = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A1:Z300"
    ).execute()
    rows = result.get("values", [])
    row_map   = {}
    date_rows = []
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

def get_sheet_id(sheets_svc, sheet_name):
    meta = sheets_svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return next(s["properties"]["sheetId"] for s in meta["sheets"]
                if s["properties"]["title"] == sheet_name)

def insert_event_row(sheets_svc, sheet_name, store, d):
    _, date_rows = load_sheet_rows(sheets_svc, sheet_name)

    # 日付順の挿入位置を決定
    insert_at = 4
    for dt, row_num in date_rows:
        if d < dt:
            insert_at = row_num
            break
        insert_at = row_num + 1

    sheet_id = get_sheet_id(sheets_svc, sheet_name)
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"insertDimension": {
            "range": {
                "sheetId":    sheet_id,
                "dimension":  "ROWS",
                "startIndex": insert_at - 1,
                "endIndex":   insert_at,
            },
            "inheritFromBefore": True
        }}]}
    ).execute()

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

            # 添付ファイルを読む
            att_text = get_attachment_text(gmail_svc, m["id"], msg["payload"])

            if att_text.strip():
                ok, store, dates, diff = verify(subject, att_text)
            else:
                # 添付なし → 件名のみ
                print("    [添付なし] 件名のみで処理")
                store = normalize_store(subject)
                dates = extract_dates(subject)
                ok, diff = True, ""

            if not ok:
                print(f"    [不一致] {diff}")
                send_discrepancy_email(gmail_svc, msg, subject, diff)
                processed.add(m["id"])
                continue

            if not store or not dates:
                print("    [WARN] 店舗または日付を特定できません")
                processed.add(m["id"])
                continue

            if kind == "起案":
                for d in dates:
                    insert_event_row(sheets_svc, sheet_name, store, d)
                    updated += 1
            else:
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
                    continue

            processed.add(m["id"])

    save_processed(processed)
    print(f"\n=== 完了: {updated}件更新 ===")

if __name__ == "__main__":
    main()

# vix-event-sync

Gmail のイベントメール（起案・振り返り）を Google Sheets に自動反映するボット。
2時間ごとに GitHub Actions が起動し、添付ファイル検証・KPI 実績集計まで自動で処理する。

---

## 処理フロー

```
Gmail (from:@vix.co.jp)
  └─ 起案メール ─→ 添付(pptx/xlsx)と件名を照合
  │                 一致 → イベントシートに行挿入 + B列〇
  │                 相違 → 送信者に通知メール
  └─ 振り返りメール → 同様に照合 → C列〇

KPI スプレッドシート（貼り付けシート）
  └─ ｸﾛｰｻﾞｰ実績を日付・店舗で照合
       イベント日と一致 → X〜AA列（新規MNP/ひかり/Turbo/クレカ）
       イベント日と不一致 → AD〜AH列（店舗別戻り集計）
```

---

## GitHub Secrets 一覧

Settings → Secrets and variables → Actions に登録が必要な値：

| シークレット名 | 内容 | 取得方法 |
|--------------|------|---------|
| `GMAIL_CLIENT_ID` | Google OAuth クライアントID | GCP コンソール → OAuth 2.0 クライアント |
| `GMAIL_CLIENT_SECRET` | Google OAuth クライアントシークレット | 同上 |
| `GMAIL_REFRESH_TOKEN` | Gmail 読み取り・送信用トークン | `python3 reauth.py` で取得 |
| `SHEETS_CLIENT_ID` | Google OAuth クライアントID（Sheets用） | GCP コンソール（Gmail と同一プロジェクトなら同じ値） |
| `SHEETS_CLIENT_SECRET` | 同上 | 同上 |
| `SHEETS_REFRESH_TOKEN` | Sheets 読み書き用トークン | `python3 reauth.py` で取得 |
| `SPREADSHEET_ID` | イベント管理スプレッドシートの ID | URL の `/d/` と `/edit` の間の文字列 |
| `KPI_SPREADSHEET_ID` | KPI スプレッドシートの ID | 同上 |

---

## 通常運用

**月次作業：なし**
シート名・対象月はスクリプトが当月を自動検出するため、毎月の設定変更は不要。

**年次作業（任意）：トークンの確認**
`GMAIL_REFRESH_TOKEN` / `SHEETS_REFRESH_TOKEN` は通常ほぼ失効しないが、
長期間 Actions が止まっていた場合は下記の手順でトークンを更新する。

---

## 初期セットアップ手順（引き継ぎ時も同様）

### Step 1. 専用 Google アカウントを用意する

既存アカウントの個人依存を避けるため、専用アカウントを1つ作る。

- 例: `vix-sync@vix.co.jp`（Google Workspace）または `vix.events.sync@gmail.com`

### Step 2. メールの転送設定

全店舗からのイベントメールが新アカウントにも届くよう CC を追加する。

- 各店舗の起案・振り返りメール送信設定で `vix-sync@vix.co.jp` を CC に追加

### Step 3. スプレッドシートの共有

| スプレッドシート | 付与する権限 |
|---------------|------------|
| イベント管理シート | 編集者 |
| KPI シート | 閲覧者 |

### Step 4. GCP プロジェクトの確認（初回のみ）

既存の GCP プロジェクトを引き継ぐ場合は `GMAIL_CLIENT_ID` などを確認するだけでよい。  
新規に作る場合は以下の手順：

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクトを作成
2. API とサービス → ライブラリ → `Gmail API` と `Google Sheets API` を有効化
3. OAuth 同意画面 → 外部 → アプリ名・スコープ（gmail.modify, spreadsheets）を設定
4. 認証情報 → OAuth 2.0 クライアント ID → デスクトップアプリ → 作成
5. JSON をダウンロード（`client_secret.json`）

### Step 5. OAuth トークンを取得する

```bash
# リポジトリをクローンして実行
git clone https://github.com/khoshino-bot/vix-event-sync.git
cd vix-event-sync
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client

python3 reauth.py
# → ブラウザが開くので vix-sync アカウントでログイン・承認
# → 表示された CLIENT_ID / CLIENT_SECRET / REFRESH_TOKEN をメモ
```

`reauth.py` を2回実行する（Gmail 用・Sheets 用それぞれ）。

### Step 6. GitHub Secrets を登録する

リポジトリの Settings → Secrets and variables → Actions → New repository secret で
「GitHub Secrets 一覧」の8項目をすべて登録する。

### Step 7. 動作確認

Actions タブ → 「Run workflow」で手動実行し、ログに `=== 完了 ===` が表示されれば成功。

---

## トークンが切れたときの対応

```bash
python3 reauth.py
# ブラウザで再ログイン → 新しいトークンが表示される
```

表示された `REFRESH_TOKEN` を GitHub Secrets の `GMAIL_REFRESH_TOKEN` または
`SHEETS_REFRESH_TOKEN` に上書き登録する。

---

## 月次シートの自動作成

毎月1日 09:00（JST）に GitHub Actions が「テンプレート」シートを複製して
当月の YYMM シート（例: `2605`）を自動作成します。

### テンプレートシートの準備（初回1回だけ）

スプレッドシート内に **`テンプレート`** という名前のシートを作成する。

中身は既存の月シートと同じ構造にする：

| 行 | 内容 |
|----|------|
| 1〜3 | ヘッダー行（列名など） |
| 4 | 書式・ドロップダウン・数式が設定された空行（データなし） |

> テンプレートのデータ行（行4）には実際の値を入れず、書式だけ設定した状態にしておく。  
> sync_events.py が起案メールを検知するたびに、このテンプレート行の書式を引き継いで新しい行を挿入する。

### 手動でシートを作りたい場合

```bash
SHEETS_CLIENT_ID=xxx SHEETS_CLIENT_SECRET=xxx SHEETS_REFRESH_TOKEN=xxx \
SPREADSHEET_ID=xxx python3 create_monthly_sheet.py
```

または GitHub Actions の Actions タブ → `月次シート自動作成` → `Run workflow` で手動実行。

---

## 店舗を追加するには

`sync_events.py` を編集して2箇所追記し、push する。

```python
# 1. メール件名・添付から店舗名を認識するキーワード（37行目付近）
STORE_MAP = {
    ...
    "新店舗名": ["新店舗名", "別名があれば"],   # ← 追加
}

# 2. KPI スプレッドシートの店舗名との対応（56行目付近）
STORE_TO_KPI = {
    ...
    "新店舗名": "楽天新店舗名",   # ← 追加
}

# 3. シートでの表示順（49行目付近）
STORE_ORDER = [..., "新店舗名"]   # ← 末尾または適切な位置に追加
```

---

## ログの確認・トラブルシューティング

**ログの確認場所**  
GitHub → Actions タブ → 各実行のジョブ「sync」を開く

**よくある失敗パターン**

| エラーメッセージ | 原因 | 対処 |
|---------------|------|------|
| `invalid_grant` | トークン失効 | `python3 reauth.py` でトークン更新 |
| `[ERROR] 環境変数 XXX が設定されていません` | Secrets 未登録 | Settings → Secrets で登録 |
| `[ERROR] シート 'YYMM' を読めません` | 当月シートが未作成 | 対象シートを先に作成する |
| `HttpError 403` | スプレッドシートの共有設定漏れ | ボットアカウントに権限付与 |

---

## リポジトリのアクセス権

引き継ぎ先を GitHub リポジトリの Admin として追加する：  
Settings → Collaborators and teams → Add people → Admin 権限を付与

Admin 権限があれば Secrets の更新・Actions の管理がすべて可能になる。

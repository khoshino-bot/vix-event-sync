# vix-event-sync

GmailのイベントメールをGoogle Sheetsに自動反映するbot。

## 動作概要

30分ごとに GitHub Actions が起動し、
`@vix.co.jp` ドメインから届いた起案・振り返りメールを検知して
Google Sheetsの対応行に「〇」を自動入力する。

## 引き継ぎ手順

1. このリポジトリの Settings > Secrets and variables > Actions を開く
2. 下記シークレットを更新する

| シークレット名 | 内容 |
|--------------|------|
| `GMAIL_CLIENT_ID` | Google OAuth クライアントID |
| `GMAIL_CLIENT_SECRET` | Google OAuth クライアントシークレット |
| `GMAIL_REFRESH_TOKEN` | Gmailアクセス用リフレッシュトークン（再認証で更新） |
| `SHEETS_CLIENT_ID` | 同上（同一プロジェクトならGmailと同じ値） |
| `SHEETS_CLIENT_SECRET` | 同上 |
| `SHEETS_REFRESH_TOKEN` | Sheetsアクセス用リフレッシュトークン |
| `SPREADSHEET_ID` | 対象スプレッドシートのID |

### リフレッシュトークンの再取得

担当者が変わった場合、以下のスクリプトをローカルで実行して新しいトークンを取得する：

```bash
python3 reauth.py
```

取得したトークンを GitHub Secrets に上書き登録すれば完了。

## 店舗名の追加・変更

`sync_events.py` の `STORE_MAP` を編集する。

## ログの確認

GitHub の Actions タブから各実行のログを確認できる。

# 案件登録システム - Agent Memo

## プロジェクト概要

Googleスプレッドシートを活用した案件登録システムのスマートフォン対応ウェブアプリケーション。

## 技術スタック

- **バックエンド**: Google Apps Script (GAS)
- **フロントエンド**: Next.js (React)
- **デプロイ**: Vercel/Netlify
- **認証**: Google OAuth 2.0

## 主要機能

- 案件登録（月別シート / 月末請求シート）
- Googleアカウント認証
- マスターデータの取得（クライアント、ステータス、月）

## API URL

- **デプロイID**: `AKfycbwG8hKsFqUu7HaQWI3VZ69BBOZgUIP_UA3FSqgbseOcN0j-6KblDlvURldK_MGFh1_b`
- **URL**: `https://script.google.com/macros/s/AKfycbwG8hKsFqUu7HaQWI3VZ69BBOZgUIP_UA3FSqgbseOcN0j-6KblDlvURldK_MGFh1_b/exec`
- **バージョン**: 8 (2025/11/03 16:00)
- **ステータス**: ✅ 動作確認済み

## ファイル構成

- `コード.js` - Google Apps Scriptのバックエンドコード
- `web-app/` - Next.jsフロントエンドアプリケーション
- `appsscript.json` - Google Apps Script設定ファイル
- `.clasp.json` - clasp設定ファイル

## clasp push実行時のセットアップ

`clasp push`を実行する前に、毎回以下のセットアップが必要です：

```bash
# 1. claspでログイン（認証が必要な場合）
clasp logout  # 既存の認証をクリア
clasp login   # ブラウザで認証

# 2. コードをプッシュ
clasp push --force

# 3. デプロイを更新（必要に応じて）
clasp deploy -i AKfycbzX5QtLy5Q4K2ZCc2SgFBtRqGhaU-JMXmwN6fbo2vMXNffhZAf8q_HhGprOq0Qyjvvm -d "案件登録API - 更新"
```

**注意**: `clasp push`だけではWebアプリの設定は更新されません。デプロイの設定（Webアプリとして公開するなど）はGoogle Apps Scriptエディタで手動で設定する必要があります。


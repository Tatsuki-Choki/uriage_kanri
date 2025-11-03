# 案件登録システム

Googleスプレッドシートを活用した案件登録システムのスマートフォン対応ウェブアプリケーション。

## 技術スタック

- **バックエンド**: Google Apps Script (GAS)
- **フロントエンド**: Next.js (React) + TypeScript
- **デプロイ**: Vercel
- **認証**: Google OAuth 2.0
- **データベース**: Google Spreadsheet

## 主要機能

- 案件登録（月別シート / 月末請求シート）
- Googleアカウント認証
- マスターデータの取得（クライアント、ステータス、月）
- スマートフォン対応UI

## セットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/Tatsuki-Choki/uriage_kanri.git
cd uriage_kanri
```

### 2. Google Apps Script のデプロイ

1. `clasp` を使用してデプロイ

```bash
# claspでログイン
clasp login

# コードをプッシュ
clasp push --force
```

2. Google Apps ScriptエディタでWebアプリとして公開
   - 「デプロイ」→「デプロイを管理」
   - 種類: 「ウェブアプリ」を選択
   - アクセスできるユーザー: 「全員」を選択

### 3. フロントエンドのセットアップ

```bash
cd web-app
npm install
```

### 4. 環境変数の設定

`web-app/.env.local` を作成:

```env
NEXT_PUBLIC_GAS_API_URL=https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
NEXT_PUBLIC_GOOGLE_CLIENT_ID=YOUR_GOOGLE_CLIENT_ID
```

### 5. 開発サーバーの起動

```bash
npm run dev
```

## デプロイ

### Vercelへのデプロイ

```bash
cd web-app
vercel --prod
```

環境変数はVercelのダッシュボードで設定してください。

## API URL

- **デプロイID**: `AKfycbwG8hKsFqUu7HaQWI3VZ69BBOZgUIP_UA3FSqgbseOcN0j-6KblDlvURldK_MGFh1_b`
- **URL**: `https://script.google.com/macros/s/AKfycbwG8hKsFqUu7HaQWI3VZ69BBOZgUIP_UA3FSqgbseOcN0j-6KblDlvURldK_MGFh1_b/exec`
- **バージョン**: 8 (2025/11/03 16:00)

## ファイル構成

```
uriage_kanri/
├── コード.js              # Google Apps Scriptのバックエンドコード
├── appsscript.json        # Google Apps Script設定ファイル
├── .clasp.json            # clasp設定ファイル（.gitignoreに含まれる）
├── .claspignore           # clasp無視ファイル
├── agent.md               # 開発メモ
├── .gitignore             # Git無視ファイル
└── web-app/               # Next.jsフロントエンドアプリケーション
    ├── app/               # Next.js App Router
    ├── components/        # Reactコンポーネント
    ├── lib/               # ユーティリティ関数
    ├── public/            # 静的ファイル
    └── package.json       # 依存関係
```

## clasp push実行時のセットアップ

`clasp push`を実行する前に、毎回以下のセットアップが必要です：

```bash
# 1. claspでログイン（認証が必要な場合）
clasp logout  # 既存の認証をクリア
clasp login   # ブラウザで認証

# 2. コードをプッシュ
clasp push --force

# 3. デプロイを更新（必要に応じて）
clasp deploy -i AKfycbwG8hKsFqUu7HaQWI3VZ69BBOZgUIP_UA3FSqgbseOcN0j-6KblDlvURldK_MGFh1_b -d "案件登録API - 更新"
```

**注意**: `clasp push`だけではWebアプリの設定は更新されません。デプロイの設定（Webアプリとして公開するなど）はGoogle Apps Scriptエディタで手動で設定する必要があります。

## ライセンス

このプロジェクトは個人利用のためのものです。


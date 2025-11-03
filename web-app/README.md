# 案件登録ウェブアプリケーション

スプレッドシート連携の案件登録システムです。

## セットアップ

### 1. 環境変数の設定

`.env.local`ファイルを作成し、以下の環境変数を設定してください：

```env
NEXT_PUBLIC_GAS_API_URL=https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec
NEXT_PUBLIC_GOOGLE_CLIENT_ID=your-google-client-id.apps.googleusercontent.com
```

### 2. Google Apps Scriptの設定

1. Google Apps Scriptでスプレッドシートのスクリプトを開く
2. `コード.js`をデプロイ（Webアプリとして公開）
3. 公開設定：
   - 実行ユーザー: 自分
   - アクセスできるユーザー: 全員（または特定のユーザー）
4. デプロイ後に生成されるURLを`NEXT_PUBLIC_GAS_API_URL`に設定

### 3. Google OAuth設定

1. [Google Cloud Console](https://console.cloud.google.com/)でプロジェクトを作成
2. OAuth同意画面を設定
3. OAuth 2.0クライアントIDを作成
4. 承認済みのリダイレクトURIに以下を追加：
   - `http://localhost:3000` (開発環境)
   - `https://your-domain.com` (本番環境)
5. クライアントIDを`NEXT_PUBLIC_GOOGLE_CLIENT_ID`に設定

### 4. 依存関係のインストール

```bash
npm install
```

### 5. 開発サーバーの起動

```bash
npm run dev
```

ブラウザで [http://localhost:3000](http://localhost:3000) を開いてください。

## デプロイ

### Vercelへのデプロイ

1. [Vercel](https://vercel.com/)にログイン
2. プロジェクトをインポート
3. 環境変数を設定：
   - `NEXT_PUBLIC_GAS_API_URL`
   - `NEXT_PUBLIC_GOOGLE_CLIENT_ID`
4. デプロイ

### Netlifyへのデプロイ

1. [Netlify](https://www.netlify.com/)にログイン
2. プロジェクトをインポート
3. ビルド設定：
   - Build command: `npm run build`
   - Publish directory: `.next`
4. 環境変数を設定：
   - `NEXT_PUBLIC_GAS_API_URL`
   - `NEXT_PUBLIC_GOOGLE_CLIENT_ID`
5. デプロイ

## 使用方法

1. Googleアカウントでログイン
2. 案件情報を入力：
   - 登録先シート（月別シートまたは月末請求シート）
   - 月
   - 登録日
   - クライアント名
   - 業種（任意）
   - 経費（任意）
   - 経費（ドル）（月別シートの場合、任意）
   - 売上（必須）
   - ステータス（必須）
   - 備考（必須）
3. 「登録」ボタンをクリック
4. 登録完了画面で確認

## 機能

- Googleアカウント認証
- 案件登録（月別シート/月末請求シート）
- マスターデータの自動取得（クライアント、ステータス、月）
- レスポンシブデザイン（スマートフォン対応）

## 技術スタック

- Next.js 14
- React
- TypeScript
- Tailwind CSS
- Google OAuth 2.0
- Google Apps Script API

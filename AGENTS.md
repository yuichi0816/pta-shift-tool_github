# リポジトリガイドライン

## プロジェクト構成とモジュール整理
- `index.html`, `tool.html`, `manual.html`, `validator.html`, `region.html`, `gas-config.html`: ブラウザ向けUIページです。
- `js/script.js`, `js/script2.js`, `js/validator.js`: クライアントサイドの主要ロジックです。
- `css/style.css`, `css/style_gpt.css`: スタイル定義です。
- `functions/`: エッジ/サーバーレスのハンドラーです（`_middleware.js`, `functions/api/*`）。
- `gas/`: Google Apps Script 関連のユーティリティと仕様書です（`reminder_email.gs`, `SPECIFICATION.md`）。
- `input_data/`: 手動検証に使うサンプルスプレッドシートです。

## ビルド・テスト・開発コマンド
このリポジトリは静的な HTML/CSS/JS アプリで、ビルドパイプラインやパッケージマネージャーはありません。
- ローカル実行: `index.html` または `tool.html` をブラウザで直接開きます。
- 任意の静的サーバー: `python -m http.server 8000` を実行し、`http://localhost:8000/` にアクセスします。

## コーディングスタイルと命名規則
- JavaScript は 4 スペースインデント、セミコロンあり、`const`/`let`、関数宣言ベースで統一します。
- コメントとユーザー向け文言は、既存ファイルに合わせて日本語を基本とします。
- ファイル名は小文字を基本に、ハイフンまたはアンダースコアを使用します（例: `region-stats.js`, `css/style_gpt.css`）。

## テスト方針
- 自動テストフレームワークは未導入です。
- 手動確認として、`validator.html` で `input_data/` のサンプルファイルを検証し、`tool.html` の主要フロー（アップロード、割り当て、出力）を確認してください。

## コミットとプルリクエストのルール
- コミットメッセージは、履歴に合わせて短く具体的に記述します（日本語中心、必要に応じて `feat:` 形式）。
- PR では `.github/PULL_REQUEST_TEMPLATE.md` を使用し、概要、確認内容、UI変更時のスクリーンショットを記載してください。関連Issueがあればリンクしてください。

## Issue運用
- `.github/ISSUE_TEMPLATE/` 配下のテンプレート（バグ報告/機能要望）を使用してください。
- 一般的な質問や相談は Discussions（`config.yml` で設定）を利用してください。

## 設定とセキュリティ
- エッジハンドラーは環境バインディングを参照します（例: `functions/_middleware.js` の `REGION_KV`）。ホスティング環境で設定し、秘密情報はコミットしないでください。

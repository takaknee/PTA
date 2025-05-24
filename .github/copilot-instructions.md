# Copilot 指示ファイル

## プロジェクト情報
このプロジェクトは G-suits（Google Workspace）および Microsoft 365（M365）の各種スクリプトを生成・保管することを目的としています。

## コーディング規約
- 命名規則: camelCase を使用
- コメント: 日本語または英語でわかりやすく記述
- エラー処理: try-catch で適切に処理すること

## プロジェクト固有の要件
- スクリプトは再利用可能なモジュールとして設計する
- 機密情報（API キーなど）はコード内にハードコーディングせず、環境変数または設定ファイルから読み込む
- ログ記録を適切に実装し、実行状況を追跡可能にする

## セキュリティ要件
- 認証情報の適切な管理
- 入力値の検証によるインジェクション攻撃の防止
- API 利用時のレート制限への配慮

## 推奨ライブラリ
### Google Apps Script
- SpreadsheetApp, DriveApp, GmailApp などの Google サービス

### Microsoft 365
- Microsoft Graph API
- PowerShell モジュール (Microsoft.Graph, ExchangeOnlineManagement)

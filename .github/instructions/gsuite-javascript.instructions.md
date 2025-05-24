---
applyTo: "**/*.gs,**/*.js"
---
# G-suits (Google Apps Script) コーディング基準

[一般コーディングガイドライン](./.github/instructions/general-coding.instructions.md)をすべてのコードに適用してください。

## Google Apps Script ガイドライン
- Google サービス（SpreadsheetApp, DriveApp, GmailApp など）を適切に活用する
- スクリプトの実行時間制限（6分）を考慮し、長時間実行が必要な場合は分割実行を検討する
- トリガーを使用する場合は、重複実行を防止する仕組みを実装する
- ユーザー認証が必要な場合は、OAuth2 を適切に実装する

## JavaScript ガイドライン
- ES6+ の機能を積極的に活用する（Google Apps Script の対応状況を確認）
- 関数は小さく保ち、一つの機能に集中させる
- 変数のスコープを最小限に保つ
- グローバル変数の使用を避け、必要な場合は明確に命名する

## ログ記録
- Logger.log() を適切に使用し、実行状況を追跡可能にする
- 本番環境では console.log() ではなく適切なログ記録メカニズムを使用する
- エラー発生時には詳細情報をログに記録する

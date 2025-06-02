# GitHub Copilot 日本語化設定

このディレクトリには、GitHub Copilot を日本語で効果的に利用するための設定ファイルが含まれています。

## ファイル構成

### 基本設定ファイル
- `copilot-instructions.md` - Copilot の基本指示（日本語対応強化版）
- `instructions.md` - Copilot 利用指示書（日本語対応版）
- `prompt.md` - 日本語プロンプト例集

### 詳細指示ファイル（`instructions/` ディレクトリ）
- `general-coding.instructions.md` - 一般的なコーディング基準
- `gsuite-javascript.instructions.md` - Google Apps Script 用指示
- `m365-powershell.instructions.md` - Microsoft 365 PowerShell 用指示
- `copilot-commit-message.instructions.md` - コミットメッセージ作成指示
- `copilot-review-instructions.md` - コードレビュー指示
- `pull-request-description.instructions.md` - PR説明文テンプレート
- `test-instructions.md` - テスト作成ガイドライン

## 日本語化の特徴

### 1. 完全日本語対応
- すべての指示とプロンプトが日本語
- コード生成時のコメントも日本語
- エラーメッセージとログ出力も日本語

### 2. VS Code設定連携
`.vscode/settings.json` で以下を設定済み：
- `"github.copilot.chat.localeOverride": "ja"` - Copilot Chat の日本語化
- 各種指示ファイルの日本語テキスト設定
- コード生成時の日本語応答指示

### 3. PTA情報配信システム特化
- PTA特有の要件に対応したプロンプト例
- メール配信、アンケート管理、スケジュール通知等の機能に特化
- セキュリティとプライバシー保護の強化

## 使い方

### 1. Copilot Chat での日本語指示
```
「PTAメンバーにメール一括配信する機能を日本語コメント付きで作成して」
「アンケート結果を自動集計する機能を実装して、エラーハンドリングも日本語で」
```

### 2. コード生成時の日本語要求
- 「日本語でコメントを付けて」
- 「エラーメッセージも日本語で」
- 「ログ出力も日本語で」

### 3. プロンプト例の活用
`prompt.md` に記載された日本語プロンプト例を参考にして、具体的で効果的な指示を行う。

## 注意事項

- すべてのやり取りは日本語で行う
- 機密情報はコードに含めない
- セキュリティ要件を常に考慮する
- 既存のコードスタイルと整合性を保つ

## 設定の確認

VS Code で以下が有効になっていることを確認：
1. GitHub Copilot 拡張機能がインストール済み
2. `.vscode/settings.json` の設定が適用済み
3. `github.copilot.chat.localeOverride` が "ja" に設定済み

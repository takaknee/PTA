# GitHub Copilot 日本語化設定 + タスクベストプラクティス

このディレクトリには、GitHub Copilot を日本語で効果的に利用し、タスクベースの開発を効率化するための設定ファイルが含まれています。

## ファイル構成

### コア設定ファイル
- `copilot-instructions.md` - **Copilot メイン設定** (タスクベストプラクティス統合版)
- `instructions.md` - Copilot 利用指示書（日本語対応版）
- `prompt.md` - 日本語プロンプト例集

### テンプレート集 (`templates/` ディレクトリ)
- `copilot-development-guide.md` - **開発ガイド** (プロンプト例・トラブルシューティング・チーム開発)
- `gas-function-template.gs` - Google Apps Script 関数テンプレート
- `vba-module-template.bas` - VBA モジュールテンプレート
- `github-workflow-template.md` - **NEW!** GitHub ワークフローテンプレート集

### 詳細指示ファイル（`instructions/` ディレクトリ）
- `general-coding.instructions.md` - 一般的なコーディング基準
- `gsuite-javascript.instructions.md` - Google Apps Script 用指示
- `m365-powershell.instructions.md` - Microsoft 365 PowerShell 用指示
- `copilot-commit-message.instructions.md` - コミットメッセージ作成指示
- `copilot-review-instructions.md` - コードレビュー指示
- `pull-request-description.instructions.md` - PR説明文テンプレート
- `test-instructions.md` - テスト作成ガイドライン

## 🆕 タスクベース開発の新機能

### 1. 効果的なタスク分解
```
大きなタスクを段階的に分解：
要件分析 → 設計 → 実装 → 検証
```

### 2. GitHub Workflow 統合
- Issue-Driven Development
- Pull Request ベストプラクティス
- コミット規約とブランチ命名規約

### 3. 品質保証の強化
- 段階的テストアプローチ
- コード品質チェックリスト
- セキュリティとパフォーマンス要件

### 4. チーム開発支援
- ペアプログラミングガイド
- 知識継承のドキュメント化
- レビュー効率化

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

### 🎯 タスクベース開発の活用

#### 1. タスク開始時の効果的なプロンプト
```
このPTAプロジェクトで「メール配信システムの改修」を実装します。

要件:
- 技術: Google Apps Script
- 既存パターン: src/gsuite/pta/informationDistribution.gs 参考
- 制約: 既存の会員データ形式を維持
- 目標: 送信エラー率を50%削減

段階的に実装を進めてください。まず要件分析から始めましょう。
```

#### 2. Issue作成時の活用
```markdown
## 機能概要
メール配信の信頼性向上

## 要件
- [ ] エラーハンドリングの強化
- [ ] 再送機能の実装
- [ ] 配信状況の可視化

## 受け入れ条件
- [ ] 送信成功率95%以上
- [ ] エラー時の自動再送3回
- [ ] 管理者向けダッシュボード表示
```

#### 3. 段階的実装の指示
```
Step 1: 基本機能の確認
- 現在のメール送信処理をテストしてください
- エラー発生パターンを特定してください

Step 2: エラーハンドリング強化
- 各種例外ケースの処理を実装してください
- 日本語エラーメッセージを追加してください

Step 3: テストと検証
- 単体テストを作成してください
- 統合テストを実行してください
```

### 💻 従来の日本語指示（引き続き活用）

#### Copilot Chat での日本語指示
```
「PTAメンバーにメール一括配信する機能を日本語コメント付きで作成して」
「アンケート結果を自動集計する機能を実装して、エラーハンドリングも日本語で」
```

#### コード生成時の日本語要求
- 「日本語でコメントを付けて」
- 「エラーメッセージも日本語で」
- 「ログ出力も日本語で」

#### プロンプト例の活用
各テンプレートファイルに記載された日本語プロンプト例を参考にして、具体的で効果的な指示を行う。

### 📋 テンプレート活用ガイド

#### 新機能開発時
1. `github-workflow-template.md` でIssue作成
2. `copilot-development-guide.md` でプロンプト例確認
3. `gas-function-template.gs` または `vba-module-template.bas` で実装開始

#### バグ修正時
1. `github-workflow-template.md` のバグ報告テンプレート使用
2. `copilot-development-guide.md` のトラブルシューティング参照
3. 段階的デバッグ手法を適用

#### コードレビュー時
1. `copilot-development-guide.md` のチェックリスト使用
2. セキュリティ・パフォーマンス・保守性を重点確認
3. レビューコメントは日本語で記述

## 注意事項とベストプラクティス

### 📋 基本原則
- **すべてのやり取りは日本語で行う**
- **機密情報はコードに含めない**
- **セキュリティ要件を常に考慮する**
- **既存のコードスタイルと整合性を保つ**

### 🔧 タスク実行時の注意点
- **段階的な実装**: 小さな単位で確実に進行
- **テスト駆動**: 各段階でテスト実行
- **ドキュメント更新**: 実装と同時にドキュメント更新
- **レビュー重視**: 品質チェックリストの活用

### 🚫 禁止事項
- ❌ 英語でのコメントや説明
- ❌ テストなしでの本番デプロイ
- ❌ ドキュメント更新の忘れ
- ❌ セキュリティリスクのあるコード

## 🛠️ 設定の確認

### VS Code 設定
VS Code で以下が有効になっていることを確認：
1. **GitHub Copilot 拡張機能**がインストール済み
2. **`.vscode/settings.json` の設定**が適用済み
3. **`github.copilot.chat.localeOverride`** が "ja" に設定済み

### GitHub設定確認
1. **Issue テンプレート**: `.github/ISSUE_TEMPLATE/` に配置
2. **PR テンプレート**: `.github/pull_request_template.md` に配置
3. **ブランチ保護ルール**: mainブランチへの直接pushを制限

## 📚 さらなる学習リソース

### 内部リソース
- `docs/copilot-efficiency-improvements.md` - 改善の詳細記録
- `docs/outlook/specifications.md` - VBA実装仕様
- `src/` ディレクトリ - 実装例とパターン集

### 推奨する学習順序
1. **基本**: `copilot-instructions.md` で全体把握
2. **実践**: `templates/copilot-development-guide.md` でプロンプト学習
3. **ワークフロー**: `templates/github-workflow-template.md` で作業効率化
4. **応用**: 実際のタスクでパターン適用

## 🔄 継続的改善

このCopilot設定は定期的に更新されます：
- **月次**: ベストプラクティスの見直し
- **四半期**: 新機能・パターンの追加
- **年次**: 全体アーキテクチャの見直し

改善提案やフィードバックはIssueで受け付けています。

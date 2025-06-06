# PTA情報配信システム

PTA関連の情報配信を自動化するシステムです。Google Workspace（G Suite）とMicrosoft 365の連携スクリプトを提供します。

## 機能概要

### Google Workspace連携
- **PTA情報配信システム** (`src/gsuite/pta/`): PTAメンバー管理、情報配信、アンケート管理
- **Gmail整理スクリプト** (`src/gsuite/gmail/`): メール自動整理・クリーンアップ

### Microsoft Outlook連携  
- **Outlook AI Helper** (`src/outlook/`): OpenAI APIを使用したメール分析・作成支援

## CI/CD・品質管理

このプロジェクトでは、コード品質とセキュリティを自動的にチェックするGitHub Actionsを導入しています：

- ✅ **セキュリティスキャン**: API キー漏洩、脆弱性の検出
- ✅ **コード品質チェック**: ESLintによる静的解析
- ✅ **VBA セキュリティ分析**: 危険な関数の使用チェック
- ✅ **依存関係監査**: 脆弱性のある依存関係の検出

詳細は [CI/CD設定ドキュメント](docs/CI_CD_README.md) をご覧ください。

## クイックスタート

### 開発環境のセットアップ

```bash
# リポジトリをクローン
git clone https://github.com/takaknee/PTA.git
cd PTA

# 開発依存関係をインストール
npm install

# コード品質チェック
npm run lint

# セキュリティ監査
npm run security
```

### 各システムのセットアップ

- **PTA情報配信システム**: [セットアップガイド](src/gsuite/pta/SETUP.md)
- **Gmail整理スクリプト**: [使用方法](src/gsuite/gmail/README.md)  
- **Outlook AI Helper**: [クイックスタートガイド](docs/outlook/quickstart.md)

## セキュリティ注意事項

- API キーやパスワードをコード内にハードコーディングしないでください
- 本番環境では設定ファイルまたは環境変数を使用してください
- 定期的にアクセス権限を見直してください

## 貢献

プルリクエストやIssueを歓迎します。変更前に必ずCI/CDチェックが通ることを確認してください。

## ライセンス

MIT License

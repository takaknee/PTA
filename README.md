# PTA情報配信システム

PTA関連の情報配信を自動化するシステムです。Google Workspace（G Suite）とMicrosoft 365の連携スクリプトを提供します。

## 機能概要

### Google Workspace連携
- **PTA情報配信システム** (`src/gsuite/pta/`): PTAメンバー管理、情報配信、アンケート管理
- **Gmail整理スクリプト** (`src/gsuite/gmail/`): メール自動整理・クリーンアップ

### Microsoft Outlook連携  
- **Outlook AI Helper (VBA)** (`src/outlook/`): OpenAI APIを使用したメール分析・作成支援（従来版）
- **Outlook AI Helper (Office Add-in)** (`src/outlook-addin/`): 最新のOffice Add-ins技術による軽量版（推奨）
- **Outlook VSTO アドイン** (`src/outlook-vsto/`): C#による高機能版（レガシー）

### Edge拡張機能（NEW!）
- **PTA Edge拡張機能** (`src/edge-extension/`): Outlookアドインの代替案としてのEdgeブラウザ拡張機能
  - Web版Outlook・Gmail対応
  - Azure OpenAI API統合
  - ポリシー制限回避ソリューション

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
- **Outlook AI Helper (VBA)**: [クイックスタートガイド](docs/outlook/quickstart.md)
- **Edge拡張機能**: [インストール・設定ガイド](src/edge-extension/README.md)
- **Outlook Add-in (推奨)**: [セットアップガイド](src/outlook-addin/README.md)

## ✨ 最新情報：Office Add-ins 移行完了

**VSTOからOffice Add-ins（Office拡張機能）への移行が完了しました！**

### 🎯 推奨の移行パス

1. **従来のVBA版** → **Office Add-in版**（推奨）
2. 既存のVSTO版は将来的に廃止予定

### 🌟 Office Add-in版の利点

- ✅ **クロスプラットフォーム**: Windows/Mac/Web全対応
- ✅ **軽量配信**: マニフェストファイルのみで配信
- ✅ **自動更新**: ブラウザキャッシュによる透明な更新
- ✅ **簡単インストール**: ClickOnce不要
- ✅ **将来保証**: Microsoft推奨の最新技術

詳細は [Office Add-in移行ガイド](src/outlook-addin/MIGRATION.md) をご覧ください。

## セキュリティ注意事項

- API キーやパスワードをコード内にハードコーディングしないでください
- 本番環境では設定ファイルまたは環境変数を使用してください
- 定期的にアクセス権限を見直してください

## 貢献

プルリクエストやIssueを歓迎します。変更前に必ずCI/CDチェックが通ることを確認してください。

## ライセンス

MIT License

# PTA 情報配信システム

PTA 関連の情報配信を自動化するシステムです。Google Workspace（G Suite）と Microsoft 365 の連携スクリプトを提供します。

## 機能概要

### Google Workspace 連携

- **PTA 情報配信システム** (`src/gsuite/pta/`): PTA メンバー管理、情報配信、アンケート管理
- **Gmail 整理スクリプト** (`src/gsuite/gmail/`): メール自動整理・クリーンアップ

### Microsoft 365 連携

- **Excel AI Helper (VBA)** (`src/excel/`): OpenAI API を使用した Excel データ分析・処理支援

### Edge 拡張機能（NEW!）

- **PTA Edge 拡張機能** (`src/edge-extension/`): Web ブラウザ向けメール処理拡張機能
  - Web 版 Outlook・Gmail 対応
  - Azure OpenAI API 統合
  - ポリシー制限回避ソリューション

## CI/CD・品質管理

このプロジェクトでは、コード品質とセキュリティを自動的にチェックする GitHub Actions を導入しています：

- ✅ **セキュリティスキャン**: API キー漏洩、脆弱性の検出
- ✅ **コード品質チェック**: ESLint による静的解析
- ✅ **VBA セキュリティ分析**: 危険な関数の使用チェック
- ✅ **依存関係監査**: 脆弱性のある依存関係の検出

詳細は [CI/CD 設定ドキュメント](docs/CI_CD_README.md) をご覧ください。

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

- **PTA 情報配信システム**: [セットアップガイド](src/gsuite/pta/SETUP.md)
- **Gmail 整理スクリプト**: [使用方法](src/gsuite/gmail/README.md)
- **Outlook AI Helper (VBA)**: [クイックスタートガイド](docs/outlook/quickstart.md)
- **Edge 拡張機能**: [インストール・設定ガイド](src/edge-extension/README.md)
- **Outlook Add-in (推奨)**: [セットアップガイド](src/outlook-addin/README.md)

## ✨ 最新情報：Office Add-ins 移行完了

**VSTO から Office Add-ins（Office 拡張機能）への移行が完了しました！**

### 🎯 推奨の移行パス

1. **従来の VBA 版** → **Office Add-in 版**（推奨）
2. 既存の VSTO 版は将来的に廃止予定

### 🌟 Office Add-in 版の利点

- ✅ **クロスプラットフォーム**: Windows/Mac/Web 全対応
- ✅ **軽量配信**: マニフェストファイルのみで配信
- ✅ **自動更新**: ブラウザキャッシュによる透明な更新
- ✅ **簡単インストール**: ClickOnce 不要
- ✅ **将来保証**: Microsoft 推奨の最新技術

詳細は [Office Add-in 移行ガイド](src/outlook-addin/MIGRATION.md) をご覧ください。

## 🚀 最新アップデート：GPT-4o Mini 対応

2024 年 6 月更新で、全システムが最新の GPT-4o Mini モデルに対応しました：

- **⚡ 高速化**: 従来の GPT-4 比で約 50%の応答速度向上
- **💰 コスト効率**: 約 60%のコスト削減を実現
- **🎯 品質向上**: 日本語理解と文脈処理の改善
- **🔄 完全互換**: 既存設定からシームレスな移行

**対応コンポーネント**:

- Outlook Add-in: デフォルトモデルを`gpt-4o-mini`に更新
- Edge 拡張機能: 最新モデル選択肢を追加
- VBA（Outlook/Excel）: デフォルト設定を更新
- VSTO: 設定とエンドポイントを最適化

詳細は [GPT-4o Mini 移行ガイド](docs/GPT4O_MINI_MIGRATION.md) をご覧ください。

## セキュリティ注意事項

- API キーやパスワードをコード内にハードコーディングしないでください
- 本番環境では設定ファイルまたは環境変数を使用してください
- 定期的にアクセス権限を見直してください

## 貢献

プルリクエストや Issue を歓迎します。変更前に必ず CI/CD チェックが通ることを確認してください。

## ライセンス

MIT License

# Outlook OpenAI マクロ

Microsoft Outlook で Azure OpenAI API を活用してメール処理を効率化するVBAマクロ集です。

## 概要

このマクロは以下の課題を解決します：

### 受信メール処理の改善
- 複数回転送されたメールの元内容抽出
- 枕詞や挨拶文の除去による内容明確化  
- 期限、対象者、アクション項目の自動抽出

### 送信メール作成の効率化
- 営業メールに対する丁寧な断りメール自動生成
- 依頼・提案への適切な承諾メール自動生成

### 検索・分析機能の強化
- 検索フォルダの使用状況分析と最適化提案
- フォルダ内メール分析による新規検索フォルダ作成支援

## 主な機能

### 📧 メール内容解析 (`AI_EmailAnalyzer.bas`)
- 転送メールの元内容抽出
- メール内容の明確化・要約
- 重要情報抽出（期限・対象者・アクション）
- 包括的メール分析

### ✍️ メール作成支援 (`AI_EmailComposer.bas`)
- 営業メール断り文作成
- 承諾・同意メール作成
- カスタムビジネスメール生成

### 🔍 検索フォルダ分析 (`AI_SearchAnalyzer.bas`)
- 既存検索フォルダの使用状況分析
- フォルダ内メール傾向分析
- 重複・類似検索フォルダの検出
- 最適化提案とレポート生成

### ⚙️ 設定管理 (`AI_ConfigManager.bas`)
- API設定の確認・変更
- プロンプトカスタマイズ
- 設定のインポート・エクスポート
- 接続テスト機能

### 🚀 インストール支援 (`OutlookAI_Installer.bas`)
- ガイド付き初回セットアップ
- システム要件確認
- API設定ガイダンス
- 動作テスト実行

## 必要な環境

- Microsoft Outlook 2016 以降
- VBA マクロ実行権限
- Azure OpenAI サービスのアカウント
- インターネット接続

## インストール方法

### 1. ファイルのダウンロード
必要なVBAファイルをダウンロードします：
- `AI_Common.bas` - 共通関数・設定
- `AI_ApiConnector.bas` - OpenAI API接続
- `AI_EmailAnalyzer.bas` - メール解析機能
- `AI_EmailComposer.bas` - メール作成支援
- `AI_SearchAnalyzer.bas` - 検索フォルダ分析
- `AI_ConfigManager.bas` - 設定管理
- `OutlookAI_Installer.bas` - インストール支援

### 2. VBAモジュールのインポート
1. Outlook を起動
2. `Alt + F11` でVBAエディタを開く
3. `ファイル` → `ファイルのインポート` で各.basファイルをインポート

### 3. Azure OpenAI の設定
1. Azure ポータルでOpenAI リソースを作成
2. GPT-4 モデルをデプロイ
3. API キーとエンドポイントを取得

### 4. マクロの設定
1. `AI_Common.bas` を開く
2. 以下の定数を編集：
```vba
Public Const OPENAI_API_ENDPOINT As String = "https://your-resource.openai.azure.com/..."
Public Const OPENAI_API_KEY As String = "your-api-key"
```

### 5. セットアップガイドの実行
1. `Alt + F8` でマクロダイアログを開く
2. `OutlookAI_Installer.RunInitialSetup` を実行
3. ガイドに従って設定を完了

## 使用方法

### 基本的な操作手順
1. 分析・処理したいメールを選択
2. `Alt + F8` でマクロダイアログを開く
3. 実行したい機能のマクロを選択
4. 画面の指示に従って操作

### よく使用する機能
- **メール解析**: `AI_EmailAnalyzer.AnalyzeSelectedEmail`
- **営業断りメール**: `AI_EmailComposer.CreateRejectionEmail`  
- **承諾メール**: `AI_EmailComposer.CreateAcceptanceEmail`
- **検索フォルダ分析**: `AI_SearchAnalyzer.AnalyzeSearchFolders`
- **設定管理**: `AI_ConfigManager.ManageConfiguration`

## 設定のカスタマイズ

### プロンプトの調整
`AI_Common.bas` 内のシステムプロンプト定数を編集することで、AI の応答をカスタマイズできます：
- `SYSTEM_PROMPT_ANALYZER` - メール分析用
- `SYSTEM_PROMPT_COMPOSER` - メール作成用  
- `SYSTEM_PROMPT_SEARCH` - 検索分析用

### パフォーマンス設定
- `MAX_CONTENT_LENGTH` - 最大処理文字数
- `REQUEST_TIMEOUT` - API タイムアウト時間

## トラブルシューティング

### よくある問題

#### API接続エラー
- APIキーとエンドポイントを確認
- ネットワーク接続を確認
- Azure OpenAI サービスの制限を確認

#### マクロが表示されない
- VBAファイルが正しくインポートされているか確認
- マクロセキュリティ設定を確認
- Outlook を再起動

#### 処理が遅い・タイムアウト
- メール本文が大きすぎる可能性（50KB制限）
- API制限による待機時間
- ネットワーク環境の確認

## セキュリティについて

### API キーの管理
- コード内にAPIキーをハードコーディングしない（本番環境では）
- 定期的にAPIキーをローテーション
- 使用しないキーは無効化

### データの取り扱い
- 機密メールの処理時は内容を事前確認
- 処理ログは定期的に削除
- 外部送信データの最小化

## 開発・カスタマイズ

### アーキテクチャ
- モジュール分割設計による保守性確保
- 共通関数による重複排除
- エラーハンドリングの統一

### 拡張方法
- 新機能は独立したモジュールとして追加
- `AI_Common.bas` の共通関数を活用
- プロンプトテンプレートのカスタマイズ

## ライセンス

MIT License - 詳細は LICENSE ファイルを参照

## サポート

### バグ報告・機能要望
問題が発生した場合は、以下の情報を含めて報告してください：
- Outlook バージョン
- エラーメッセージ
- 実行していた操作
- VBA デバッグ情報

### コントリビューション
プルリクエストやイシューの報告を歓迎します。

## 更新履歴

### v1.0.0 (2024)
- 初回リリース
- 基本的なメール解析・作成機能
- 検索フォルダ分析機能
- 設定管理機能
- インストールガイド

## 関連ドキュメント

- [設計書・仕様書](docs/outlook/specifications.md)
- [設計案比較](docs/outlook/design-comparison.md)  
- [クイックスタートガイド](docs/outlook/quickstart.md)

---

このマクロを使用することで、日々のメール処理が大幅に効率化されることを願っています。
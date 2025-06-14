# PTA Outlook Add-in

Microsoft Office Add-ins技術を使用したPTA情報配信システム用のOutlookアドインです。

## 概要

VSTOからの移行により、より軽量で展開しやすいWeb技術ベースのOutlookアドインとして再実装されました。

### 主要機能

- **AI解析機能**: OpenAI/Azure OpenAI APIを使用したメール内容の自動解析
- **メール作成支援**: AIによる文章生成と挿入機能
- **タスクペイン**: 詳細な設定と履歴管理
- **リボン統合**: Outlookリボンからの直接アクセス

### 技術仕様

- **基盤技術**: Microsoft Office Add-ins (Office JavaScript API)
- **対応環境**: Outlook on Windows, Mac, Web
- **言語**: JavaScript/HTML/CSS
- **AIプロバイダー**: OpenAI, Azure OpenAI, Claude（予定）

## ファイル構成

```
src/outlook-addin/
├── manifest/
│   └── manifest.xml          # Office Add-in マニフェストファイル
├── src/
│   ├── read.html            # メール読み取り画面
│   ├── read.js              # メール解析機能
│   ├── compose.html         # メール作成画面
│   ├── compose.js           # メール作成支援機能
│   ├── taskpane.html        # 詳細パネル画面
│   ├── taskpane.js          # 詳細機能
│   ├── functions.html       # 関数実行用HTML
│   └── functions.js         # リボンボタン用関数
├── assets/
│   ├── styles.css           # 共通スタイルシート
│   ├── icon-16.png          # アイコン（16x16）
│   ├── icon-32.png          # アイコン（32x32）
│   ├── icon-80.png          # アイコン（80x80）
│   └── logo-filled.png      # PTAロゴ
├── package.json             # NPMパッケージ設定
└── README.md               # このファイル
```

## セットアップ

### 1. 依存関係のインストール

```bash
cd src/outlook-addin
npm install
```

### 2. 開発サーバーの起動

```bash
npm start
```

### 3. アドインのサイドロード

```bash
npm run sideload
```

## 開発

### ローカル開発サーバー

```bash
npm run start:web
```

ブラウザで https://localhost:3000 にアクセスして開発可能です。

### マニフェスト検証

```bash
npm run validate
```

### リント

```bash
npm run lint
npm run lint:fix  # 自動修正
```

## 配信

### サイドロード配信

1. `manifest/manifest.xml` をOutlookにインポート
2. 開発サーバーまたは静的ホスティングでファイルを提供

### 企業内配信

1. SharePoint App Catalogまたは
2. Microsoft 365 管理センターでの一元配信

### Microsoft AppSource

1. Partner Center での審査申請
2. 一般公開配信

## API設定

### OpenAI API

1. アドインの設定タブを開く
2. プロバイダーで「OpenAI」を選択
3. APIキーを入力
4. モデル（GPT-4/GPT-3.5）を選択

### Azure OpenAI

1. プロバイダーで「Azure」を選択
2. APIキーとエンドポイントを設定
3. デプロイメント名を指定

## 使用方法

### メール解析

1. Outlookでメールを選択
2. リボンの「メール解析」ボタンをクリック
3. 解析結果が通知に表示

### メール作成支援

1. 新規メール作成画面を開く
2. リボンの「作成支援」ボタンをクリック
3. フォームに内容を入力して生成
4. 生成された文章をメールに挿入

### 詳細機能

1. リボンの「詳細表示」ボタンでタスクペインを開く
2. 「解析」タブで詳細な解析オプション
3. 「設定」タブでAPI設定
4. 「履歴」タブで過去の操作履歴

## トラブルシューティング

### よくある問題

**Q: APIキーエラーが発生する**
A: 設定タブでAPIキーが正しく設定されているか確認してください。

**Q: アドインが表示されない**
A: マニフェストファイルが正しくインポートされているか確認してください。

**Q: HTTPS証明書エラー**
A: `npm run start` で開発証明書をインストールしてください。

### ログ確認

ブラウザの開発者ツールでコンソールログを確認してください。

## ライセンス

MIT License - 詳細は [LICENSE](../../LICENSE) ファイルを参照

## サポート

- GitHub Issues: https://github.com/takaknee/PTA/issues
- プロジェクトWiki: https://github.com/takaknee/PTA/wiki

## 移行について

### VSTOからの移行利点

- **軽量性**: Web技術による軽量実装
- **クロスプラットフォーム**: Windows, Mac, Web対応
- **簡単配信**: マニフェストファイルベースの配信
- **メンテナンス**: ブラウザ技術によるデバッグ・開発の容易さ

### 移行後の機能比較

| 機能 | VSTO版 | Office Add-in版 |
|------|--------|-----------------|
| メール解析 | ✅ | ✅ |
| メール作成支援 | ✅ | ✅ |
| AI API統合 | ✅ | ✅ |
| リボン統合 | ✅ | ✅ |
| タスクペイン | ✅ | ✅ |
| 設定管理 | ✅ | ✅ |
| 履歴管理 | ✅ | ✅ |
| 自動更新 | ClickOnce | Web自動更新 |
| 配信方法 | MSI/ClickOnce | マニフェスト |
| 対応OS | Windows のみ | Windows/Mac/Web |

詳細な移行ガイドは [MIGRATION.md](MIGRATION.md) を参照してください。
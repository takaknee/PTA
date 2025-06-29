# Edge 拡張機能 アップデート概要

## 🚀 M365統合＆VSCode設定解析機能追加 (v1.1.0)

### 新機能概要

#### 🔗 Microsoft 365統合機能
- **Teams Chat転送**: ページ情報をログインユーザーのTeams chatに自動転送
- **予定表登録**: 現在日時でページレビュー予定を自動作成
- **フォールバック機能**: API認証失敗時はWeb版アプリケーションを自動起動

#### ⚙️ VSCode設定解析機能
- **自動判定**: VSCodeドキュメントページの自動検出
- **設定抽出**: AI活用によるページ内設定項目の抽出・解析
- **サンプル生成**: 実用的なsettings.jsonサンプル自動生成
- **ワンクリックコピー**: 解析結果の簡単コピー機能

### 技術実装

#### 1. manifest.json
- Microsoft Graph API権限追加: `https://graph.microsoft.com/*`
- Chrome identity API権限追加: `identity`
- Microsoft認証エンドポイント追加: `https://login.microsoftonline.com/*`

#### 2. background.js
- `getMicrosoftGraphToken()`: Chrome identity APIによるMicrosoft Graph認証
- `handleForwardToTeams()`: Teams chat API連携とWeb版フォールバック
- `handleAddToCalendar()`: Calendar API連携とOutlook Web版フォールバック
- `handleAnalyzeVSCodeSettings()`: AI活用VSCode設定解析

#### 3. content.js
- 新UIボタン追加: 💬 Teams chatに転送、📅 予定表に追加、⚙️ VSCode設定解析
- コンテキストメニュー統合: 右クリックから各機能への直接アクセス
- フォールバック対応UI: Web版アプリ起動時の適切なガイダンス表示

### 対応サイト拡張
- **M365統合**: 全サイト対応（outlook.office.com、teams.microsoft.com連携）
- **VSCode解析**: code.visualstudio.com、marketplace.visualstudio.com対応

### セキュリティ＆可用性
- **段階的認証**: Chrome identity API → Microsoft Graph API
- **フォールバック**: API失敗時の自動Web版起動
- **エラーハンドリング**: 詳細なエラーメッセージと解決提案

## 🔄 リブランド対応 (PTA → AI 業務支援ツール)

### 変更概要

- **名称変更**: 「PTA 支援ツール」→「AI 業務支援ツール」
- **対象拡大**: PTA 特化 → 一般的な業務支援
- **機能追加**: URL 抽出、ページ情報コピー、翻訳機能

### 修正済みファイル

#### 1. manifest.json

- name: "AI 業務支援ツール"
- description: "Web ページの要約・翻訳・URL 抽出を行う AI 対応業務効率化アシスタント"
- author: "Development Team"

#### 2. background.js

- PTA → AI 業務支援ツールへのコメント・メッセージ変更
- 設定キー: 'ai_settings' → 'ai_settings'
- 右クリックメニュー刷新:
  - 🤖 選択文を要約・分析
  - 🌐 選択文を翻訳
  - 📄 このページを要約
  - 🌐 このページを翻訳
  - 🔗 URL を抽出してコピー
  - 📋 ページ情報をコピー

#### 3. options.html

- タイトル: "AI 業務支援ツール - 設定"
- ロゴ: 🏫 → 🤖
- 統計ラベル: "メール解析回数" → "テキスト解析回数"
- 開発者: "AI Business Support Team"

#### 4. popup.html

- タイトル・ヘッダー: "AI 業務支援ツール"
- ロゴ: 🏫 → 🤖
- "メール解析" → "コンテンツ解析"
- "返信作成" → "文書作成"
- メール種類に「報告書」「提案書」追加

#### 5. content.js

- コメント・変数名の更新 (PTA → AI 業務支援)
- 関数名リネーム:
  - `addPTAButton()` → `addAISupportButton()`
  - `createPTADialog()` → `createAIDialog()`
- クラス名修正: `pta-*` → `ai-*`
- ダイアログタイトル: 🏫 PTA 支援ツール → 🤖 AI 業務支援ツール
- 右クリックメニューからの新機能ハンドラー追加:
  - `handleSelectionTranslation()` - 選択文翻訳
  - `handlePageTranslation()` - ページ翻訳
  - `handleUrlExtraction()` - URL 抽出・コピー
  - `handlePageInfoCopy()` - ページ情報コピー
- ダイアログの閉じる処理をイベントリスナー方式に修正（右クリックメニューからのダイアログが閉じられない問題を解決）
- 通知機能の追加 (`showNotification()`)

#### 6. content.css

- 全クラス名修正: `.pta-*` → `.ai-*`
- 通知スタイル追加: `.ai-notification`
- アニメーション修正: `pta-slide-in` → `ai-slide-in`
- ダークモード・レスポンシブ対応の更新

#### 7. options.js

- ヘッダーコメント更新
- ストレージキー修正:
  - `ai_settings` → `ai_settings`
  - `ai_history` → `ai_history`
  - `pta_statistics` → `ai_statistics`
- エクスポートファイル名変更:
  - `pta-settings.json` → `ai-settings.json`
  - `pta-history.json` → `ai-history.json`
  - `handlePageTranslation()`
  - `handleUrlExtraction()`
  - `handlePageInfoCopy()`
- ダイアログの閉じる問題修正（イベントリスナー方式採用）
- CSS クラス名更新: pta- → ai-

#### 6. content.css

- クラス名統一: pta- → ai-
- アニメーション名更新: pta-spin → ai-spin

## 🐛 バグ修正

### 右クリックメニューのウィンドウが閉じられない問題

**原因**: `onclick="this.closest('.ai-dialog').remove()"` の使用
**修正**: イベントリスナー方式に変更

```javascript
const closeButton = dialog.querySelector("#ai-close-btn");
closeButton.addEventListener("click", function () {
  dialog.remove();
});
```

### 追加修正

- オーバーレイクリックで閉じる機能
- ESC キーで閉じる機能
- 安全性チェック（null チェック）追加

## ✨ 新機能

### 1. URL 抽出＆コピー（Markdown 形式）

- ページ内のリンクを抽出
- `[リンクテキスト](URL)` 形式でコピー

### 2. ページ情報コピー（Markdown 形式）

```markdown
# ページタイトル

**URL**: https://example.com

## 要約

ページの要約内容...
```

### 3. 翻訳機能

- 選択テキストの翻訳
- ページ全体の翻訳
- 日本語 ⇔ 英語の自動判定

## 🔧 設定ボタンの役割

設定ボタンでは以下の設定が可能:

- **AI プロバイダー**: Azure OpenAI / OpenAI
- **API キー**: セキュアな保存
- **モデル選択**: GPT-4o, GPT-4.1 Mini など
- **動作設定**: 自動検出、通知、履歴保存
- **データ管理**: エクスポート/インポート/削除

## 📊 ページ要約ボタンについて

現在のページ要約ボタンは以下の用途で残しています:

- **手動要約**: ユーザーが任意のタイミングで実行
- **詳細設定**: 要約の詳細度や観点を指定可能
- **履歴管理**: 要約結果を履歴として保存

## 🧪 テスト手順

### 1. 基本動作確認

1. **拡張機能のインストール**

   - Edge で「拡張機能の管理」を開く
   - 「開発者モード」を有効化
   - 「展開して読み込み」で src/edge-extension フォルダを選択

2. **設定確認**
   - ツールバーの 🤖 アイコンをクリック
   - 「設定」ボタンで API キー設定
   - プロバイダー選択とテスト実行

### 2. 新機能テスト

#### 右クリックメニューテスト

1. **文字選択テスト**

   - ページ内でテキストを選択
   - 右クリック → 「🤖 選択文を要約・分析」
   - 右クリック → 「🌐 選択文を翻訳」

2. **ページ全体テスト**
   - 任意のページで右クリック
   - 「📄 このページを要約」
   - 「🌐 このページを翻訳」
   - 「🔗 URL を抽出してコピー」
   - 「📋 ページ情報をコピー」

#### UI テスト

1. **ダイアログ操作**

   - ダイアログが正常に開く
   - × ボタンで閉じる
   - オーバーレイクリックで閉じる
   - ESC キーで閉じる

2. **通知表示**
   - 成功通知の表示
   - エラー通知の表示
   - 自動消去の動作

### 3. バグ修正確認

1. **API キー問題**: 設定画面でキーを保存後、各機能が正常動作することを確認
2. **ダイアログ閉じる問題**: 右クリックメニューからのダイアログが正常に閉じることを確認
3. **変数名エラー**: コンソールにエラーが出ないことを確認

## ✅ 完了状況

- [x] manifest.json のリブランド完了
- [x] UI ファイル（options.html, popup.html）の更新完了
- [x] JavaScript ファイルの PTA → AI 修正完了
- [x] CSS クラス名の統一完了
- [x] ストレージキー名の統一完了
- [x] 新機能（URL 抽出、翻訳、ページ情報コピー）実装完了
- [x] バグ修正（API キー、ダイアログ閉じる、変数名）完了
- [x] エラーチェック完了（JavaScript 構文エラーなし）

## 🎯 今後の改善提案

### 機能拡張

- **要約設定**: 長さ調整、観点指定
- **履歴検索**: キーワード検索、日付フィルター
- **テンプレート機能**: よく使う処理のテンプレート化
- **ショートカットキー**: キーボードショートカット対応

### パフォーマンス最適化

- **API レート制限**: 連続リクエスト制御
- **キャッシュ機能**: 同一ページの重複処理回避
- **バックグラウンド処理**: 重い処理の非同期化

### セキュリティ強化

- **API キー暗号化**: より安全な保存方式
- **データ有効期限**: 履歴の自動削除
- **プライバシー設定**: データ収集の細かい制御

右クリックメニューとの使い分け:

- **右クリック**: クイック操作（即座に実行）
- **ポップアップ**: 詳細設定や履歴確認

## 🎯 今後の改善予定

1. **UI/UX 改善**
   - ダイアログデザインの統一
   - レスポンシブ対応強化
2. **機能拡張**
   - PDF 抽出サポート
   - 画像内テキスト認識
   - 多言語翻訳対応
3. **パフォーマンス**
   - キャッシュ機能
   - バックグラウンド処理最適化

## 🔍 テスト推奨項目

### 基本機能
- [ ] 右クリックメニューからの各機能動作
- [ ] ダイアログの開閉（× ボタン、ESC、オーバーレイクリック）
- [ ] URL 抽出＆コピー機能
- [ ] ページ情報の Markdown 形式コピー
- [ ] 翻訳機能（選択テキスト・ページ全体）
- [ ] 設定画面の動作
- [ ] ポップアップの文書作成機能

### 新機能（M365統合・VSCode解析）
- [ ] **Teams Chat転送**
  - [ ] Microsoft 365ログイン状態での API転送
  - [ ] 認証失敗時のWeb版フォールバック
  - [ ] ページ情報の適切な転送
- [ ] **予定表登録**
  - [ ] 現在日時での予定作成
  - [ ] Outlook Web版フォールバック
  - [ ] イベント詳細情報の正確性
- [ ] **VSCode設定解析**
  - [ ] VSCodeドキュメントページでの機能有効化
  - [ ] 一般ページでの適切なエラーメッセージ
  - [ ] 設定項目抽出の精度
  - [ ] settings.jsonサンプルの生成
  - [ ] コピー機能の動作

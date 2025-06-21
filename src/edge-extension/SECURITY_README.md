# URL 検証セキュリティ対策とテスト手順

## 🔒 セキュリティ脆弱性の修正

### 修正前の問題点

- `String.includes()` メソッドを使用した URL 検証は、悪意のある URL 偽装攻撃に脆弱
- 例：`https://malicious.com/code.visualstudio.com/fake` のような URL が検証を通過してしまう

### 修正後の対策

1. **URL オブジェクトを使用した厳密な検証**

   - `new URL()` コンストラクタで URL を正しく解析
   - `hostname` プロパティで正確なホスト名を取得

2. **許可リストによる完全一致チェック**

   - ホスト名の完全一致チェック
   - パスプレフィックスの厳密な検証

3. **エラーハンドリングの強化**
   - 無効な URL 形式の適切な処理
   - セキュリティログの記録

## 📁 ファイル構成

```
src/edge-extension/
├── infrastructure/
│   └── url-validator.js              # 共通URL検証ユーティリティ
├── background/
│   └── background.js                 # Background Script（修正済み）
├── content/
│   └── content.js                    # Content Script（修正済み）
└── tests/unit/
    └── url-validator-security.test.js # セキュリティテスト
```

## 🛠️ 実装の変更点

### 1. 共通ユーティリティの作成

- `infrastructure/url-validator.js` - 再利用可能な URL 検証ロジック
- セキュアな検証アルゴリズムの統一実装

### 2. Background Script の修正

```javascript
// 修正前（脆弱）
const isVSCodeDoc =
  data.pageUrl &&
  (data.pageUrl.includes("code.visualstudio.com") ||
    data.pageUrl.includes("vscode.docs") ||
    data.pageUrl.includes(
      "docs.microsoft.com/ja-jp/azure/developer/javascript/"
    ) ||
    data.pageUrl.includes("marketplace.visualstudio.com"));

// 修正後（セキュア）
const isVSCodeDoc = isVSCodeDocumentPage(data.pageUrl);
```

### 3. Content Script の修正

```javascript
// 修正前（脆弱）
const isVSCodeDoc =
  dialogData.pageUrl &&
  (dialogData.pageUrl.includes("code.visualstudio.com") ||
    dialogData.pageUrl.includes("vscode") ||
    dialogData.pageUrl.includes("marketplace.visualstudio.com"));

// 修正後（セキュア）
const isVSCodeDoc =
  window.UrlValidator &&
  window.UrlValidator.isVSCodeDocumentPage(dialogData.pageUrl);
```

## 🧪 セキュリティテストの実行

### 1. テスト環境の準備

1. Chrome DevTools を開く
2. 拡張機能の Background Script コンソールを開く
3. 以下のファイルを読み込む：
   - `infrastructure/url-validator.js`
   - `tests/unit/url-validator-security.test.js`

### 2. テストの実行

```javascript
// セキュリティテストの実行
const testResults = window.SecurityTester.runFullSecurityTest(
  window.UrlValidator.isVSCodeDocumentPage
);
```

### 3. テストケース

テストには以下の脅威シナリオが含まれています：

#### 悪意のある URL 偽装攻撃

- `https://malicious.com/code.visualstudio.com/fake`
- `https://code.visualstudio.com.malicious.com/`
- `https://fake-marketplace.visualstudio.com/`
- `https://code-visualstudio.com/`
- `https://evil.com?redirect=code.visualstudio.com`

#### エッジケース

- 無効な URL 形式
- null/undefined 値
- 異なるプロトコル
- 大文字・小文字の混在

## 📊 期待されるテスト結果

### 成功した場合

```
🎉 すべてのセキュリティテストが成功しました！
総テスト数: 18
成功: 18
失敗: 0
成功率: 100.0%
```

### 失敗した場合

失敗したテストがある場合は、セキュリティ脆弱性が存在する可能性があります。

## 🔧 トラブルシューティング

### 1. manifest.json の設定確認

```json
"content_scripts": [
  {
    "js": [
      "infrastructure/url-validator.js",
      "content/content.js"
    ]
  }
]
```

### 2. Background Script での使用

```javascript
// 直接関数を使用
const isValid = isVSCodeDocumentPage(url);
```

### 3. Content Script での使用

```javascript
// window.UrlValidator経由で使用
const isValid =
  window.UrlValidator && window.UrlValidator.isVSCodeDocumentPage(url);
```

## 🛡️ セキュリティベストプラクティス

### 1. URL 検証の原則

- **完全一致チェック**: ホスト名は必ず完全一致で検証
- **プロトコル検証**: 必要に応じて HTTPS/HTTP プロトコルを限定
- **パス検証**: パスプレフィックスの厳密なチェック

### 2. 入力値検証

- **null/undefined チェック**: 入力値の存在確認
- **型チェック**: 文字列型の確認
- **長さ制限**: 異常に長い URL の拒否

### 3. エラーハンドリング

- **例外捕捉**: URL 解析エラーの適切な処理
- **ログ記録**: セキュリティ関連の操作ログ
- **フォールバック**: 検証エラー時は安全側に判定

## 📝 継続的なセキュリティ監視

### 1. 定期的なテスト実行

- 新機能追加時のセキュリティテスト実行
- 定期的な脆弱性スキャン

### 2. 脅威情報の更新

- 新しい攻撃手法の調査
- 許可ドメインリストの定期見直し

### 3. CodeQL 等の静的解析ツール活用

- CI/CD パイプラインでの自動セキュリティチェック
- セキュリティ脆弱性の早期発見

## 🚀 デプロイ前チェックリスト

- [ ] セキュリティテストの実行と全テスト成功
- [ ] CodeQL 警告の解消
- [ ] URL 許可リストの妥当性確認
- [ ] エラーハンドリングの動作確認
- [ ] ログ出力の適切性確認

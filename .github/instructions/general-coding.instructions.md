---
applyTo: "**"
---
# プロジェクト一般コーディング基準

## 命名規則
- クラス名やインターフェイス名には PascalCase を使用
- 変数名、関数名、メソッド名には camelCase を使用
- プライベートなクラスメンバーには先頭にアンダースコア (_) を付ける
- 定数には ALL_CAPS を使用

## エラー処理
- 非同期操作には try/catch ブロックを使用
- エラーログには常に文脈情報を含める
- エラーは適切に処理し、ユーザーにわかりやすいメッセージを表示する

## コードスタイル
- コードには適切なコメントを日本語または英語で記述する
- 1つの関数やメソッドは1つの責任に集中する
- 再利用可能なコードはモジュール化する

## コード品質管理
- **未使用変数・インポートの除去**: 使用されていない変数、関数、インポートは削除する
- **未使用パラメータの処理**: 必要に応じて `_` プレフィックスを付けるか、削除を検討する
- **デッドコードの除去**: 到達不可能なコードや使用されない関数は削除する
- **一時的なデバッグコード**: 本番環境に不要なconsole.log()等は削除する

```javascript
// ❌ 悪い例：未使用変数
import { API_CLIENT, UNUSED_UTIL } from './utils.js'; // UNUSED_UTILが未使用
const unusedVariable = 'not used';

function processData(data, unusedParam) { // unusedParamが未使用
    return API_CLIENT.process(data);
}

// ✅ 良い例：クリーンなコード
import { API_CLIENT } from './utils.js';

function processData(data) {
    return API_CLIENT.process(data);
}

// ✅ 意図的に未使用の場合はプレフィックスを使用
function eventHandler(event, _unusedContext) {
    // _unusedContextは意図的に未使用
    return event.preventDefault();
}
```

## セキュリティ対策
- 機密情報はハードコーディングせず、環境変数または設定ファイルから読み込む
- ユーザー入力は常に検証し、インジェクション攻撃を防止する
- API利用時にはレート制限を考慮する

## 【重要】HTML/テキストサニタイゼーション要件

### 不完全サニタイゼーション問題の防止
以下の実装は**セキュリティ脆弱性の原因**となるため禁止：

```javascript
// ❌ 危険：不完全な正規表現サニタイゼーション
.replace(/<[^>]*>/g, '')           // HTMLタグ除去（バイパス可能）
.replace(/javascript:/gi, '')      // プロトコル除去（不完全）
.replace('{{placeholder}}', input) // 直接置換（エスケープなし）
```

### 必須：統一セキュリティサニタイザーの使用
```javascript
// ✅ 安全：DOMPurifyベースの統一サニタイザー
const sanitizer = getSecuritySanitizer();
const safeHTML = sanitizer.sanitizeHTML(input);
const plainText = sanitizer.extractPlainText(input);
const securePrompt = sanitizer.buildSecurePrompt(template, variables);
```

### コード品質チェック
- 全ての外部入力は統一サニタイザー経由で処理すること
- 正規表現による直接HTML処理は禁止
- プロンプトテンプレート構築では `buildSecurePrompt()` を使用
- "incomplete multi-character sanitization" 警告の回避を徹底

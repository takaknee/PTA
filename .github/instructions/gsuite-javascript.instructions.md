---
applyTo: "**/*.gs,**/*.js"
---
# G-suits (Google Apps Script) コーディング基準

[一般コーディングガイドライン](./.github/instructions/general-coding.instructions.md)をすべてのコードに適用してください。

## Google Apps Script ガイドライン
- Google サービス（SpreadsheetApp, DriveApp, GmailApp など）を適切に活用する
- スクリプトの実行時間制限（6分）を考慮し、長時間実行が必要な場合は分割実行を検討する
- トリガーを使用する場合は、重複実行を防止する仕組みを実装する
- ユーザー認証が必要な場合は、OAuth2 を適切に実装する

## JavaScript ガイドライン
- ES6+ の機能を積極的に活用する（Google Apps Script の対応状況を確認）
- 関数は小さく保ち、一つの機能に集中させる
- 変数のスコープを最小限に保つ
- グローバル変数の使用を避け、必要な場合は明確に命名する

## ログ記録
- Logger.log() を適切に使用し、実行状況を追跡可能にする
- 本番環境では console.log() ではなく適切なログ記録メカニズムを使用する
- エラー発生時には詳細情報をログに記録する

## 【重要】セキュリティ・サニタイゼーション要件

### 不完全多文字サニタイゼーション問題の回避
以下の実装パターンは**絶対に禁止**（セキュリティ脆弱性の原因）：

```javascript
// ❌ 禁止パターン：不完全な正規表現サニタイゼーション
text.replace(/<[^>]*>/g, '')       // HTMLタグ除去（不完全）
text.replace(/&[^;]*;/g, ' ')      // HTMLエンティティ処理（不完全）
text.replace(/javascript:/gi, '')  // プロトコル除去（不完全）
text.replace('{{var}}', userInput) // 直接置換（エスケープなし）

// ❌ テンプレート文字列での直接挿入
`テンプレート: ${userInput}`       // エスケープなし挿入
```

### 【重要】DOMPurify必須アーキテクチャ実装

```javascript
// ✅ 推奨：DOMPurify必須の統一サニタイザー使用
const sanitizer = globalThis.PTASanitizer; // DOMPurify必須

// 必須：DOMPurify利用可能性チェック
if (!sanitizer.isDOMPurifyAvailable()) {
    throw new Error('DOMPurifyが必要です。ライブラリを読み込んでください。');
}

// HTMLサニタイゼーション（DOMPurifyベース）
const cleanHTML = sanitizer.sanitizeHTML(htmlContent);

// プレーンテキスト抽出
const plainText = sanitizer.extractSafeText(htmlContent);

// 高速タグ除去
const fastStrip = sanitizer.fastStripTags(htmlContent);
```

### コード生成時の必須チェック項目（更新版）
1. **全ての外部入力**（ユーザー入力、API応答、ページコンテンツ）は統一サニタイザー経由で処理
2. **DOMPurifyが必須依存関係として含まれていること**
3. **正規表現による直接HTML処理は絶対禁止**
4. **フォールバック実装は作成しないこと**（セキュリティリスク）
5. **DOMPurify利用不可時は適切にエラーハンドリングすること**
6. **未使用変数・インポートの完全除去**
7. **デッドコード・一時的デバッグコードの削除**

### 未使用変数・インポート管理
```javascript
// ❌ 禁止：未使用インポート
import { USED_CONST, UNUSED_CONST } from './constants.js';

// ❌ 禁止：未使用変数
const unusedVariable = 'value';

// ❌ 禁止：未使用パラメータ（意図的でない場合）
function handler(event, unusedParam) {
    return event.target.value;
}

// ✅ 推奨：必要なもののみインポート
import { USED_CONST } from './constants.js';

// ✅ 推奨：意図的未使用は_プレフィックス
function handler(event, _metadata) {
    // _metadataは意図的に未使用（API仕様上必要）
    return event.target.value;
}
```

### VSCodeエラー対応（DOMPurify必須版）
"incomplete multi-character sanitization" 警告の根本的解決：
- **DOMPurifyを必須として使用**し、複雑な正規表現処理を排除
- **フォールバック実装は作成禁止**（セキュリティリスク）
- **文字列置換処理はDOMPurify処理後に実行**
- **DOMPurify利用不可時は安全にエラーを投げる**

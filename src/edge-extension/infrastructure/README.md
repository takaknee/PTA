# HTML サニタイゼーション共通モジュール

## 概要

PTA Edge Extension プロジェクトにおいて、HTML サニタイゼーション処理を共通化・モジュール化したライブラリです。
CodeQL セキュリティスキャンで指摘された「Incomplete multi-character sanitization」脆弱性に対応し、
セキュリティを強化した統一的なサニタイゼーション機能を提供します。

## 🔒 セキュリティ機能

### 主要な保護機能

- **XSS 攻撃防止**: 危険な HTML タグと属性の除去
- **HTML インジェクション防止**: コメント、CDATA、処理命令の安全な除去
- **JavaScript 実行防止**: イベントハンドラーとプロトコルの無効化
- **コード実行防止**: Script、Style、Object タグの完全除去

### CodeQL 脆弱性対応

- **段階的サニタイゼーション**: 複数文字パターンを単一文字レベルで処理
- **無限ループ防止**: 最大反復回数制限の実装
- **多層防御**: 従来の正規表現 + 安全な代替方式
- **包括的エラーハンドリング**: 予期しないエラーに対する安全な回復

## 📁 ファイル構成

```
src/edge-extension/infrastructure/
└── html-sanitizer.js          # HTMLサニタイゼーション共通モジュール
```

## 🛠️ 使用方法

### ES Modules（推奨）

```javascript
// 動的インポート
import sanitizer from "./infrastructure/html-sanitizer.js";

// 使用例
const safeText = sanitizer.extractSafeText(htmlContent);
```

### Service Worker 環境

```javascript
// 動的インポートでの初期化
async function initializeSanitizer() {
  const sanitizerModule = await import("./infrastructure/html-sanitizer.js");
  return sanitizerModule.default;
}

// 使用例
const sanitizer = await initializeSanitizer();
const safeText = sanitizer.extractSafeText(htmlContent);
```

### レガシー環境

```javascript
// グローバルオブジェクトからアクセス
const safeText = globalThis.PTASanitizer.extractSafeText(htmlContent);
```

## 🚀 API リファレンス

### メイン関数

#### `extractSafeText(html)`

HTML コンテンツから安全なテキストを抽出する包括的な関数

```javascript
const safeText = sanitizer.extractSafeText(
  '<p onclick="alert(1)">Hello <script>alert(2)</script>World</p>'
);
// 結果: "Hello World"
```

#### `stripHTMLTags(html)`

基本的な HTML タグ除去（軽量版）

```javascript
const plainText = sanitizer.stripHTMLTags("<p>Hello <b>World</b></p>");
// 結果: "Hello World"
```

### 特殊化された関数

#### `sanitizeMultiCharacterPatterns(input)`

複数文字パターンの安全な除去（CodeQL 対応）

#### `removeDangerousTags(html)`

危険な HTML タグの除去

#### `removeDangerousAttributes(html)`

危険な属性とプロトコルの除去

#### `decodeHTMLEntities(text)`

HTML エンティティの安全なデコード

#### `normalizeText(text)`

テキストの正規化（空白、制御文字の処理）

### テスト機能

#### `testSanitization()`

サニタイゼーション機能のテスト実行

```javascript
const testResult = sanitizer.testSanitization();
// コンソールにテスト結果を出力し、成功/失敗を返す
```

## 🧪 テストケース

### 対応する攻撃パターン

1. **基本的な HTML コメント**: `<!-- comment -->text`
2. **ネストされた HTML コメント攻撃**: `<!<!--- comment --->>text`
3. **CDATA 攻撃**: `<![CDATA[<script>alert(1)</script>]]>text`
4. **処理命令攻撃**: `<?php echo "test"; ?>text`
5. **複合攻撃**: `<!--<![CDATA[--><script>alert(1)</script><!--]]>-->text`
6. **JavaScript プロトコル**: `<a href="javascript:alert(1)">text</a>`
7. **イベントハンドラー属性**: `<div onclick="alert(1)">text</div>`
8. **HTML エンティティ**: `&lt;script&gt;alert(1)&lt;/script&gt;`

## 🔧 設定とカスタマイズ

### 危険なタグのカスタマイズ

```javascript
// デフォルトで除去される危険なタグ
const DANGEROUS_TAGS = [
  "script",
  "style",
  "iframe",
  "frame",
  "frameset",
  "noframes",
  "object",
  "embed",
  "applet",
  "meta",
  "link",
  "base",
  "form",
  "input",
  "button",
  "select",
  "textarea",
  "option",
  "img",
  "audio",
  "video",
  "source",
  "track",
  "svg",
  "math",
  "canvas",
];
```

### 危険な属性のカスタマイズ

```javascript
// デフォルトで除去される危険な属性
const DANGEROUS_ATTRIBUTES = [
  "onabort",
  "onblur",
  "onchange",
  "onclick" /* ... */,
  "javascript:",
  "vbscript:",
  "data:",
  "blob:",
  "file:",
];
```

## ⚡ パフォーマンス

### 軽量版と完全版の使い分け

- **軽量版** (`stripHTMLTags`): 基本的な HTML タグ除去のみ
- **完全版** (`extractSafeText`): 包括的なセキュリティ処理

### 処理時間の目安

- 小さなテキスト（～ 1KB）: < 1ms
- 中程度の HTML（～ 10KB）: < 10ms
- 大きな HTML（～ 100KB）: < 100ms

## 🛡️ セキュリティ考慮事項

### 制限事項

- **長さ制限**: デフォルトで 10,000 文字まで（設定可能）
- **ネスト制限**: 最大 100 回の反復処理
- **メモリ制限**: 大量の HTML コンテンツには注意

### 推奨事項

1. **入力検証**: サニタイゼーション前の入力サイズ確認
2. **エラーハンドリング**: try-catch ブロックでの包含
3. **ログ記録**: セキュリティイベントの記録
4. **定期更新**: 新しい攻撃パターンへの対応

## 🔄 移行ガイド

### 既存コードからの移行

```javascript
// 旧: インライン処理
function oldSanitize(html) {
  return html.replace(/<[^>]*>/g, "");
}

// 新: 共通モジュール使用
import sanitizer from "./infrastructure/html-sanitizer.js";
function newSanitize(html) {
  return sanitizer.extractSafeText(html);
}
```

### 段階的移行手順

1. **共通モジュールのインストール**
2. **既存のサニタイゼーション関数の特定**
3. **段階的な置き換え実行**
4. **テスト実行と検証**
5. **古いコードの削除**

## 📊 メトリクスと監視

### 開発時テスト

```javascript
// 開発環境でのテスト実行
if (manifest.name.includes("Dev")) {
  sanitizer.testSanitization();
}
```

### 本番監視

```javascript
// エラー率の監視
try {
  const result = sanitizer.extractSafeText(html);
} catch (error) {
  console.error("サニタイゼーション失敗:", error);
  // エラー報告システムに送信
}
```

## 🤝 貢献とメンテナンス

### コードレビューポイント

- セキュリティ要件の遵守
- パフォーマンスへの影響
- 既存機能への互換性
- テストケースの充実

### 更新頻度

- **セキュリティ修正**: 即座
- **機能追加**: 月次
- **パフォーマンス改善**: 四半期
- **メジャーアップデート**: 年次

## 📝 ライセンス

このモジュールは PTA プロジェクトの一部として開発され、
プロジェクトの共通ライセンス下で提供されます。

## 🔗 関連リンク

- [OWASP XSS Prevention Cheat Sheet](https://owasp.org/www-project-cheat-sheets/cheatsheets/Cross_Site_Scripting_Prevention_Cheat_Sheet.html)
- [CodeQL Security Queries](https://codeql.github.com/codeql-query-help/javascript/)
- [Chrome Extension Security](https://developer.chrome.com/docs/extensions/mv3/security/)

---

**作成日**: 2025 年 6 月 18 日  
**最終更新**: 2025 年 6 月 18 日  
**担当**: Development Team

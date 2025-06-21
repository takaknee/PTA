# セキュリティ修正レポート: Bad HTML filtering regexp

## 修正概要

CodeQL セキュリティ検査で検出された「Bad HTML filtering regexp」脆弱性を修正しました。

### 問題の詳細

**Rule ID**: `js/bad-tag-filter`  
**重要度**: セキュリティ重要  
**検出場所**: `src/edge-extension/content/content.js:1091`

**問題のあった正規表現**:
```javascript
// ❌ 脆弱性のあるパターン
{ name: 'script タグ', pattern: /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi }
```

**問題点**:
- `</script >`のような終了タグ（スペースを含む）をマッチできない
- 悪意のあるスクリプトがフィルターを回避する可能性
- XSS攻撃やコード実行のリスクが存在

## 修正内容

### 1. 正規表現パターンの強化

**修正前**:
```javascript
/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi
```

**修正後**:
```javascript
// 開始タグと終了タグを個別に検知
{ name: 'script 開始タグ', pattern: /<script[\s\S]*?>/gi, severity: 'critical' }
{ name: 'script 終了タグ', pattern: /<\/script[\s]*>/gi, severity: 'critical' }
```

### 2. セキュリティ検知機能の包括的強化

#### 新しいセキュリティパターン検知

```javascript
// より堅牢で包括的なパターン検知
const securityPatterns = [
  // JavaScript実行関連
  { name: 'script 開始タグ検知', pattern: /<script(?:\s+[^>]*)?>/gi, severity: 'critical' },
  { name: 'script 終了タグ検知', pattern: /<\/script(?:\s*)>/gi, severity: 'critical' },
  { name: 'JavaScript プロトコル', pattern: /(?:javascript|vbscript|data|blob|file)\s*:/gi, severity: 'high' },
  
  // イベントハンドラー
  { name: 'イベントハンドラー', pattern: /\bon\w+\s*=(?:\s*["']?[^"'>]*["']?|[^>\s]*)/gi, severity: 'high' },
  
  // 外部コンテンツ埋め込み
  { name: 'iframe 開始タグ', pattern: /<iframe(?:\s+[^>]*)?>/gi, severity: 'high' },
  { name: 'iframe 終了タグ', pattern: /<\/iframe(?:\s*)>/gi, severity: 'high' },
  
  // その他の危険なタグ
  { name: 'object/embed タグ', pattern: /<(?:object|embed)(?:\s+[^>]*)?>/gi, severity: 'medium' },
  { name: 'form/input タグ', pattern: /<(?:form|input|button|select|textarea)(?:\s+[^>]*)?>/gi, severity: 'medium' },
  
  // 高度な回避パターン
  { name: 'HTMLエンコード回避', pattern: /&(?:#(?:x[0-9a-f]+|[0-9]+)|[a-z]+);/gi, severity: 'low' },
  { name: 'Base64疑似パターン', pattern: /(?:data:[\w\/]+;base64,|btoa|atob)[A-Za-z0-9+\/=]{20,}/gi, severity: 'medium' }
];
```

#### リスクレベル分類システム

- **Critical**: JavaScript実行が可能な要素
- **High**: 外部リソース読み込みやイベント実行
- **Medium**: プラグイン実行やフォーム要素
- **Low**: 監視対象だが直接的な脅威ではない

### 3. エラーハンドリングと監査ログの強化

#### セキュリティ警告の段階的表示

```javascript
function showSecurityWarning(detectedPatterns, riskLevel) {
  const riskConfig = {
    critical: { icon: '🚨', color: '#ff4444', message: 'クリティカルリスク' },
    high: { icon: '⚠️', color: '#ff9800', message: '高リスク' },
    medium: { icon: '📝', color: '#ffc107', message: '中リスク' },
    low: { icon: '📝', color: '#2196f3', message: '低リスク' }
  };
  
  // リスクレベルに応じた適切な警告表示
  // ユーザー通知、コンソールログ、監査ログを段階的に実装
}
```

#### 包括的な監査ログ

```javascript
const auditLog = {
  timestamp: new Date().toISOString(),
  riskLevel: riskLevel,
  patternsDetected: detectedPatterns.length,
  summary: {
    criticalCount: detectedPatterns.filter(p => p.severity === 'critical').length,
    highCount: detectedPatterns.filter(p => p.severity === 'high').length,
    // ...その他の統計情報
  },
  details: detectedPatterns.map(p => ({
    name: p.name,
    count: p.count,
    severity: p.severity,
    description: p.description
  }))
};
```

### 4. HTMLサニタイザーモジュールの強化

#### 共通ライブラリでの高度検知機能追加

`src/edge-extension/infrastructure/html-sanitizer.js` に以下を追加:

- `detectAdvancedSecurityPatterns()` 関数
- より堅牢な正規表現パターン
- Service Worker 対応の実装
- 包括的なエラーハンドリング

#### モジュール間の連携強化

- Content Script での高度検知機能の活用
- フォールバック機構の実装
- エラー時の安全な処理

## セキュリティ効果

### Before (修正前)
- `</script >`のような終了タグが検出されない
- 単純な正規表現による限定的な検知
- セキュリティリスクの段階的評価なし

### After (修正後)
- ✅ スペースを含む終了タグも確実に検知
- ✅ 開始・終了タグの個別検知でより確実
- ✅ クリティカル〜低リスクまでの段階的評価
- ✅ 包括的な危険パターンの検知
- ✅ 詳細な監査ログとユーザー警告
- ✅ モジュール化による保守性向上

## テスト検証

### 検証対象パターン

1. **基本的なscriptタグ**
   ```html
   <script>alert('test')</script>
   <script type="text/javascript">alert('test')</script>
   ```

2. **スペースを含む終了タグ**
   ```html
   <script>alert('test')</script >
   <script>alert('test')</script   >
   ```

3. **複雑な属性を持つタグ**
   ```html
   <script src="malicious.js" defer async>
   <iframe src="data:text/html,<script>alert('xss')</script>">
   ```

4. **イベントハンドラー**
   ```html
   <img onerror="alert('xss')" src="invalid">
   <div onclick="maliciousFunction()">
   ```

### 検証結果

すべてのパターンが適切に検知され、リスクレベルに応じた警告が表示されることを確認。

## 今後の改善予定

1. **DOMPurify などの専門ライブラリとの統合検討**
2. **CSP (Content Security Policy) との連携強化**
3. **機械学習による異常検知の導入**
4. **リアルタイムセキュリティ監視の実装**

## 関連ドキュメント

- [OWASP XSS Prevention Cheat Sheet](https://owasp.org/www-project-cheat-sheets/cheatsheets/Cross_Site_Scripting_Prevention_Cheat_Sheet.html)
- [Chrome Extension Security](https://developer.chrome.com/docs/extensions/mv3/security/)
- [CodeQL JavaScript Security Queries](https://codeql.github.com/codeql-query-help/javascript/)

---

**修正日時**: 2025年6月21日  
**修正者**: Development Team  
**レビュー**: セキュリティチーム承認済み

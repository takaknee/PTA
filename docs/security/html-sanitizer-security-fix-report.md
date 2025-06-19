# HTML サニタイザー セキュリティ修正レポート

## 修正概要

**日付**: 2025 年 6 月 19 日  
**対象ファイル**: `src/edge-extension/infrastructure/html-sanitizer.js`  
**脆弱性 ID**: CodeQL js/bad-tag-filter  
**重要度**: 高

## 問題の詳細

### 発見された脆弱性

CodeQL スキャンにより、HTML コメント処理において不正な形式のコメント終了タグを適切に処理できていない問題が発見されました。

**問題のあったコード**:

```javascript
// 109行目付近
result = result.replace(/<!--/g, "").replace(/-->/g, "");
```

### セキュリティリスク

- `--!>` のような不正な HTML コメント終了タグが適切に処理されない
- XSS 攻撃やコード実行攻撃のバイパスに利用される可能性
- HTML 構造の破壊によるレンダリング問題

## 実装した修正

### 1. 専用の安全な処理関数を作成

新しい関数 `removeHTMLCommentsSafely()` を実装:

```javascript
function removeHTMLCommentsSafely(input) {
  // 複数の不正パターンに対応した安全な処理
  // 1. 正規のHTMLコメント (<!--...-->)
  // 2. 不正なコメント開始タグ (<!--)
  // 3. 不正なコメント終了タグ (--!>, -!>, --a> など)
  // 4. 不完全なコメント構造
}
```

### 2. 既存サニタイジング関数との統合

`sanitizeMultiCharacterPatterns()` 関数内の HTML コメント処理部分を、新しい安全な関数を使用するように修正:

```javascript
{
  name: 'HTMLコメント',
  regex: /<!--[\s\S]*?-->/g,
  safeAlternative: (str) => removeHTMLCommentsSafely(str)
}
```

### 3. 包括的なテストケースの追加

新しいテストファイル `html-sanitizer-security.test.js` を作成:

- 正規の HTML コメント処理テスト
- 不正なコメント終了タグのテスト
- XSS 攻撃パターンのテスト
- パフォーマンステスト
- 回帰テスト

## 修正後のセキュリティ強化点

### 1. 多様な不正パターンへの対応

- `--!>`, `-!>`, `--a>` などの変形パターンに対応
- 不完全なコメント構造の安全な処理
- ネストしたコメント攻撃への対応

### 2. 無限ループ防止

- 最大反復回数の制限（100 回）
- 処理の変化検出による効率的な終了判定

### 3. 堅牢なエラーハンドリング

- 入力値の型チェック
- 例外処理の充実
- ログ出力による可視性向上

## テスト結果

### セキュリティテスト

以下の攻撃パターンに対して適切に防御することを確認:

1. **不正コメント終了タグ**: `<!--攻撃--!>` → 完全除去
2. **XSS コメント攻撃**: `<!--<script>alert(1)--!>` → 完全除去
3. **複合攻撃パターン**: 複数の不正パターンの組み合わせ → 安全に処理

### パフォーマンステスト

- 大量データ処理: 10KB のコメントデータを適切に処理
- 大量タグ処理: 1000 個の不正コメントタグを効率的に除去

## 影響範囲

### 修正対象

- HTML サニタイザーのコメント処理ロジック
- セキュリティテストスイート

### 既存機能への影響

- **破壊的変更なし**: 既存の API 仕様は維持
- **パフォーマンス向上**: より効率的な処理アルゴリズム
- **セキュリティ強化**: より包括的な脅威への対応

## 今後の推奨事項

### 1. 定期的なセキュリティスキャン

- CodeQL スキャンの定期実行
- 新しい脆弱性パターンの監視
- セキュリティテストの継続的な更新

### 2. 他のサニタイゼーション機能の点検

- 他の HTML 要素処理の安全性確認
- 属性処理ロジックの見直し
- CSS サニタイゼーションの強化

### 3. ドキュメントの充実

- セキュリティガイドラインの整備
- 開発者向けのベストプラクティス文書
- インシデント対応手順の策定

## 関連リソース

- [OWASP HTML Sanitization Cheat Sheet](https://cheatsheetseries.owasp.org/cheatsheets/Cross_Site_Scripting_Prevention_Cheat_Sheet.html)
- [CodeQL Query Reference](https://codeql.github.com/codeql-query-help/javascript/)
- [Chrome Extension Security Best Practices](https://developer.chrome.com/docs/extensions/mv3/security/)

---

**修正担当**: PTA プロジェクト開発チーム  
**レビュー**: セキュリティチーム（要実施）

# DOMPurify組み込み手順

## Edge拡張でDOMPurifyを使用する場合のmanifest.json更新例

```json
{
  "manifest_version": 3,
  "name": "AI業務支援ツール",
  "version": "1.0.1",
  "description": "Webページの要約・翻訳・URL抽出を行うAI対応業務効率化アシスタント（DOMPurify強化版）",
  
  "content_scripts": [
    {
      "matches": [
        "https://*/*",
        "http://*/*"
      ],
      "js": [
        "vendor/dompurify.min.js",           // ← DOMPurifyを最初に読み込み
        "vendor/secure-sanitizer.js",        // ← セキュリティ強化スクリプト
        "infrastructure/html-sanitizer.js",  // ← 既存のHTMLサニタイザー
        "infrastructure/url-validator.js",
        "content/content.js"
      ],
      "css": [
        "content/content.css"
      ],
      "run_at": "document_idle"
    }
  ],
  
  "web_accessible_resources": [
    {
      "resources": [
        "assets/*",
        "offscreen/*",
        "vendor/*",                          // ← ベンダーライブラリも公開
        "infrastructure/html-sanitizer.js"
      ],
      "matches": [
        "<all_urls>"
      ]
    }
  ]
}
```

## 導入手順

### 1. DOMPurifyのインストール
```bash
cd src/edge-extension
npm install
npm run install-dompurify
```

### 2. ビルド実行
```bash
npm run build
```

### 3. 拡張機能の再読み込み
- Edge の拡張機能ページで「再読み込み」
- 開発者ツールでDOMPurifyが読み込まれていることを確認

### 4. 動作確認
```javascript
// コンソールで確認
console.log('DOMPurify available:', typeof DOMPurify !== 'undefined');
console.log('Sanitizer function:', window.activeSanitizer.name);
```

## パフォーマンス比較

### ファイルサイズ
- **カスタム実装のみ**: ~50KB
- **DOMPurify組み込み**: ~95KB (+45KB)

### セキュリティレベル
- **カスタム実装**: CodeQL脆弱性修正済み、包括的検知
- **DOMPurify版**: 業界標準 + カスタム検知のハイブリッド

## 今後の展開

1. **段階的導入**
   - 開発環境でのテスト実装
   - 本番環境での限定的展開
   - 全面移行

2. **ハイブリッド戦略**
   - DOMPurifyによる主要サニタイゼーション
   - カスタム検知による追加監視
   - パフォーマンス最適化

3. **継続的な改善**
   - セキュリティアップデートの自動化
   - パフォーマンス監視
   - ユーザーフィードバックの反映

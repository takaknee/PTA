# HTMLサニタイゼーションライブラリ検討レポート

## CodeQL推奨事項と Edge拡張の制約

### CodeQLの推奨事項
> **Recommendation**: Use a well-tested sanitization or parser library if at all possible. These libraries are much more likely to handle corner cases correctly than a custom implementation.

### Edge拡張機能の制約

1. **Manifest V3セキュリティ制約**
   - 外部CDNからの読み込み禁止
   - Content Security Policy (CSP) の厳格な制限
   - インラインスクリプト禁止

2. **バンドルサイズ制約**
   - 拡張機能全体の推奨サイズ: 100KB以下
   - ダウンロード時間とメモリ使用量への影響

3. **Service Worker制限**
   - ES Modules 使用不可
   - DOM API の制限
   - `importScripts()` のみサポート

## 主要ライブラリの評価

### 1. DOMPurify 🌟 推奨
```javascript
// サイズ: ~45KB (minified)
// 特徴: 最も信頼性が高いHTMLサニタイゼーションライブラリ
// GitHub: https://github.com/cure53/DOMPurify

// Edge拡張での使用方法
import DOMPurify from 'dompurify';

function sanitizeWithDOMPurify(html) {
    return DOMPurify.sanitize(html, {
        ALLOWED_TAGS: ['b', 'i', 'em', 'strong', 'p', 'br', 'ul', 'ol', 'li'],
        ALLOWED_ATTR: ['class']
    });
}
```

**メリット**:
- 業界標準のセキュリティライブラリ
- 包括的なXSS防止
- 豊富な設定オプション
- 定期的なセキュリティ更新

**デメリット**:
- ファイルサイズが大きい（45KB）
- Manifest V3での組み込みが複雑

### 2. sanitize-html
```javascript
// サイズ: ~80KB+ (依存関係含む)
// Node.js中心のライブラリ

const sanitizeHtml = require('sanitize-html');

function sanitizeContent(html) {
    return sanitizeHtml(html, {
        allowedTags: ['b', 'i', 'em', 'strong', 'p'],
        allowedAttributes: {}
    });
}
```

**メリット**:
- 高い柔軟性
- 詳細な設定が可能

**デメリット**:
- ファイルサイズが大きい
- ブラウザ環境での使用が複雑
- Edge拡張には不適

### 3. js-xss
```javascript
// サイズ: ~25KB
// 軽量なXSS防止ライブラリ

import xss from 'xss';

function sanitizeWithXSS(html) {
    return xss(html, {
        whiteList: {
            p: [],
            b: [],
            i: [],
            strong: [],
            em: []
        }
    });
}
```

**メリット**:
- 比較的軽量
- シンプルなAPI

**デメリット**:
- DOMPurifyより機能が限定的
- 設定の柔軟性が低い

## Edge拡張でのライブラリ組み込み方法

### 方法1: ダイレクトバンドリング (推奨)

```bash
# 1. ライブラリのダウンロード
npm install dompurify

# 2. ビルドプロセスでバンドル
# webpack.config.js
module.exports = {
    entry: './src/content/content.js',
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: 'content.bundle.js'
    },
    resolve: {
        fallback: {
            "buffer": false,
            "crypto": false,
            "stream": false
        }
    }
};
```

**ディレクトリ構造**:
```
src/edge-extension/
├── vendor/
│   └── dompurify.min.js    # 外部ライブラリを直接配置
├── content/
│   └── content.js
└── manifest.json
```

**manifest.jsonでの読み込み**:
```json
{
    "content_scripts": [
        {
            "matches": ["<all_urls>"],
            "js": [
                "vendor/dompurify.min.js",
                "content/content.js"
            ]
        }
    ]
}
```

### 方法2: CDN版の手動ダウンロード

```bash
# DOMPurifyの最新版をダウンロード
wget https://cdn.jsdelivr.net/npm/dompurify@3.0.5/dist/purify.min.js \
     -O src/edge-extension/vendor/dompurify.min.js
```

### 方法3: ESビルドでのバンドリング

```javascript
// build.js
const esbuild = require('esbuild');

esbuild.build({
    entryPoints: ['src/content/content.js'],
    bundle: true,
    outfile: 'dist/content.bundle.js',
    platform: 'browser',
    format: 'iife',
    external: []
}).catch(() => process.exit(1));
```

## 実装計画: DOMPurifyの組み込み

### ステップ1: DOMPurifyの導入

```javascript
// src/edge-extension/content/content-with-dompurify.js

// DOMPurifyを使用したセキュアなサニタイゼーション
function sanitizeHtmlResponseSecure(html) {
    if (!html || typeof html !== 'string') {
        return '';
    }

    // DOMPurifyを使用した安全なサニタイゼーション
    const cleanHtml = DOMPurify.sanitize(html, {
        // AI応答表示用の安全な設定
        ALLOWED_TAGS: [
            'p', 'br', 'div', 'span', 'strong', 'b', 'em', 'i', 'u',
            'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
            'ul', 'ol', 'li',
            'blockquote', 'pre', 'code',
            'table', 'thead', 'tbody', 'tr', 'td', 'th'
        ],
        ALLOWED_ATTR: [
            'class', 'style'
        ],
        ALLOWED_STYLES: {
            'color': true,
            'background-color': true,
            'font-weight': true,
            'font-style': true,
            'text-decoration': true
        },
        // 危険な要素を完全に除去
        FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'form', 'input'],
        FORBID_ATTR: ['onclick', 'onerror', 'onload', 'href', 'src'],
        // セキュリティ強化オプション
        SAFE_FOR_TEMPLATES: true,
        WHOLE_DOCUMENT: false,
        RETURN_DOM: false,
        RETURN_DOM_FRAGMENT: false,
        RETURN_TRUSTED_TYPE: false
    });

    // 追加のセキュリティチェック（既存の検知機能と併用）
    const securityResult = detectSuspiciousPatterns(cleanHtml);
    if (securityResult.riskLevel === 'high' || securityResult.riskLevel === 'critical') {
        console.warn('🔒 DOMPurify処理後もリスクパターンが検出されました:', securityResult);
    }

    return cleanHtml;
}
```

### ステップ2: フォールバック機構

```javascript
// ライブラリ使用可能性の確認とフォールバック
function getSanitizeFunction() {
    // DOMPurifyが利用可能かチェック
    if (typeof DOMPurify !== 'undefined' && DOMPurify.sanitize) {
        console.log('✅ DOMPurifyを使用したセキュアなサニタイゼーション');
        return sanitizeHtmlResponseSecure;
    } else {
        console.warn('⚠️ DOMPurifyが利用できません。カスタム実装を使用します');
        return sanitizeHtmlResponse; // 既存のカスタム実装
    }
}

// 動的な関数選択
const activeSanitizer = getSanitizeFunction();
```

### ステップ3: 段階的移行計画

```javascript
// Phase 1: 両方の実装を並行実行（検証期間）
function sanitizeWithValidation(html) {
    const customResult = sanitizeHtmlResponse(html);
    
    if (typeof DOMPurify !== 'undefined') {
        const libraryResult = sanitizeHtmlResponseSecure(html);
        
        // 結果の比較ログ
        if (customResult !== libraryResult) {
            console.log('🔍 サニタイゼーション結果の差異を検出:', {
                custom: customResult.length,
                library: libraryResult.length,
                original: html.length
            });
        }
        
        return libraryResult; // ライブラリ結果を優先
    }
    
    return customResult;
}
```

## 推奨実装戦略

### 短期計画 (現在の状況)
1. ✅ **カスタム実装の強化** (完了)
   - CodeQLの脆弱性を修正
   - 包括的なパターン検知
   - 詳細なログ記録

### 中期計画 (次のマイルストーン)
2. **DOMPurifyの段階的導入**
   - ビルドプロセスの設定
   - 外部ライブラリの組み込み
   - フォールバック機構の実装

### 長期計画 (安定化)
3. **ハイブリッド実装**
   - DOMPurifyによる主要サニタイゼーション
   - カスタム検知機能による追加監視
   - パフォーマンス最適化

## セキュリティ考慮事項

### ライブラリ使用時の注意点

1. **定期的な更新**
   ```bash
   # セキュリティ更新の確認
   npm audit
   npm update dompurify
   ```

2. **設定の適切性**
   ```javascript
   // 過度に制限的な設定は避ける
   // AI応答の表示品質とセキュリティのバランス
   ```

3. **パフォーマンス監視**
   ```javascript
   // サニタイゼーション処理時間の監視
   console.time('sanitization');
   const result = DOMPurify.sanitize(html);
   console.timeEnd('sanitization');
   ```

## まとめ

**現在の実装判断の妥当性**:
- ✅ Edge拡張の制約を適切に考慮
- ✅ CodeQLの脆弱性を修正
- ✅ 包括的なセキュリティ検知

**今後の改善方向**:
- 🔄 DOMPurifyの段階的導入検討
- 🔄 ビルドプロセスの最適化
- 🔄 ハイブリッド実装による最適解の実現

Edge拡張特有の制約により、**カスタム実装から開始したのは適切な判断**でした。今後、DOMPurifyなどの業界標準ライブラリとの統合を検討することで、さらなるセキュリティ強化が可能です。

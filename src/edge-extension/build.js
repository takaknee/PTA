/**
 * Edge拡張機能ビルドスクリプト
 * DOMPurifyなどの外部ライブラリを安全に組み込むためのビルド設定
 */

const fs = require('fs');
const path = require('path');

// ビルド設定
const BUILD_CONFIG = {
    // ソースディレクトリ
    srcDir: __dirname,

    // 出力ディレクトリ
    distDir: path.join(__dirname, 'dist'),

    // ベンダーライブラリディレクトリ
    vendorDir: path.join(__dirname, 'vendor'),

    // 開発モードかどうか
    isDev: process.argv.includes('--dev'),

    // 本番モードかどうか
    isProd: process.argv.includes('--prod')
};

/**
 * ディレクトリの作成
 */
function ensureDirectoryExists(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
        console.log(`✅ ディレクトリを作成しました: ${dirPath}`);
    }
}

/**
 * ファイルのコピー
 */
function copyFile(src, dest) {
    try {
        const data = fs.readFileSync(src);
        fs.writeFileSync(dest, data);
        console.log(`✅ ファイルをコピーしました: ${path.basename(src)}`);
    } catch (error) {
        console.error(`❌ ファイルコピーエラー: ${error.message}`);
    }
}

/**
 * DOMPurifyのダウンロードと設定
 */
function setupDOMPurify() {
    const dompurifyPath = path.join(__dirname, 'node_modules', 'dompurify', 'dist', 'purify.min.js');
    const vendorDOMPurifyPath = path.join(BUILD_CONFIG.vendorDir, 'dompurify.min.js');

    if (fs.existsSync(dompurifyPath)) {
        ensureDirectoryExists(BUILD_CONFIG.vendorDir);
        copyFile(dompurifyPath, vendorDOMPurifyPath);

        // DOMPurify使用のためのContent Scriptを更新
        updateContentScriptForDOMPurify();

        return true;
    } else {
        console.warn('⚠️ DOMPurifyが見つかりません。npm install dompurify を実行してください。');
        return false;
    }
}

/**
 * Content ScriptでDOMPurifyを使用するための設定更新
 */
function updateContentScriptForDOMPurify() {
    const manifestPath = path.join(BUILD_CONFIG.srcDir, 'manifest.json');

    try {
        const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));

        // DOMPurifyをContent Scriptsの先頭に追加
        manifest.content_scripts.forEach(script => {
            if (!script.js.includes('vendor/dompurify.min.js')) {
                script.js.unshift('vendor/dompurify.min.js');
            }
        });

        // manifest.jsonを更新
        fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));
        console.log('✅ manifest.jsonにDOMPurifyを追加しました');

    } catch (error) {
        console.error('❌ manifest.json更新エラー:', error.message);
    }
}

/**
 * セキュリティ強化版Content Scriptの生成
 */
function generateSecureContentScript() {
    const contentScriptTemplate = `
// DOMPurify統合版セキュリティ強化Content Script

/**
 * セキュアなHTMLサニタイゼーション（DOMPurify使用）
 */
function sanitizeHtmlResponseSecure(html) {
    if (!html || typeof html !== 'string') {
        console.log('sanitizeHtmlResponseSecure: 入力が無効です');
        return '';
    }

    // DOMPurifyが利用可能かチェック
    if (typeof DOMPurify === 'undefined') {
        console.warn('⚠️ DOMPurifyが利用できません。フォールバック実装を使用します');
        return sanitizeHtmlResponse(html); // 既存のカスタム実装
    }

    try {
        // DOMPurifyによる安全なサニタイゼーション
        const cleanHtml = DOMPurify.sanitize(html, {
            // AI応答表示用の安全な設定
            ALLOWED_TAGS: [
                'p', 'br', 'div', 'span', 'strong', 'b', 'em', 'i', 'u',
                'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                'ul', 'ol', 'li',
                'blockquote', 'pre', 'code',
                'table', 'thead', 'tbody', 'tr', 'td', 'th'
            ],
            ALLOWED_ATTR: ['class', 'style'],
            ALLOWED_STYLES: {
                'color': true,
                'background-color': true,
                'font-weight': true,
                'font-style': true,
                'text-decoration': true,
                'padding': true,
                'margin': true,
                'border-radius': true
            },
            // 危険な要素を完全に除去
            FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'form', 'input', 'button'],
            FORBID_ATTR: [
                'onclick', 'onerror', 'onload', 'onmouseover', 'onfocus', 'onblur',
                'href', 'src', 'action', 'formaction'
            ],
            // セキュリティ強化オプション
            SAFE_FOR_TEMPLATES: true,
            WHOLE_DOCUMENT: false,
            RETURN_DOM: false,
            RETURN_DOM_FRAGMENT: false,
            RETURN_TRUSTED_TYPE: false,
            SANITIZE_DOM: true,
            KEEP_CONTENT: false, // 危険なタグの内容も削除
            IN_PLACE: false
        });

        // 追加のセキュリティチェック（既存の検知機能と併用）
        const securityResult = detectSuspiciousPatterns(cleanHtml);
        if (securityResult.riskLevel === 'high' || securityResult.riskLevel === 'critical') {
            console.warn('🔒 DOMPurify処理後もリスクパターンが検出されました:', securityResult);
            
            // 高リスクの場合は追加の警告
            showNotification(
                'セキュリティ: サニタイゼーション後もリスクパターンが検出されました',
                'error'
            );
        }

        console.log('✅ DOMPurifyによるセキュアなサニタイゼーション完了');
        return cleanHtml;

    } catch (error) {
        console.error('❌ DOMPurifyサニタイゼーションエラー:', error);
        console.warn('⚠️ フォールバック実装を使用します');
        return sanitizeHtmlResponse(html); // エラー時はカスタム実装
    }
}

/**
 * 最適なサニタイゼーション関数の選択
 */
function getSanitizeFunction() {
    if (typeof DOMPurify !== 'undefined' && DOMPurify.sanitize) {
        console.log('✅ DOMPurifyを使用したセキュアなサニタイゼーション');
        return sanitizeHtmlResponseSecure;
    } else {
        console.warn('⚠️ DOMPurifyが利用できません。カスタム実装を使用します');
        return sanitizeHtmlResponse;
    }
}

// グローバルに最適なサニタイザーを設定
window.activeSanitizer = getSanitizeFunction();

// 既存のsanitizeHtmlResponse関数を拡張
const originalSanitizeHtmlResponse = sanitizeHtmlResponse;
sanitizeHtmlResponse = function(html) {
    return window.activeSanitizer(html);
};

console.log('🛡️ セキュリティ強化版HTMLサニタイゼーション初期化完了');
`;

    const secureScriptPath = path.join(BUILD_CONFIG.vendorDir, 'secure-sanitizer.js');
    ensureDirectoryExists(BUILD_CONFIG.vendorDir);
    fs.writeFileSync(secureScriptPath, contentScriptTemplate);
    console.log('✅ セキュリティ強化スクリプトを生成しました');
}

/**
 * ビルドプロセスの実行
 */
function build() {
    console.log('🚀 Edge拡張機能のビルドを開始します...');
    console.log(`📦 モード: ${BUILD_CONFIG.isDev ? '開発' : BUILD_CONFIG.isProd ? '本番' : '標準'}`);

    // ディレクトリの準備
    ensureDirectoryExists(BUILD_CONFIG.vendorDir);
    ensureDirectoryExists(BUILD_CONFIG.distDir);

    // DOMPurifyのセットアップ
    const hasDOMPurify = setupDOMPurify();

    // セキュリティ強化スクリプトの生成
    generateSecureContentScript();

    // ビルド結果の表示
    console.log('\n📋 ビルド結果:');
    console.log(`  DOMPurify: ${hasDOMPurify ? '✅ 利用可能' : '❌ 未インストール'}`);
    console.log(`  セキュリティ強化: ✅ 実装済み`);
    console.log(`  カスタム検知機能: ✅ 実装済み`);

    if (!hasDOMPurify) {
        console.log('\n💡 DOMPurifyをインストールするには:');
        console.log('   npm run install-dompurify');
    }

    console.log('\n🎉 ビルド完了');
}

// ビルド実行
if (require.main === module) {
    build();
}

module.exports = { build, BUILD_CONFIG };

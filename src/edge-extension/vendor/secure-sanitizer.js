
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

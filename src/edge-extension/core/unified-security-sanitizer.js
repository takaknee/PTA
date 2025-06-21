/**
 * 統一HTMLサニタイゼーション・セキュリティ管理モジュール
 * 
 * 【重要】このモジュールは全てのHTML/テキスト処理で使用する統一サニタイザーです
 * - DOMPurifyベースの包括的サニタイゼーション
 * - XSS攻撃防止
 * - テンプレート処理のセキュリティ強化
 * - 不完全な多文字サニタイゼーション問題の解消
 * 
 * 開発時の注意事項:
 * 1. 全ての外部入力（ユーザー入力、API応答、URL、ページコンテンツ）は必ずこのモジュールを通すこと
 * 2. 正規表現による直接的なHTML除去は禁止（不完全サニタイゼーションの原因）
 * 3. テンプレート置換時は必ずescapeUserInput()を使用すること
 * 4. プロンプトテンプレート内の動的コンテンツは事前サニタイゼーション必須
 */

/**
 * 統一セキュリティサニタイザークラス
 */
class UnifiedSecuritySanitizer {
    constructor() {
        this.domPurifyAvailable = false;
        this.initializeDOMPurify();
    }

    /**
     * DOMPurifyの初期化確認
     */
    initializeDOMPurify() {
        try {
            // Service Worker環境でのDOMPurify確認
            if (typeof DOMPurify !== 'undefined') {
                this.domPurifyAvailable = true;
                console.log('✅ UnifiedSecuritySanitizer: DOMPurify利用可能');
            } else if (typeof globalThis !== 'undefined' && globalThis.DOMPurify) {
                this.domPurifyAvailable = true;
                console.log('✅ UnifiedSecuritySanitizer: globalThis.DOMPurify利用可能');
            } else {
                console.warn('⚠️ UnifiedSecuritySanitizer: DOMPurify未利用、フォールバック使用');
            }
        } catch (error) {
            console.error('❌ UnifiedSecuritySanitizer: DOMPurify初期化エラー:', error);
        }
    }

    /**
     * 包括的HTMLサニタイゼーション（DOMPurifyベース）
     * 【重要】これは正規表現ベースの不完全サニタイゼーションを置き換える統一メソッドです
     * 
     * @param {string} html - サニタイゼーション対象のHTML
     * @param {Object} options - サニタイゼーションオプション
     * @returns {string} - サニタイズされた安全なHTML
     */
    sanitizeHTML(html, options = {}) {
        if (!html || typeof html !== 'string') {
            return '';
        }

        // 長さ制限（DoS攻撃防止）
        if (html.length > (options.maxLength || 100000)) {
            console.warn('UnifiedSecuritySanitizer: 入力が長すぎるため切り詰めます');
            html = html.substring(0, options.maxLength || 100000);
        }

        try {
            if (this.domPurifyAvailable) {
                // DOMPurifyが利用可能な場合
                const config = {
                    ALLOWED_TAGS: options.allowedTags || ['p', 'br', 'strong', 'em', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'blockquote', 'code', 'pre', 'div', 'span'],
                    ALLOWED_ATTR: options.allowedAttributes || ['class', 'id'],
                    FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input', 'textarea', 'select', 'button'],
                    FORBID_ATTR: ['onclick', 'onload', 'onerror', 'onmouseover', 'onfocus', 'onblur', 'onchange', 'onsubmit'],
                    ALLOW_DATA_ATTR: false,
                    WHOLE_DOCUMENT: false,
                    RETURN_DOM: false,
                    RETURN_DOM_FRAGMENT: false
                };

                const sanitizer = this.domPurifyAvailable === true ? DOMPurify : globalThis.DOMPurify;
                return sanitizer.sanitize(html, config);
            } else {
                // フォールバック: 改良された安全な正規表現処理
                return this.fallbackSanitizeHTML(html, options);
            }
        } catch (error) {
            console.error('UnifiedSecuritySanitizer: サニタイゼーションエラー:', error);
            return this.fallbackSanitizeHTML(html, options);
        }
    }

    /**
     * テキストのみ抽出（HTMLタグ完全除去）
     * 正規表現の不完全サニタイゼーション問題を解決する統一メソッド
     * 
     * @param {string} html - HTML文字列
     * @param {Object} options - 抽出オプション
     * @returns {string} - プレーンテキスト
     */
    extractPlainText(html, options = {}) {
        if (!html || typeof html !== 'string') {
            return '';
        }

        try {
            if (this.domPurifyAvailable) {
                // DOMPurifyを使用してテキストのみ抽出
                const sanitizer = this.domPurifyAvailable === true ? DOMPurify : globalThis.DOMPurify;
                const cleaned = sanitizer.sanitize(html, {
                    ALLOWED_TAGS: [],
                    ALLOWED_ATTR: [],
                    KEEP_CONTENT: true
                });

                return this.normalizeWhitespace(cleaned, options);
            } else {
                // フォールバック処理
                return this.fallbackExtractPlainText(html, options);
            }
        } catch (error) {
            console.error('UnifiedSecuritySanitizer: テキスト抽出エラー:', error);
            return this.fallbackExtractPlainText(html, options);
        }
    }

    /**
     * ユーザー入力の安全なエスケープ処理
     * プロンプトテンプレート等での動的コンテンツ挿入時に使用
     * 
     * @param {string} userInput - ユーザー入力文字列
     * @param {Object} options - エスケープオプション
     * @returns {string} - エスケープされた安全な文字列
     */
    escapeUserInput(userInput, options = {}) {
        if (!userInput || typeof userInput !== 'string') {
            return '';
        }

        // 長さ制限
        if (userInput.length > (options.maxLength || 10000)) {
            userInput = userInput.substring(0, options.maxLength || 10000);
        }

        // HTML特殊文字のエスケープ
        const htmlEscapeMap = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#x27;',
            '/': '&#x2F;',
            '`': '&#x60;',
            '=': '&#x3D;'
        };

        let escaped = userInput.replace(/[&<>"'`=\/]/g, (match) => htmlEscapeMap[match]);

        // JavaScript/URL プロトコルの除去
        escaped = escaped.replace(/javascript:/gi, 'javascript-removed:');
        escaped = escaped.replace(/vbscript:/gi, 'vbscript-removed:');
        escaped = escaped.replace(/data:/gi, 'data-removed:');

        // 制御文字の正規化
        escaped = this.normalizeWhitespace(escaped, options);

        return escaped;
    }

    /**
     * プロンプトテンプレートの安全な構築
     * 動的コンテンツ挿入時の統一セキュリティ処理
     * 
     * @param {string} template - プロンプトテンプレート
     * @param {Object} variables - 置換変数オブジェクト
     * @param {Object} options - テンプレート処理オプション
     * @returns {string} - 安全に構築されたプロンプト
     */
    buildSecurePrompt(template, variables = {}, options = {}) {
        if (!template || typeof template !== 'string') {
            return '';
        }

        let securePrompt = template;

        // 各変数を安全にエスケープしてから置換
        Object.keys(variables).forEach(key => {
            const placeholder = `{{${key}}}`;
            const rawValue = variables[key] || '';

            // 値の型チェック
            const stringValue = typeof rawValue === 'string' ? rawValue : String(rawValue);

            // コンテンツタイプに応じた適切なサニタイゼーション
            let sanitizedValue;
            if (options.preserveHTML && (key === 'pageContent' || key.includes('html'))) {
                // HTMLを保持する場合
                sanitizedValue = this.sanitizeHTML(stringValue, {
                    allowedTags: ['p', 'br', 'strong', 'em', 'h1', 'h2', 'h3', 'ul', 'ol', 'li', 'code'],
                    maxLength: 50000
                });
            } else {
                // プレーンテキストとして処理
                sanitizedValue = this.escapeUserInput(stringValue, {
                    maxLength: options.maxVariableLength || 10000
                });
            }

            // グローバル置換（replaceAllの代替実装）
            securePrompt = securePrompt.split(placeholder).join(sanitizedValue);
        });

        return securePrompt;
    }

    /**
     * フォールバック用HTMLサニタイゼーション
     * DOMPurifyが利用できない場合の安全な代替処理
     */
    fallbackSanitizeHTML(html, options = {}) {
        // 危険なタグを除去（より安全な正規表現使用）
        let cleaned = html
            // スクリプト系タグの完全除去
            .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
            .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
            .replace(/<iframe[^>]*>[\s\S]*?<\/iframe>/gi, '')
            .replace(/<object[^>]*>[\s\S]*?<\/object>/gi, '')
            .replace(/<embed[^>]*>[\s\S]*?<\/embed>/gi, '')
            .replace(/<form[^>]*>[\s\S]*?<\/form>/gi, '')
            // 危険な属性の除去
            .replace(/\s+on\w+\s*=\s*["'][^"']*["']/gi, '')
            .replace(/\s+on\w+\s*=\s*[^\s>]+/gi, '')
            // 危険なプロトコルの除去
            .replace(/javascript:/gi, 'javascript-removed:')
            .replace(/vbscript:/gi, 'vbscript-removed:')
            .replace(/data:/gi, 'data-removed:');

        return this.normalizeWhitespace(cleaned, options);
    }

    /**
     * フォールバック用プレーンテキスト抽出
     */
    fallbackExtractPlainText(html, options = {}) {
        // 段階的な安全なHTML除去
        let text = html;
        let previous;
        // 危険なコンテンツの除去（繰り返し処理）
        do {
            previous = text;
            text = text
                .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
                .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                .replace(/<iframe[^>]*>[\s\S]*?<\/iframe>/gi, '');
        } while (text !== previous);
        // HTMLタグの除去（改良版）
        text = text.replace(/<[^>]+>/g, ' ');
        // HTMLエンティティのデコード（安全な文字のみ）
        text = text
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/&quot;/g, '"')
            .replace(/&#x27;/g, "'");
        // 危険なプロトコルの除去（改良版）
        text = text
            .replace(/javascript\s*:/gi, '')
            .replace(/vbscript\s*:/gi, '')
            .replace(/data\s*:/gi, '');

        return this.normalizeWhitespace(text, options);
    }

    /**
     * 空白文字の正規化
     */
    normalizeWhitespace(text, options = {}) {
        if (!text || typeof text !== 'string') {
            return '';
        }

        return text
            // 制御文字を空白に置換
            .replace(/[\r\n\t\f\v]/g, ' ')
            // 連続する空白を単一空白に統合
            .replace(/\s+/g, ' ')
            // 前後の空白を除去
            .trim()
            // 長さ制限
            .substring(0, options.maxLength || 50000);
    }

    /**
     * URLの安全性検証
     */
    validateURL(url) {
        if (!url || typeof url !== 'string') {
            return false;
        }

        try {
            const urlObj = new URL(url);
            // HTTPSまたはHTTPのみ許可
            return ['https:', 'http:'].includes(urlObj.protocol);
        } catch (error) {
            return false;
        }
    }

    /**
     * セキュリティテスト実行
     */
    runSecurityTests() {
        console.log('UnifiedSecuritySanitizer: セキュリティテストを開始');

        const testCases = [
            {
                name: 'XSSスクリプト除去テスト',
                input: '<p>Hello <script>alert("XSS")</script>World</p>',
                expected: (result) => result.includes('Hello') && result.includes('World') && !result.includes('script')
            },
            {
                name: 'イベントハンドラー除去テスト',
                input: '<div onclick="alert(1)">Click me</div>',
                expected: (result) => result.includes('Click me') && !result.includes('onclick')
            },
            {
                name: 'プロトコル除去テスト',
                input: 'Visit javascript:alert(1) or vbscript:msgbox(1)',
                expected: (result) => result.includes('Visit') && !result.includes('javascript:') && !result.includes('vbscript:')
            }
        ];

        let passed = 0;
        testCases.forEach((test, index) => {
            try {
                const result = this.sanitizeHTML(test.input);
                const testPassed = test.expected(result);
                console.log(`テスト ${index + 1} (${test.name}): ${testPassed ? '✅ 合格' : '❌ 不合格'}`);
                console.log(`  入力: ${test.input}`);
                console.log(`  結果: ${result}`);
                if (testPassed) passed++;
            } catch (error) {
                console.error(`テスト ${index + 1} エラー:`, error);
            }
        });

        console.log(`セキュリティテスト結果: ${passed}/${testCases.length} 合格`);
        return passed === testCases.length;
    }
}

// グローバルインスタンスの作成
const unifiedSanitizer = new UnifiedSecuritySanitizer();

// 後方互換性のための統一エクスポート
globalThis.UnifiedSecuritySanitizer = UnifiedSecuritySanitizer;
globalThis.unifiedSanitizer = unifiedSanitizer;

// Service Worker環境での利用を考慮した複数の公開方法
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { UnifiedSecuritySanitizer, unifiedSanitizer };
}

console.log('✅ 統一セキュリティサニタイザーモジュールが初期化されました');

/*
 * PTA AI業務支援ツール - 定数定義モジュール
 * Copyright (c) 2024 PTA Development Team
 */

/**
 * アプリケーション全体の定数定義
 */
export const APP_CONSTANTS = {
    // アプリケーション情報
    APP_NAME: 'PTA AI業務支援ツール',
    VERSION: '1.0.0',

    // API設定
    API: {
        TIMEOUT: 30000, // 30秒
        MAX_RETRIES: 3,
        RETRY_DELAY: 1000, // 1秒
        RATE_LIMIT_DELAY: 100, // 100ms
    },

    // コンテンツ制限
    CONTENT_LIMITS: {
        MAX_TEXT_LENGTH: 50000,
        MAX_HTML_LENGTH: 100000,
        MAX_SELECTION_LENGTH: 10000,
        SUMMARY_MIN_LENGTH: 100,
    },

    // UI設定
    UI: {
        ANIMATION_DURATION: 300,
        DEBOUNCE_DELAY: 300,
        NOTIFICATION_TIMEOUT: 5000,
        LOADING_MIN_DISPLAY: 500,
    },

    // セキュリティ設定
    SECURITY: {
        ALLOWED_PROTOCOLS: ['https:', 'http:'],
        BLOCKED_ELEMENTS: ['script', 'style', 'iframe', 'object', 'embed'],
        BLOCKED_ATTRIBUTES: ['onclick', 'onload', 'onerror', 'javascript:', 'vbscript:'],
        MAX_INPUT_SIZE: 1000000, // 1MB
    },

    // サービス固有設定
    SERVICES: {
        OUTLOOK: {
            DOMAINS: ['outlook.office.com', 'outlook.live.com'],
            SELECTORS: {
                EMAIL_CONTENT: '[role="main"] [role="document"]',
                EMAIL_SUBJECT: '[aria-label*="件名"]',
                EMAIL_FROM: '[aria-label*="差出人"]',
                EMAIL_DATE: '[aria-label*="受信日時"]',
            }
        },
        GMAIL: {
            DOMAINS: ['mail.google.com'],
            SELECTORS: {
                EMAIL_CONTENT: '.ii.gt .a3s.aiL',
                EMAIL_SUBJECT: '.hP',
                EMAIL_FROM: '.go .g2',
                EMAIL_DATE: '.g3',
            }
        }
    },

    // ストレージキー
    STORAGE_KEYS: {
        API_SETTINGS: 'apiSettings',
        USER_PREFERENCES: 'userPreferences',
        THEME_SETTINGS: 'themeSettings',
        ANALYSIS_HISTORY: 'analysisHistory',
        CACHE_SETTINGS: 'cacheSettings',
    },

    // エラーメッセージ
    ERROR_MESSAGES: {
        API_KEY_MISSING: 'APIキーが設定されていません。設定画面で設定してください。',
        API_CONNECTION_FAILED: 'APIへの接続に失敗しました。設定とネットワーク接続を確認してください。',
        CONTENT_TOO_LARGE: 'コンテンツが大きすぎます。選択範囲を小さくしてください。',
        INVALID_FORMAT: '無効な形式のデータです。',
        PERMISSION_DENIED: '必要な権限がありません。',
        TIMEOUT_ERROR: '処理がタイムアウトしました。再試行してください。',
        UNKNOWN_ERROR: '予期しないエラーが発生しました。',
    },

    // 成功メッセージ
    SUCCESS_MESSAGES: {
        ANALYSIS_COMPLETE: '分析が完了しました',
        SETTINGS_SAVED: '設定が保存されました',
        COPIED_TO_CLIPBOARD: 'クリップボードにコピーしました',
        EMAIL_FORWARDED: 'メールを転送しました',
        CALENDAR_ADDED: 'カレンダーに追加しました',
    },

    // プロンプトテンプレート
    PROMPTS: {
        EMAIL_ANALYSIS: {
            SYSTEM: 'あなたは日本語での業務メール分析の専門家です。メール内容を分析し、要点、重要度、推奨アクションを日本語で提供してください。',
            USER_TEMPLATE: 'メール件名: {{subject}}\n差出人: {{from}}\n日時: {{date}}\n\nメール内容:\n{{content}}\n\n上記のメールを分析し、以下の形式で回答してください：\n1. 要約（2-3行）\n2. 重要度（高・中・低）\n3. 推奨アクション\n4. 回答期限（もしあれば）',
        },
        QUICK_REPLY: {
            SYSTEM: 'あなたは日本語での丁寧で効率的なビジネスメール作成の専門家です。',
            USER_TEMPLATE: '以下のメールに対する返信を作成してください：\n\n{{originalEmail}}\n\n返信のトーン: {{tone}}\n特記事項: {{notes}}',
        },
        PAGE_SUMMARY: {
            SYSTEM: 'あなたは日本語でのWebページ要約の専門家です。重要な情報を簡潔にまとめてください。',
            USER_TEMPLATE: 'ページタイトル: {{title}}\nURL: {{url}}\n\nページ内容:\n{{content}}\n\n上記のページの要約を作成してください。',
        }
    }
};

/**
 * 開発/本番環境固有の設定
 */
export const ENV_CONSTANTS = {
    DEVELOPMENT: {
        DEBUG_MODE: true,
        LOG_LEVEL: 'debug',
        API_TIMEOUT: 60000, // 開発時は長めに設定
    },
    PRODUCTION: {
        DEBUG_MODE: false,
        LOG_LEVEL: 'error',
        API_TIMEOUT: 30000,
    }
};

/**
 * レスポンシブブレークポイント
 */
export const BREAKPOINTS = {
    MOBILE: 768,
    TABLET: 1024,
    DESKTOP: 1200,
};

/**
 * カラーテーマ定数
 */
export const COLORS = {
    PRIMARY: '#0078d4',
    SECONDARY: '#6c757d',
    SUCCESS: '#28a745',
    WARNING: '#ffc107',
    ERROR: '#dc3545',
    INFO: '#17a2b8',
    LIGHT: '#f8f9fa',
    DARK: '#343a40',
};

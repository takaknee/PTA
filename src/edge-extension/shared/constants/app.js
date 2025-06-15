/*
 * AI業務支援ツール Edge拡張機能 - アプリケーション定数
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * メッセージアクション定数
 */
export const MESSAGE_ACTIONS = Object.freeze({
  // コンテンツ解析
  ANALYZE_EMAIL: 'analyzeEmail',
  ANALYZE_PAGE: 'analyzePage',
  ANALYZE_SELECTION: 'analyzeSelection',

  // メール関連
  COMPOSE_EMAIL: 'composeEmail',
  REPLY_EMAIL: 'replyEmail',

  // 翻訳機能
  TRANSLATE_SELECTION: 'translateSelection',
  TRANSLATE_PAGE: 'translatePage',

  // ユーティリティ
  EXTRACT_URLS: 'extractUrls',
  COPY_PAGE_INFO: 'copyPageInfo',
  SUMMARIZE_CONTENT: 'summarizeContent',

  // システム
  TEST_API: 'testAPI',
  CONNECTION_TEST: 'connectionTest',
  SYSTEM_DIAGNOSTICS: 'systemDiagnostics',

  // 設定
  GET_SETTINGS: 'getSettings',
  SAVE_SETTINGS: 'saveSettings',
  RESET_SETTINGS: 'resetSettings',

  // 履歴
  GET_HISTORY: 'getHistory',
  SAVE_HISTORY: 'saveHistory',
  CLEAR_HISTORY: 'clearHistory',
  DELETE_HISTORY_ITEM: 'deleteHistoryItem',

  // 統計
  GET_STATISTICS: 'getStatistics',
  UPDATE_STATISTICS: 'updateStatistics',
  RESET_STATISTICS: 'resetStatistics',

  // 通知
  SHOW_NOTIFICATION: 'showNotification',
  HIDE_NOTIFICATION: 'hideNotification'
});

/**
 * メッセージターゲット定数
 */
export const MESSAGE_TARGETS = Object.freeze({
  BACKGROUND: 'background',
  CONTENT: 'content',
  OFFSCREEN: 'offscreen',
  POPUP: 'popup',
  OPTIONS: 'options'
});

/**
 * 履歴項目タイプ定数
 */
export const HISTORY_TYPES = Object.freeze({
  EMAIL_ANALYSIS: 'emailAnalysis',
  PAGE_ANALYSIS: 'pageAnalysis',
  SELECTION_ANALYSIS: 'selectionAnalysis',
  EMAIL_COMPOSITION: 'emailComposition',
  EMAIL_REPLY: 'emailReply',
  TRANSLATION: 'translation',
  URL_EXTRACTION: 'urlExtraction',
  CONTENT_SUMMARY: 'contentSummary'
});

/**
 * 通知タイプ定数
 */
export const NOTIFICATION_TYPES = Object.freeze({
  SUCCESS: 'success',
  ERROR: 'error',
  WARNING: 'warning',
  INFO: 'info'
});

/**
 * コンテキストメニューID定数
 */
export const CONTEXT_MENU_IDS = Object.freeze({
  ANALYZE_PAGE: 'ai-analyze-page',
  ANALYZE_SELECTION: 'ai-analyze-selection',
  TRANSLATE_SELECTION: 'ai-translate-selection',
  TRANSLATE_PAGE: 'ai-translate-page',
  EXTRACT_URLS: 'ai-extract-urls',
  COPY_PAGE_INFO: 'ai-copy-page-info',
  SUMMARIZE_CONTENT: 'ai-summarize-content'
});

/**
 * サポートされるサービス定数
 */
export const SUPPORTED_SERVICES = Object.freeze({
  OUTLOOK: 'outlook',
  GMAIL: 'gmail',
  GENERAL: 'general'
});

/**
 * ストレージキー定数
 */
export const STORAGE_KEYS = Object.freeze({
  SETTINGS: 'aiToolSettings',
  HISTORY: 'aiToolHistory',
  STATISTICS: 'aiToolStatistics',
  BUTTON_POSITION: 'aiButtonPosition',
  USER_PREFERENCES: 'aiUserPreferences',
  CACHE: 'aiToolCache'
});

/**
 * システム設定定数
 */
export const SYSTEM_CONFIG = Object.freeze({
  // 履歴管理
  MAX_HISTORY_ITEMS: 100,
  HISTORY_CLEANUP_INTERVAL: 24 * 60 * 60 * 1000, // 24時間

  // パフォーマンス
  BATCH_SIZE: 50,
  DEBOUNCE_DELAY: 300,
  CONNECTION_TIMEOUT: 10000,

  // コンテンツ処理
  MAX_TEXT_LENGTH: 10000,
  MAX_SELECTION_LENGTH: 5000,

  // リトライ設定
  RETRY_ATTEMPTS: 3,
  RETRY_DELAY: 1000,
  EXPONENTIAL_BACKOFF: true,

  // キャッシュ設定
  CACHE_DURATION: 5 * 60 * 1000, // 5分
  MAX_CACHE_ENTRIES: 50,

  // 診断設定
  DIAGNOSTIC_INTERVAL: 60 * 1000, // 1分
  MAX_LOG_ENTRIES: 1000
});

/**
 * デフォルト設定
 */
export const DEFAULT_SETTINGS = Object.freeze({
  api: {
    provider: 'azure',
    model: 'gpt-4o-mini',
    apiKey: '',
    azureEndpoint: '',
    azureDeploymentName: '',
    maxTokens: 2000,
    temperature: 0.7,
    requestTimeout: 30000
  },

  ui: {
    theme: 'auto', // 'light', 'dark', 'auto'
    showNotifications: true,
    notificationDuration: 3000,
    buttonPosition: { x: 20, y: 20 },
    expandByDefault: false
  },

  features: {
    autoDetect: true,
    saveHistory: true,
    enableStatistics: true,
    enableCache: true,
    enableDiagnostics: false
  },

  advanced: {
    debugMode: false,
    logLevel: 'info',
    enableTelemetry: false
  }
});

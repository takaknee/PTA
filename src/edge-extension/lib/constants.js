/*
 * AI業務支援ツール Edge拡張機能 - 定数定義
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * APIプロバイダー定数
 */
export const API_PROVIDERS = {
  OPENAI: 'openai',
  AZURE: 'azure'
};

/**
 * AIモデル定数
 */
export const AI_MODELS = {
  GPT_4O_MINI: 'gpt-4o-mini',
  GPT_4O: 'gpt-4o',
  GPT_35_TURBO: 'gpt-3.5-turbo',
  GPT_4: 'gpt-4'
};

/**
 * メッセージアクション定数
 */
export const MESSAGE_ACTIONS = {
  ANALYZE_EMAIL: 'analyzeEmail',
  ANALYZE_PAGE: 'analyzePage',
  ANALYZE_SELECTION: 'analyzeSelection',
  COMPOSE_EMAIL: 'composeEmail',
  TRANSLATE_SELECTION: 'translateSelection',
  TRANSLATE_PAGE: 'translatePage',
  EXTRACT_URLS: 'extractUrls',
  COPY_PAGE_INFO: 'copyPageInfo',
  TEST_API: 'testAPI',
  CONNECTION_TEST: 'connectionTest',
  FETCH_API: 'fetchAPI',
  SYSTEM_DIAGNOSTICS: 'systemDiagnostics'
};

/**
 * メッセージターゲット定数
 */
export const MESSAGE_TARGETS = {
  BACKGROUND: 'background',
  CONTENT: 'content',
  OFFSCREEN: 'offscreen',
  POPUP: 'popup',
  OPTIONS: 'options'
};

/**
 * デフォルト設定
 */
export const DEFAULT_SETTINGS = {
  provider: API_PROVIDERS.AZURE,
  model: AI_MODELS.GPT_4O_MINI,
  apiKey: '',
  azureEndpoint: '',
  azureDeploymentName: '',
  autoDetect: true,
  showNotifications: true,
  saveHistory: true,
  maxTokens: 2000,
  temperature: 0.7,
  requestTimeout: 30000
};

/**
 * APIエンドポイント定数
 */
export const API_ENDPOINTS = {
  OPENAI: 'https://api.openai.com/v1/chat/completions',
  AZURE_TEMPLATE: 'https://{resource}.openai.azure.com/openai/deployments/{deployment}/chat/completions'
};

/**
 * エラーメッセージ定数
 */
export const ERROR_MESSAGES = {
  INVALID_API_KEY: 'APIキーが無効です。設定を確認してください。',
  ACCESS_DENIED: 'APIアクセスが拒否されました。権限またはデプロイメント名を確認してください。',
  ENDPOINT_NOT_FOUND: 'APIエンドポイントが見つかりません。URLまたはデプロイメント名を確認してください。',
  RATE_LIMIT_EXCEEDED: 'API利用制限に達しました。しばらく待ってから再試行してください。',
  SERVER_ERROR: 'APIサーバーエラーが発生しました。しばらく待ってから再試行してください。',
  NETWORK_ERROR: 'ネットワーク接続エラーが発生しました。インターネット接続を確認してください。',
  INVALID_RESPONSE: 'AIからの有効な応答が得られませんでした。',
  SETTINGS_NOT_CONFIGURED: 'API設定が完了していません。設定画面で設定を行ってください。',
  NO_CONTENT: 'コンテンツが見つかりません。',
  PROCESSING_FAILED: '処理に失敗しました。'
};

/**
 * 通知タイプ定数
 */
export const NOTIFICATION_TYPES = {
  SUCCESS: 'success',
  ERROR: 'error',
  WARNING: 'warning',
  INFO: 'info'
};

/**
 * ストレージキー定数
 */
export const STORAGE_KEYS = {
  SETTINGS: 'aiToolSettings',
  HISTORY: 'aiToolHistory',
  STATISTICS: 'aiToolStatistics',
  BUTTON_POSITION: 'aiButtonPosition'
};

/**
 * HTML要素ID定数
 */
export const ELEMENT_IDS = {
  AI_BUTTON: 'ai-support-button',
  AI_DIALOG: 'ai-dialog',
  NOTIFICATION: 'ai-notification',
  LOADING: 'ai-loading'
};

/**
 * CSS クラス定数
 */
export const CSS_CLASSES = {
  AI_BUTTON: 'ai-support-button',
  AI_DIALOG: 'ai-dialog',
  NOTIFICATION: 'ai-notification',
  LOADING: 'ai-loading',
  SUCCESS: 'success',
  ERROR: 'error',
  WARNING: 'warning',
  INFO: 'info',
  DARK_THEME: 'dark-theme',
  LIGHT_THEME: 'light-theme'
};

/**
 * 履歴項目タイプ定数
 */
export const HISTORY_TYPES = {
  EMAIL_ANALYSIS: 'emailAnalysis',
  PAGE_ANALYSIS: 'pageAnalysis',
  SELECTION_ANALYSIS: 'selectionAnalysis',
  EMAIL_COMPOSITION: 'emailComposition',
  TRANSLATION: 'translation',
  URL_EXTRACTION: 'urlExtraction'
};

/**
 * コンテキストメニューID定数
 */
export const CONTEXT_MENU_IDS = {
  ANALYZE_PAGE: 'analyzePage',
  ANALYZE_SELECTION: 'analyzeSelection',
  TRANSLATE_SELECTION: 'translateSelection',
  EXTRACT_URLS: 'extractUrls',
  COPY_PAGE_INFO: 'copyPageInfo'
};

/**
 * システム設定定数
 */
export const SYSTEM_CONFIG = {
  MAX_HISTORY_ITEMS: 100,
  BATCH_SIZE: 50,
  RETRY_ATTEMPTS: 3,
  RETRY_DELAY: 1000,
  CONNECTION_TIMEOUT: 10000,
  MAX_TEXT_LENGTH: 10000,
  DEBOUNCE_DELAY: 300
};

/**
 * デバッグ設定
 */
export const DEBUG_CONFIG = {
  ENABLED: true,
  LOG_LEVEL: 'info',
  MAX_LOG_ENTRIES: 1000
};

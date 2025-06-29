/*
 * AI業務支援ツール Edge拡張機能 - エラー定数
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * エラーカテゴリー定数
 */
export const ERROR_CATEGORIES = Object.freeze({
  API: 'api',
  NETWORK: 'network',
  VALIDATION: 'validation',
  PERMISSION: 'permission',
  STORAGE: 'storage',
  UI: 'ui',
  PARSING: 'parsing',
  SYSTEM: 'system'
});

/**
 * エラーコード定数
 */
export const ERROR_CODES = Object.freeze({
  // API関連エラー
  API_INVALID_KEY: 'API_INVALID_KEY',
  API_ACCESS_DENIED: 'API_ACCESS_DENIED',
  API_ENDPOINT_NOT_FOUND: 'API_ENDPOINT_NOT_FOUND',
  API_RATE_LIMIT_EXCEEDED: 'API_RATE_LIMIT_EXCEEDED',
  API_SERVER_ERROR: 'API_SERVER_ERROR',
  API_INVALID_RESPONSE: 'API_INVALID_RESPONSE',
  API_TIMEOUT: 'API_TIMEOUT',

  // ネットワーク関連エラー
  NETWORK_CONNECTION_FAILED: 'NETWORK_CONNECTION_FAILED',
  NETWORK_TIMEOUT: 'NETWORK_TIMEOUT',
  NETWORK_OFFLINE: 'NETWORK_OFFLINE',
  NETWORK_CORS_ERROR: 'NETWORK_CORS_ERROR',

  // バリデーション関連エラー
  VALIDATION_REQUIRED_FIELD: 'VALIDATION_REQUIRED_FIELD',
  VALIDATION_INVALID_FORMAT: 'VALIDATION_INVALID_FORMAT',
  VALIDATION_LENGTH_EXCEEDED: 'VALIDATION_LENGTH_EXCEEDED',
  VALIDATION_INVALID_VALUE: 'VALIDATION_INVALID_VALUE',

  // 権限関連エラー
  PERMISSION_DENIED: 'PERMISSION_DENIED',
  PERMISSION_STORAGE_ACCESS: 'PERMISSION_STORAGE_ACCESS',
  PERMISSION_HOST_ACCESS: 'PERMISSION_HOST_ACCESS',

  // ストレージ関連エラー
  STORAGE_QUOTA_EXCEEDED: 'STORAGE_QUOTA_EXCEEDED',
  STORAGE_ACCESS_DENIED: 'STORAGE_ACCESS_DENIED',
  STORAGE_CORRUPTION: 'STORAGE_CORRUPTION',

  // UI関連エラー
  UI_ELEMENT_NOT_FOUND: 'UI_ELEMENT_NOT_FOUND',
  UI_RENDER_FAILED: 'UI_RENDER_FAILED',
  UI_INTERACTION_FAILED: 'UI_INTERACTION_FAILED',

  // 解析関連エラー
  PARSING_CONTENT_FAILED: 'PARSING_CONTENT_FAILED',
  PARSING_EMAIL_FAILED: 'PARSING_EMAIL_FAILED',
  PARSING_PAGE_FAILED: 'PARSING_PAGE_FAILED',

  // システム関連エラー
  SYSTEM_INITIALIZATION_FAILED: 'SYSTEM_INITIALIZATION_FAILED',
  SYSTEM_SERVICE_UNAVAILABLE: 'SYSTEM_SERVICE_UNAVAILABLE',
  SYSTEM_CONFIGURATION_ERROR: 'SYSTEM_CONFIGURATION_ERROR',
  SYSTEM_UNKNOWN_ERROR: 'SYSTEM_UNKNOWN_ERROR'
});

/**
 * エラーメッセージ定数
 */
export const ERROR_MESSAGES = Object.freeze({
  // API関連メッセージ
  [ERROR_CODES.API_INVALID_KEY]: 'APIキーが無効です。設定を確認してください。',
  [ERROR_CODES.API_ACCESS_DENIED]: 'APIアクセスが拒否されました。権限またはデプロイメント名を確認してください。',
  [ERROR_CODES.API_ENDPOINT_NOT_FOUND]: 'APIエンドポイントが見つかりません。URLまたはデプロイメント名を確認してください。',
  [ERROR_CODES.API_RATE_LIMIT_EXCEEDED]: 'API利用制限に達しました。しばらく待ってから再試行してください。',
  [ERROR_CODES.API_SERVER_ERROR]: 'APIサーバーエラーが発生しました。しばらく待ってから再試行してください。',
  [ERROR_CODES.API_INVALID_RESPONSE]: 'AIからの有効な応答が得られませんでした。',
  [ERROR_CODES.API_TIMEOUT]: 'API応答がタイムアウトしました。しばらく待ってから再試行してください。',

  // ネットワーク関連メッセージ
  [ERROR_CODES.NETWORK_CONNECTION_FAILED]: 'ネットワーク接続に失敗しました。インターネット接続を確認してください。',
  [ERROR_CODES.NETWORK_TIMEOUT]: 'ネットワーク接続がタイムアウトしました。',
  [ERROR_CODES.NETWORK_OFFLINE]: 'オフライン状態です。インターネット接続を確認してください。',
  [ERROR_CODES.NETWORK_CORS_ERROR]: 'CORS制限により接続できませんでした。',

  // バリデーション関連メッセージ
  [ERROR_CODES.VALIDATION_REQUIRED_FIELD]: '必須項目が入力されていません。',
  [ERROR_CODES.VALIDATION_INVALID_FORMAT]: '入力形式が正しくありません。',
  [ERROR_CODES.VALIDATION_LENGTH_EXCEEDED]: '入力が最大長を超えています。',
  [ERROR_CODES.VALIDATION_INVALID_VALUE]: '無効な値が入力されています。',

  // 権限関連メッセージ
  [ERROR_CODES.PERMISSION_DENIED]: '権限が不足しています。',
  [ERROR_CODES.PERMISSION_STORAGE_ACCESS]: 'ストレージへのアクセス権限がありません。',
  [ERROR_CODES.PERMISSION_HOST_ACCESS]: 'このサイトへのアクセス権限がありません。',

  // ストレージ関連メッセージ
  [ERROR_CODES.STORAGE_QUOTA_EXCEEDED]: 'ストレージ容量が不足しています。',
  [ERROR_CODES.STORAGE_ACCESS_DENIED]: 'ストレージにアクセスできません。',
  [ERROR_CODES.STORAGE_CORRUPTION]: 'ストレージデータが破損しています。',

  // UI関連メッセージ
  [ERROR_CODES.UI_ELEMENT_NOT_FOUND]: '必要なUI要素が見つかりません。',
  [ERROR_CODES.UI_RENDER_FAILED]: 'UI表示に失敗しました。',
  [ERROR_CODES.UI_INTERACTION_FAILED]: 'UI操作に失敗しました。',

  // 解析関連メッセージ
  [ERROR_CODES.PARSING_CONTENT_FAILED]: 'コンテンツの解析に失敗しました。',
  [ERROR_CODES.PARSING_EMAIL_FAILED]: 'メールの解析に失敗しました。',
  [ERROR_CODES.PARSING_PAGE_FAILED]: 'ページの解析に失敗しました。',

  // システム関連メッセージ
  [ERROR_CODES.SYSTEM_INITIALIZATION_FAILED]: 'システムの初期化に失敗しました。',
  [ERROR_CODES.SYSTEM_SERVICE_UNAVAILABLE]: 'サービスが利用できません。',
  [ERROR_CODES.SYSTEM_CONFIGURATION_ERROR]: '設定にエラーがあります。',
  [ERROR_CODES.SYSTEM_UNKNOWN_ERROR]: '予期しないエラーが発生しました。'
});

/**
 * エラー重要度レベル
 */
export const ERROR_SEVERITY = Object.freeze({
  LOW: 'low',
  MEDIUM: 'medium',
  HIGH: 'high',
  CRITICAL: 'critical'
});

/**
 * エラー重要度マッピング
 */
export const ERROR_SEVERITY_MAP = Object.freeze({
  [ERROR_CODES.API_INVALID_KEY]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.API_ACCESS_DENIED]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.API_ENDPOINT_NOT_FOUND]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.API_RATE_LIMIT_EXCEEDED]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.API_SERVER_ERROR]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.API_INVALID_RESPONSE]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.API_TIMEOUT]: ERROR_SEVERITY.LOW,

  [ERROR_CODES.NETWORK_CONNECTION_FAILED]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.NETWORK_TIMEOUT]: ERROR_SEVERITY.LOW,
  [ERROR_CODES.NETWORK_OFFLINE]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.NETWORK_CORS_ERROR]: ERROR_SEVERITY.MEDIUM,

  [ERROR_CODES.VALIDATION_REQUIRED_FIELD]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.VALIDATION_INVALID_FORMAT]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.VALIDATION_LENGTH_EXCEEDED]: ERROR_SEVERITY.LOW,
  [ERROR_CODES.VALIDATION_INVALID_VALUE]: ERROR_SEVERITY.MEDIUM,

  [ERROR_CODES.PERMISSION_DENIED]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.PERMISSION_STORAGE_ACCESS]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.PERMISSION_HOST_ACCESS]: ERROR_SEVERITY.HIGH,

  [ERROR_CODES.STORAGE_QUOTA_EXCEEDED]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.STORAGE_ACCESS_DENIED]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.STORAGE_CORRUPTION]: ERROR_SEVERITY.CRITICAL,

  [ERROR_CODES.UI_ELEMENT_NOT_FOUND]: ERROR_SEVERITY.LOW,
  [ERROR_CODES.UI_RENDER_FAILED]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.UI_INTERACTION_FAILED]: ERROR_SEVERITY.LOW,

  [ERROR_CODES.PARSING_CONTENT_FAILED]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.PARSING_EMAIL_FAILED]: ERROR_SEVERITY.MEDIUM,
  [ERROR_CODES.PARSING_PAGE_FAILED]: ERROR_SEVERITY.MEDIUM,

  [ERROR_CODES.SYSTEM_INITIALIZATION_FAILED]: ERROR_SEVERITY.CRITICAL,
  [ERROR_CODES.SYSTEM_SERVICE_UNAVAILABLE]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.SYSTEM_CONFIGURATION_ERROR]: ERROR_SEVERITY.HIGH,
  [ERROR_CODES.SYSTEM_UNKNOWN_ERROR]: ERROR_SEVERITY.MEDIUM
});

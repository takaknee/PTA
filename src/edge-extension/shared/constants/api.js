/*
 * AI業務支援ツール Edge拡張機能 - API定数
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * APIプロバイダー定数
 */
export const API_PROVIDERS = Object.freeze({
  OPENAI: 'openai',
  AZURE: 'azure'
});

/**
 * サポートされるAIモデル定数
 */
export const AI_MODELS = Object.freeze({
  GPT_4O_MINI: 'gpt-4o-mini',
  GPT_4O: 'gpt-4o',
  GPT_35_TURBO: 'gpt-3.5-turbo',
  GPT_4: 'gpt-4'
});

/**
 * APIエンドポイント定数
 */
export const API_ENDPOINTS = Object.freeze({
  OPENAI: 'https://api.openai.com/v1/chat/completions',
  AZURE_TEMPLATE: 'https://{resource}.openai.azure.com/openai/deployments/{deployment}/chat/completions?api-version=2024-02-15-preview'
});

/**
 * HTTPステータスコード定数
 */
export const HTTP_STATUS = Object.freeze({
  OK: 200,
  BAD_REQUEST: 400,
  UNAUTHORIZED: 401,
  FORBIDDEN: 403,
  NOT_FOUND: 404,
  TOO_MANY_REQUESTS: 429,
  INTERNAL_SERVER_ERROR: 500,
  BAD_GATEWAY: 502,
  SERVICE_UNAVAILABLE: 503,
  GATEWAY_TIMEOUT: 504
});

/**
 * APIエラーコード定数
 */
export const API_ERROR_CODES = Object.freeze({
  INVALID_API_KEY: 'invalid_api_key',
  ACCESS_DENIED: 'access_denied',
  ENDPOINT_NOT_FOUND: 'endpoint_not_found',
  RATE_LIMIT_EXCEEDED: 'rate_limit_exceeded',
  SERVER_ERROR: 'server_error',
  NETWORK_ERROR: 'network_error',
  INVALID_RESPONSE: 'invalid_response',
  TIMEOUT: 'timeout',
  UNKNOWN_ERROR: 'unknown_error'
});

/**
 * リクエスト設定定数
 */
export const REQUEST_CONFIG = Object.freeze({
  DEFAULT_TIMEOUT: 30000,
  MAX_RETRIES: 3,
  RETRY_DELAY: 1000,
  MAX_TOKENS: 4000,
  DEFAULT_TEMPERATURE: 0.7
});

/**
 * APIバリデーションルール
 */
export const API_VALIDATION = Object.freeze({
  API_KEY_MIN_LENGTH: 20,
  API_KEY_PATTERN: /^[a-zA-Z0-9\-_\.]+$/,
  AZURE_ENDPOINT_PATTERN: /^https:\/\/[\w\-]+\.openai\.azure\.com/,
  DEPLOYMENT_NAME_PATTERN: /^[a-zA-Z0-9\-_]+$/,
  MAX_INPUT_LENGTH: 10000
});

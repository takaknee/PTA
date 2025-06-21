/*
 * AI業務支援ツール Edge拡張機能 - ユーティリティ関数
 * Copyright (c) 2024 AI Business Support Team
 */

import { SYSTEM_CONFIG, ERROR_MESSAGES, NOTIFICATION_TYPES } from './constants.js';

/**
 * 非同期処理のスリープ
 * @param {number} ms 待機時間（ミリ秒）
 * @returns {Promise} Promise
 */
export function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * デバウンス処理
 * @param {Function} func 実行する関数
 * @param {number} wait 待機時間（ミリ秒）
 * @returns {Function} デバウンス処理された関数
 */
export function debounce(func, wait = SYSTEM_CONFIG.DEBOUNCE_DELAY) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

/**
 * スロットル処理
 * @param {Function} func 実行する関数
 * @param {number} wait 待機時間（ミリ秒）
 * @returns {Function} スロットル処理された関数
 */
export function throttle(func, wait) {
  let inThrottle;
  return function executedFunction(...args) {
    if (!inThrottle) {
      func.apply(this, args);
      inThrottle = true;
      setTimeout(() => inThrottle = false, wait);
    }
  };
}

/**
 * リトライ処理
 * @param {Function} fn 実行する関数
 * @param {number} attempts リトライ回数
 * @param {number} delay 遅延（ミリ秒）
 * @returns {Promise} Promise
 */
export async function retryAsync(fn, attempts = SYSTEM_CONFIG.RETRY_ATTEMPTS, delay = SYSTEM_CONFIG.RETRY_DELAY) {
  try {
    return await fn();
  } catch (error) {
    if (attempts <= 1) {
      throw error;
    }

    console.warn(`リトライ中... 残り${attempts - 1}回`, error.message);
    await sleep(delay);
    return retryAsync(fn, attempts - 1, delay * 2); // 指数バックオフ
  }
}

/**
 * オブジェクトの深いコピー
 * @param {Object} obj コピー対象のオブジェクト
 * @returns {Object} コピーされたオブジェクト
 */
export function deepClone(obj) {
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }

  if (obj instanceof Date) {
    return new Date(obj.getTime());
  }

  if (obj instanceof Array) {
    return obj.map(item => deepClone(item));
  }

  if (typeof obj === 'object') {
    const clonedObj = {};
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        clonedObj[key] = deepClone(obj[key]);
      }
    }
    return clonedObj;
  }
}

/**
 * オブジェクトのマージ（深いマージ）
 * @param {Object} target ターゲットオブジェクト
 * @param {Object} source ソースオブジェクト
 * @returns {Object} マージされたオブジェクト
 */
export function deepMerge(target, source) {
  const result = deepClone(target);

  for (const key in source) {
    if (source.hasOwnProperty(key)) {
      if (source[key] && typeof source[key] === 'object' && !Array.isArray(source[key])) {
        result[key] = deepMerge(result[key] || {}, source[key]);
      } else {
        result[key] = source[key];
      }
    }
  }

  return result;
}

/**
 * 文字列のサニタイズ（HTML特殊文字のエスケープ）
 * @param {string} str サニタイズする文字列
 * @returns {string} サニタイズされた文字列
 */
export function sanitizeHtml(str) {
  if (!str) return '';

  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

/**
 * テキストの省略
 * @param {string} text 対象テキスト
 * @param {number} maxLength 最大長
 * @param {string} suffix 省略記号
 * @returns {string} 省略されたテキスト
 */
export function truncateText(text, maxLength = 100, suffix = '...') {
  if (!text || text.length <= maxLength) {
    return text;
  }

  return text.substring(0, maxLength - suffix.length) + suffix;
}

/**
 * 日時のフォーマット
 * @param {number|Date} timestamp タイムスタンプまたはDateオブジェクト
 * @param {string} format フォーマット（'datetime', 'date', 'time', 'relative'）
 * @returns {string} フォーマットされた日時
 */
export function formatDateTime(timestamp, format = 'datetime') {
  const date = timestamp instanceof Date ? timestamp : new Date(timestamp);

  if (isNaN(date.getTime())) {
    return '無効な日時';
  }

  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffMinutes = Math.floor(diffMs / (1000 * 60));
  const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

  switch (format) {
    case 'relative':
      if (diffMinutes < 1) return 'たった今';
      if (diffMinutes < 60) return `${diffMinutes}分前`;
      if (diffHours < 24) return `${diffHours}時間前`;
      if (diffDays < 7) return `${diffDays}日前`;
      return date.toLocaleDateString('ja-JP');

    case 'date':
      return date.toLocaleDateString('ja-JP');

    case 'time':
      return date.toLocaleTimeString('ja-JP');

    case 'datetime':
    default:
      return date.toLocaleString('ja-JP');
  }
}

/**
 * ファイルサイズのフォーマット
 * @param {number} bytes バイト数
 * @param {number} decimals 小数点以下桁数
 * @returns {string} フォーマットされたファイルサイズ
 */
export function formatFileSize(bytes, decimals = 2) {
  if (bytes === 0) return '0 Bytes';

  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];

  const i = Math.floor(Math.log(bytes) / Math.log(k));

  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

/**
 * URLの妥当性チェック
 * @param {string} url チェックするURL
 * @returns {boolean} 有効なURLの場合true
 */
export function isValidUrl(url) {
  try {
    new URL(url);
    return true;
  } catch {
    return false;
  }
}

/**
 * メールアドレスの妥当性チェック
 * @param {string} email チェックするメールアドレス
 * @returns {boolean} 有効なメールアドレスの場合true
 */
export function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * JSONの安全なパース
 * @param {string} jsonString JSON文字列
 * @param {*} defaultValue デフォルト値
 * @returns {*} パースされたオブジェクトまたはデフォルト値
 */
export function safeJsonParse(jsonString, defaultValue = null) {
  try {
    return JSON.parse(jsonString);
  } catch {
    return defaultValue;
  }
}

/**
 * JSONの安全な文字列化
 * @param {*} obj オブジェクト
 * @param {string} defaultValue デフォルト値
 * @returns {string} JSON文字列またはデフォルト値
 */
export function safeJsonStringify(obj, defaultValue = '{}') {
  try {
    return JSON.stringify(obj);
  } catch {
    return defaultValue;
  }
}

/**
 * ランダムな文字列の生成
 * @param {number} length 長さ
 * @param {string} charset 使用する文字セット
 * @returns {string} ランダム文字列
 */
export function generateRandomString(length = 10, charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789') {
  let result = '';
  for (let i = 0; i < length; i++) {
    result += charset.charAt(Math.floor(Math.random() * charset.length));
  }
  return result;
}

/**
 * UUIDの生成
 * @returns {string} UUID
 */
export function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

/**
 * エラーオブジェクトのシリアライズ
 * @param {Error} error エラーオブジェクト
 * @returns {Object} シリアライズされたエラー情報
 */
export function serializeError(error) {
  return {
    name: error.name,
    message: error.message,
    stack: error.stack,
    timestamp: Date.now()
  };
}

/**
 * HTTPステータスコードからエラーメッセージを取得
 * @param {number} statusCode HTTPステータスコード
 * @returns {string} エラーメッセージ
 */
export function getErrorMessageFromStatus(statusCode) {
  switch (statusCode) {
    case 400:
      return 'リクエストが無効です';
    case 401:
      return ERROR_MESSAGES.INVALID_API_KEY;
    case 403:
      return ERROR_MESSAGES.ACCESS_DENIED;
    case 404:
      return ERROR_MESSAGES.ENDPOINT_NOT_FOUND;
    case 429:
      return ERROR_MESSAGES.RATE_LIMIT_EXCEEDED;
    case 500:
    case 502:
    case 503:
      return ERROR_MESSAGES.SERVER_ERROR;
    default:
      return `HTTPエラー (${statusCode})`;
  }
}

/**
 * コンソールログの改良版（開発環境でのみ動作）
 */
export class Logger {
  static enabled = true;

  static log(message, ...args) {
    if (this.enabled) {
      console.log(`[AI Tool] ${message}`, ...args);
    }
  }

  static warn(message, ...args) {
    if (this.enabled) {
      console.warn(`[AI Tool] ${message}`, ...args);
    }
  }

  static error(message, ...args) {
    if (this.enabled) {
      console.error(`[AI Tool] ${message}`, ...args);
    }
  }

  static info(message, ...args) {
    if (this.enabled) {
      console.info(`[AI Tool] ${message}`, ...args);
    }
  }

  static debug(message, ...args) {
    if (this.enabled) {
      console.debug(`[AI Tool] ${message}`, ...args);
    }
  }
}

/**
 * パフォーマンス測定クラス
 */
export class PerformanceMonitor {
  constructor() {
    this.measurements = new Map();
  }

  /**
   * 測定開始
   * @param {string} name 測定名
   */
  start(name) {
    this.measurements.set(name, performance.now());
  }

  /**
   * 測定終了と結果取得
   * @param {string} name 測定名
   * @returns {number} 経過時間（ミリ秒）
   */
  end(name) {
    const startTime = this.measurements.get(name);
    if (startTime === undefined) {
      Logger.warn(`測定 '${name}' が開始されていません`);
      return 0;
    }

    const duration = performance.now() - startTime;
    this.measurements.delete(name);
    return duration;
  }

  /**
   * 測定中の項目をクリア
   */
  clear() {
    this.measurements.clear();
  }
}

/**
 * グローバルパフォーマンスモニター
 */
export const performanceMonitor = new PerformanceMonitor();

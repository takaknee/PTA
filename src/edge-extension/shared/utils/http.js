/*
 * AI業務支援ツール Edge拡張機能 - HTTPユーティリティ
 * Copyright (c) 2024 AI Business Support Team
 */

import { HTTP_STATUS, REQUEST_CONFIG, ERROR_CODES } from '../constants/index.js';
import { createLogger } from './logger.js';

const logger = createLogger('HTTPUtils');

/**
 * HTTPエラークラス
 */
export class HTTPError extends Error {
  constructor(status, message, response = null) {
    super(message);
    this.name = 'HTTPError';
    this.status = status;
    this.response = response;
  }
}

/**
 * HTTPリクエストユーティリティクラス
 */
export class HTTPUtils {
  /**
   * 基本的なFetchリクエスト
   * @param {string} url - リクエストURL
   * @param {Object} options - リクエストオプション
   * @returns {Promise<Response>} レスポンス
   */
  static async fetch(url, options = {}) {
    const defaultOptions = {
      timeout: REQUEST_CONFIG.DEFAULT_TIMEOUT,
      retries: REQUEST_CONFIG.MAX_RETRIES,
      retryDelay: REQUEST_CONFIG.RETRY_DELAY
    };

    const config = { ...defaultOptions, ...options };

    return this.fetchWithRetry(url, config);
  }

  /**
   * リトライ機能付きFetch
   * @param {string} url - リクエストURL
   * @param {Object} options - リクエストオプション
   * @returns {Promise<Response>} レスポンス
   */
  static async fetchWithRetry(url, options) {
    const { retries, retryDelay, timeout, ...fetchOptions } = options;

    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        logger.debug(`HTTP request attempt ${attempt + 1}`, { url, attempt });

        const response = await this.fetchWithTimeout(url, fetchOptions, timeout);

        // 成功レスポンスまたはリトライ不要なエラー
        if (response.ok || !this.shouldRetry(response.status)) {
          return response;
        }

        // 最後の試行の場合はエラーを投げる
        if (attempt === retries) {
          throw new HTTPError(
            response.status,
            `HTTP request failed: ${response.status} ${response.statusText}`,
            response
          );
        }

      } catch (error) {
        logger.warn(`HTTP request failed on attempt ${attempt + 1}`, {
          url,
          attempt,
          error: error.message
        });

        // 最後の試行またはリトライ不可能なエラー
        if (attempt === retries || !this.shouldRetryOnError(error)) {
          throw error;
        }
      }

      // リトライ前の待機
      if (attempt < retries) {
        await this.delay(retryDelay * Math.pow(2, attempt));
      }
    }
  }

  /**
   * タイムアウト付きFetch
   * @param {string} url - リクエストURL
   * @param {Object} options - リクエストオプション
   * @param {number} timeout - タイムアウト時間
   * @returns {Promise<Response>} レスポンス
   */
  static async fetchWithTimeout(url, options, timeout) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeout);

    try {
      const response = await fetch(url, {
        ...options,
        signal: controller.signal
      });

      clearTimeout(timeoutId);
      return response;

    } catch (error) {
      clearTimeout(timeoutId);

      if (error.name === 'AbortError') {
        throw new HTTPError(408, 'Request timeout');
      }

      throw error;
    }
  }

  /**
   * JSONリクエスト
   * @param {string} url - リクエストURL
   * @param {Object} options - リクエストオプション
   * @returns {Promise<Object>} JSONレスポンス
   */
  static async fetchJSON(url, options = {}) {
    const response = await this.fetch(url, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        ...options.headers
      }
    });

    if (!response.ok) {
      throw new HTTPError(
        response.status,
        `HTTP ${response.status}: ${response.statusText}`,
        response
      );
    }

    try {
      return await response.json();
    } catch (error) {
      throw new HTTPError(
        response.status,
        'Invalid JSON response',
        response
      );
    }
  }

  /**
   * POSTリクエスト
   * @param {string} url - リクエストURL
   * @param {Object} data - 送信データ
   * @param {Object} options - リクエストオプション
   * @returns {Promise<Object>} レスポンス
   */
  static async post(url, data, options = {}) {
    return this.fetchJSON(url, {
      ...options,
      method: 'POST',
      body: JSON.stringify(data)
    });
  }

  /**
   * PUTリクエスト
   * @param {string} url - リクエストURL
   * @param {Object} data - 送信データ
   * @param {Object} options - リクエストオプション
   * @returns {Promise<Object>} レスポンス
   */
  static async put(url, data, options = {}) {
    return this.fetchJSON(url, {
      ...options,
      method: 'PUT',
      body: JSON.stringify(data)
    });
  }

  /**
   * DELETEリクエスト
   * @param {string} url - リクエストURL
   * @param {Object} options - リクエストオプション
   * @returns {Promise<Object>} レスポンス
   */
  static async delete(url, options = {}) {
    return this.fetchJSON(url, {
      ...options,
      method: 'DELETE'
    });
  }

  /**
   * リトライ判定（ステータスコード）
   * @param {number} status - HTTPステータスコード
   * @returns {boolean} リトライ可能フラグ
   */
  static shouldRetry(status) {
    return status >= 500 || status === 429 || status === 408;
  }

  /**
   * リトライ判定（エラー）
   * @param {Error} error - エラーオブジェクト
   * @returns {boolean} リトライ可能フラグ
   */
  static shouldRetryOnError(error) {
    // ネットワークエラーやタイムアウトはリトライ可能
    return error.name === 'TypeError' ||
      error.name === 'AbortError' ||
      error.message.includes('network') ||
      error.message.includes('timeout');
  }

  /**
   * 遅延処理
   * @param {number} ms - 遅延時間（ミリ秒）
   * @returns {Promise} 遅延Promise
   */
  static delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * ステータスコードからエラーコードへの変換
   * @param {number} status - HTTPステータスコード
   * @returns {string} エラーコード
   */
  static statusToErrorCode(status) {
    switch (status) {
      case HTTP_STATUS.BAD_REQUEST:
        return ERROR_CODES.VALIDATION_INVALID_FORMAT;
      case HTTP_STATUS.UNAUTHORIZED:
        return ERROR_CODES.API_INVALID_KEY;
      case HTTP_STATUS.FORBIDDEN:
        return ERROR_CODES.API_ACCESS_DENIED;
      case HTTP_STATUS.NOT_FOUND:
        return ERROR_CODES.API_ENDPOINT_NOT_FOUND;
      case HTTP_STATUS.TOO_MANY_REQUESTS:
        return ERROR_CODES.API_RATE_LIMIT_EXCEEDED;
      case HTTP_STATUS.INTERNAL_SERVER_ERROR:
      case HTTP_STATUS.BAD_GATEWAY:
      case HTTP_STATUS.SERVICE_UNAVAILABLE:
      case HTTP_STATUS.GATEWAY_TIMEOUT:
        return ERROR_CODES.API_SERVER_ERROR;
      default:
        return ERROR_CODES.SYSTEM_UNKNOWN_ERROR;
    }
  }

  /**
   * ベース64エンコード
   * @param {string} str - エンコード対象文字列
   * @returns {string} エンコード結果
   */
  static base64Encode(str) {
    try {
      return btoa(unescape(encodeURIComponent(str)));
    } catch (error) {
      logger.error('Base64エンコードエラー', { error: error.message });
      throw new Error('Base64エンコードに失敗しました');
    }
  }

  /**
   * ベース64デコード
   * @param {string} str - デコード対象文字列
   * @returns {string} デコード結果
   */
  static base64Decode(str) {
    try {
      return decodeURIComponent(escape(atob(str)));
    } catch (error) {
      logger.error('Base64デコードエラー', { error: error.message });
      throw new Error('Base64デコードに失敗しました');
    }
  }

  /**
   * URLパラメータの生成
   * @param {Object} params - パラメータオブジェクト
   * @returns {string} URLパラメータ文字列
   */
  static buildQueryString(params) {
    const searchParams = new URLSearchParams();

    Object.entries(params).forEach(([key, value]) => {
      if (value !== null && value !== undefined) {
        searchParams.append(key, String(value));
      }
    });

    return searchParams.toString();
  }

  /**
   * Content-Typeの取得
   * @param {Response} response - レスポンスオブジェクト
   * @returns {string} Content-Type
   */
  static getContentType(response) {
    return response.headers.get('content-type') || '';
  }

  /**
   * JSONレスポンス判定
   * @param {Response} response - レスポンスオブジェクト
   * @returns {boolean} JSON判定結果
   */
  static isJSONResponse(response) {
    const contentType = this.getContentType(response);
    return contentType.includes('application/json');
  }
}

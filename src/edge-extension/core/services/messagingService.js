/*
 * AI業務支援ツール Edge拡張機能 - メッセージングサービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';
import { APP_CONSTANTS } from '../../shared/constants/app.js';

/**
 * Chrome拡張機能間メッセージングサービス
 */
export class MessagingService {
  /**
   * コンストラクタ
   */
  constructor() {
    this.logger = new Logger('MessagingService');
    this.validator = new Validator();
    this.messageHandlers = new Map();
    this.setupMessageListeners();
  }

  /**
   * メッセージリスナーを設定
   * @private
   */
  setupMessageListeners() {
    try {
      // Chrome runtime メッセージリスナー
      if (chrome?.runtime?.onMessage) {
        chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
          this.handleMessage(message, sender, sendResponse);
          return true; // 非同期レスポンスを許可
        });
      }

      this.logger.debug('メッセージリスナーが設定されました');
    } catch (error) {
      this.logger.error('メッセージリスナー設定中にエラーが発生しました', error);
    }
  }

  /**
   * メッセージを送信
   * @param {string} target - 送信先
   * @param {string} action - アクション名
   * @param {Object} data - データ
   * @param {Object} options - オプション
   * @returns {Promise<Object>} レスポンス
   */
  async sendMessage(target, action, data = {}, options = {}) {
    try {
      this.logger.debug('メッセージ送信を開始します', { target, action });

      // メッセージの妥当性検証
      const validationResult = this.validateMessage({ target, action, data });
      if (!validationResult.isValid) {
        throw new Error(`メッセージが無効です: ${validationResult.errors.join(', ')}`);
      }

      const message = {
        target,
        action,
        data,
        id: this.generateMessageId(),
        timestamp: new Date().toISOString(),
        source: options.source || 'unknown'
      };

      let response;

      switch (target) {
        case 'background':
          response = await this.sendToBackground(message, options);
          break;
        case 'content':
          response = await this.sendToContent(message, options);
          break;
        case 'popup':
          response = await this.sendToPopup(message, options);
          break;
        case 'options':
          response = await this.sendToOptions(message, options);
          break;
        case 'offscreen':
          response = await this.sendToOffscreen(message, options);
          break;
        default:
          response = await this.sendToRuntime(message, options);
          break;
      }

      this.logger.debug('メッセージ送信が完了しました', { target, action, success: !!response });
      return response || {};
    } catch (error) {
      this.logger.error('メッセージ送信中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * メッセージハンドラーを登録
   * @param {string} action - アクション名
   * @param {Function} handler - ハンドラー関数
   */
  registerHandler(action, handler) {
    try {
      if (!this.validator.isString(action) || action.trim().length === 0) {
        throw new Error('アクション名が無効です');
      }

      if (!this.validator.isFunction(handler)) {
        throw new Error('ハンドラーが無効です');
      }

      this.messageHandlers.set(action, handler);
      this.logger.debug('メッセージハンドラーが登録されました', { action });
    } catch (error) {
      this.logger.error('メッセージハンドラー登録中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * メッセージハンドラーを解除
   * @param {string} action - アクション名
   */
  unregisterHandler(action) {
    try {
      if (this.messageHandlers.has(action)) {
        this.messageHandlers.delete(action);
        this.logger.debug('メッセージハンドラーが解除されました', { action });
      }
    } catch (error) {
      this.logger.error('メッセージハンドラー解除中にエラーが発生しました', error);
    }
  }

  /**
   * メッセージを処理
   * @param {Object} message - メッセージ
   * @param {Object} sender - 送信者情報
   * @param {Function} sendResponse - レスポンス送信関数
   * @private
   */
  async handleMessage(message, sender, sendResponse) {
    try {
      this.logger.debug('メッセージを受信しました', {
        action: message.action,
        target: message.target,
        id: message.id
      });

      // メッセージの妥当性検証
      const validationResult = this.validateMessage(message);
      if (!validationResult.isValid) {
        const errorResponse = {
          success: false,
          error: `無効なメッセージです: ${validationResult.errors.join(', ')}`
        };
        sendResponse(errorResponse);
        return;
      }

      // 該当するハンドラーを検索
      const handler = this.messageHandlers.get(message.action);
      if (!handler) {
        this.logger.warn('メッセージハンドラーが見つかりません', { action: message.action });
        sendResponse({
          success: false,
          error: `未知のアクションです: ${message.action}`
        });
        return;
      }

      // ハンドラーを実行
      const result = await handler(message.data, sender, message);

      const response = {
        success: true,
        data: result,
        timestamp: new Date().toISOString(),
        messageId: message.id
      };

      this.logger.debug('メッセージ処理が完了しました', { action: message.action });
      sendResponse(response);
    } catch (error) {
      this.logger.error('メッセージ処理中にエラーが発生しました', error);
      sendResponse({
        success: false,
        error: error.message,
        timestamp: new Date().toISOString(),
        messageId: message.id
      });
    }
  }

  /**
   * バックグラウンドスクリプトにメッセージを送信
   * @param {Object} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async sendToBackground(message, options) {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error('バックグラウンドスクリプトへのメッセージ送信がタイムアウトしました'));
      }, options.timeout || APP_CONSTANTS.MESSAGE_TIMEOUT);

      chrome.runtime.sendMessage(message, (response) => {
        clearTimeout(timeout);
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(response);
        }
      });
    });
  }

  /**
   * コンテンツスクリプトにメッセージを送信
   * @param {Object} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async sendToContent(message, options) {
    return new Promise(async (resolve, reject) => {
      try {
        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (tabs.length === 0) {
          reject(new Error('アクティブなタブが見つかりません'));
          return;
        }

        const timeout = setTimeout(() => {
          reject(new Error('コンテンツスクリプトへのメッセージ送信がタイムアウトしました'));
        }, options.timeout || APP_CONSTANTS.MESSAGE_TIMEOUT);

        chrome.tabs.sendMessage(tabs[0].id, message, (response) => {
          clearTimeout(timeout);
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
          } else {
            resolve(response);
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * ポップアップにメッセージを送信
   * @param {Object} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async sendToPopup(message, options) {
    return this.sendToBackground(message, options);
  }

  /**
   * オプションページにメッセージを送信
   * @param {Object} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async sendToOptions(message, options) {
    return this.sendToBackground(message, options);
  }

  /**
   * オフスクリーンドキュメントにメッセージを送信
   * @param {Object} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async sendToOffscreen(message, options) {
    return this.sendToBackground(message, options);
  }

  /**
   * ランタイムにメッセージを送信
   * @param {Object} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<Object>} レスポンス
   * @private
   */
  async sendToRuntime(message, options) {
    return this.sendToBackground(message, options);
  }

  /**
   * メッセージの妥当性を検証
   * @param {Object} message - メッセージ
   * @returns {Object} 検証結果
   * @private
   */
  validateMessage(message) {
    const errors = [];

    try {
      // 必須フィールドの確認
      if (!this.validator.isString(message.target) || message.target.trim().length === 0) {
        errors.push('ターゲットが無効です');
      }

      if (!this.validator.isString(message.action) || message.action.trim().length === 0) {
        errors.push('アクションが無効です');
      }

      if (message.data !== undefined && !this.validator.isObject(message.data)) {
        errors.push('データの形式が無効です');
      }

      // サイズ制限の確認
      const messageSize = JSON.stringify(message).length;
      if (messageSize > APP_CONSTANTS.MAX_MESSAGE_SIZE) {
        errors.push(`メッセージサイズが上限を超えています（${messageSize} > ${APP_CONSTANTS.MAX_MESSAGE_SIZE}）`);
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('メッセージ妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['メッセージの妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * メッセージIDを生成
   * @returns {string} メッセージID
   * @private
   */
  generateMessageId() {
    return `msg_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * 全てのハンドラーを解除
   */
  clearAllHandlers() {
    try {
      this.messageHandlers.clear();
      this.logger.debug('全てのメッセージハンドラーが解除されました');
    } catch (error) {
      this.logger.error('メッセージハンドラー全解除中にエラーが発生しました', error);
    }
  }

  /**
   * 登録済みハンドラー一覧を取得
   * @returns {Array} ハンドラー一覧
   */
  getRegisteredHandlers() {
    try {
      return Array.from(this.messageHandlers.keys());
    } catch (error) {
      this.logger.error('ハンドラー一覧取得中にエラーが発生しました', error);
      return [];
    }
  }

  /**
   * ブロードキャストメッセージを送信
   * @param {string} action - アクション名
   * @param {Object} data - データ
   * @returns {Promise<Array>} レスポンス一覧
   */
  async broadcast(action, data = {}) {
    try {
      this.logger.debug('ブロードキャスト送信を開始します', { action });

      const targets = ['background', 'content', 'popup', 'options'];
      const promises = targets.map(target =>
        this.sendMessage(target, action, data, { timeout: 1000 })
          .catch(error => ({ error: error.message, target }))
      );

      const responses = await Promise.all(promises);

      this.logger.debug('ブロードキャスト送信が完了しました', { action, responseCount: responses.length });
      return responses;
    } catch (error) {
      this.logger.error('ブロードキャスト送信中にエラーが発生しました', error);
      throw error;
    }
  }
}

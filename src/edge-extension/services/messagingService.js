/*
 * AI業務支援ツール Edge拡張機能 - メッセージングサービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { Message } from '../lib/models.js';
import { MESSAGE_ACTIONS, MESSAGE_TARGETS } from '../lib/constants.js';
import { Logger, generateUUID } from '../lib/utils.js';

/**
 * Chrome拡張機能コンポーネント間のメッセージング管理サービス
 */
export class MessagingService {
  constructor() {
    this.pendingRequests = new Map();
    this.messageHandlers = new Map();
    this.responseTimeout = 30000; // 30秒
    this.isInitialized = false;
  }

  /**
   * メッセージングサービスの初期化
   */
  initialize() {
    if (this.isInitialized) {
      return;
    }

    // Chrome runtime message listener を設定
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
      this.handleIncomingMessage(message, sender, sendResponse);
      return true; // 非同期レスポンスを示す
    });

    this.isInitialized = true;
    Logger.info('メッセージングサービス初期化完了');
  }

  /**
   * メッセージを送信
   * @param {string} target 送信先
   * @param {string} action アクション
   * @param {Object} data データ
   * @param {Object} options オプション
   * @returns {Promise<*>} レスポンス
   */
  async sendMessage(target, action, data = {}, options = {}) {
    const {
      timeout = this.responseTimeout,
      retries = 1,
      tabId = null
    } = options;

    const message = new Message({
      target,
      action,
      data,
      sender: this.getCurrentContext()
    });

    Logger.debug('メッセージ送信', {
      target,
      action,
      requestId: message.requestId
    });

    try {
      return await this.sendMessageWithRetry(message, { timeout, retries, tabId });
    } catch (error) {
      Logger.error('メッセージ送信エラー', {
        target,
        action,
        error: error.message
      });
      throw error;
    }
  }

  /**
   * リトライ機能付きメッセージ送信
   * @param {Message} message メッセージ
   * @param {Object} options オプション
   * @returns {Promise<*>} レスポンス
   */
  async sendMessageWithRetry(message, options) {
    const { timeout, retries, tabId } = options;
    let lastError;

    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        return await this.sendSingleMessage(message, timeout, tabId);
      } catch (error) {
        lastError = error;

        if (attempt < retries) {
          Logger.warn(`メッセージ送信リトライ ${attempt + 1}/${retries}`, {
            requestId: message.requestId,
            error: error.message
          });

          // リトライ前の待機（指数バックオフ）
          await this.sleep(Math.pow(2, attempt) * 1000);
        }
      }
    }

    throw lastError;
  }

  /**
   * 単一メッセージ送信
   * @param {Message} message メッセージ
   * @param {number} timeout タイムアウト時間
   * @param {number} tabId タブID
   * @returns {Promise<*>} レスポンス
   */
  async sendSingleMessage(message, timeout, tabId) {
    return new Promise((resolve, reject) => {
      // タイムアウト設定
      const timeoutId = setTimeout(() => {
        this.pendingRequests.delete(message.requestId);
        reject(new Error(`メッセージタイムアウト: ${message.target}/${message.action}`));
      }, timeout);

      // レスポンス待機の登録
      this.pendingRequests.set(message.requestId, {
        resolve,
        reject,
        timeoutId,
        timestamp: Date.now()
      });

      try {
        // メッセージ送信
        if (tabId) {
          // 特定のタブに送信
          chrome.tabs.sendMessage(tabId, message.toJSON(), (response) => {
            this.handleMessageResponse(message.requestId, response);
          });
        } else {
          // Runtime（background script）に送信
          chrome.runtime.sendMessage(message.toJSON(), (response) => {
            this.handleMessageResponse(message.requestId, response);
          });
        }
      } catch (error) {
        // 送信エラー
        clearTimeout(timeoutId);
        this.pendingRequests.delete(message.requestId);
        reject(error);
      }
    });
  }

  /**
   * メッセージレスポンスを処理
   * @param {string} requestId リクエストID
   * @param {*} response レスポンス
   */
  handleMessageResponse(requestId, response) {
    const pendingRequest = this.pendingRequests.get(requestId);

    if (!pendingRequest) {
      Logger.warn('対応するリクエストが見つかりません', { requestId });
      return;
    }

    clearTimeout(pendingRequest.timeoutId);
    this.pendingRequests.delete(requestId);

    // Chrome runtime のエラーチェック
    if (chrome.runtime.lastError) {
      pendingRequest.reject(new Error(chrome.runtime.lastError.message));
      return;
    }

    // レスポンスの処理
    if (response && response.error) {
      pendingRequest.reject(new Error(response.error));
    } else {
      pendingRequest.resolve(response);
    }
  }

  /**
   * 受信メッセージを処理
   * @param {Object} message 受信メッセージ
   * @param {Object} sender 送信者情報
   * @param {Function} sendResponse レスポンス関数
   */
  async handleIncomingMessage(message, sender, sendResponse) {
    try {
      Logger.debug('メッセージ受信', {
        action: message.action,
        target: message.target,
        sender: sender.tab ? 'content' : 'extension'
      });

      // 自分宛てでないメッセージは無視
      const currentContext = this.getCurrentContext();
      if (message.target && message.target !== currentContext) {
        return;
      }

      // メッセージハンドラーの実行
      const handler = this.messageHandlers.get(message.action);
      if (handler) {
        const response = await handler(message.data, sender);
        sendResponse({ success: true, data: response });
      } else {
        Logger.warn('未処理のメッセージアクション', { action: message.action });
        sendResponse({ success: false, error: '未知のアクション' });
      }

    } catch (error) {
      Logger.error('メッセージ処理エラー', error);
      sendResponse({ success: false, error: error.message });
    }
  }

  /**
   * メッセージハンドラーを登録
   * @param {string} action アクション名
   * @param {Function} handler ハンドラー関数
   */
  registerHandler(action, handler) {
    if (typeof handler !== 'function') {
      throw new Error('ハンドラーは関数である必要があります');
    }

    this.messageHandlers.set(action, handler);
    Logger.debug('メッセージハンドラー登録', { action });
  }

  /**
   * メッセージハンドラーを削除
   * @param {string} action アクション名
   */
  unregisterHandler(action) {
    this.messageHandlers.delete(action);
    Logger.debug('メッセージハンドラー削除', { action });
  }

  /**
   * 複数のハンドラーを一括登録
   * @param {Object} handlers ハンドラーオブジェクト
   */
  registerHandlers(handlers) {
    Object.entries(handlers).forEach(([action, handler]) => {
      this.registerHandler(action, handler);
    });
  }

  /**
   * アクティブなタブに メッセージを送信
   * @param {string} action アクション
   * @param {Object} data データ
   * @returns {Promise<*>} レスポンス
   */
  async sendToActiveTab(action, data = {}) {
    try {
      // アクティブなタブを取得
      const [activeTab] = await chrome.tabs.query({ active: true, currentWindow: true });

      if (!activeTab) {
        throw new Error('アクティブなタブが見つかりません');
      }

      return await this.sendMessage(MESSAGE_TARGETS.CONTENT, action, data, {
        tabId: activeTab.id
      });

    } catch (error) {
      Logger.error('アクティブタブへのメッセージ送信エラー', error);
      throw error;
    }
  }

  /**
   * すべてのタブにブロードキャスト
   * @param {string} action アクション
   * @param {Object} data データ
   * @param {Object} filter タブフィルター
   * @returns {Promise<Array>} 全レスポンス
   */
  async broadcastToTabs(action, data = {}, filter = {}) {
    try {
      const tabs = await chrome.tabs.query(filter);
      const promises = tabs.map(tab =>
        this.sendMessage(MESSAGE_TARGETS.CONTENT, action, data, {
          tabId: tab.id,
          timeout: 5000 // ブロードキャストは短いタイムアウト
        }).catch(error => ({ error: error.message, tabId: tab.id }))
      );

      return await Promise.all(promises);

    } catch (error) {
      Logger.error('タブブロードキャストエラー', error);
      throw error;
    }
  }

  /**
   * Background scriptに メッセージを送信
   * @param {string} action アクション
   * @param {Object} data データ
   * @returns {Promise<*>} レスポンス
   */
  async sendToBackground(action, data = {}) {
    return await this.sendMessage(MESSAGE_TARGETS.BACKGROUND, action, data);
  }

  /**
   * Offscreen documentに メッセージを送信
   * @param {string} action アクション
   * @param {Object} data データ
   * @returns {Promise<*>} レスポンス
   */
  async sendToOffscreen(action, data = {}) {
    return await this.sendMessage(MESSAGE_TARGETS.OFFSCREEN, action, data);
  }

  /**
   * 現在のコンテキストを取得
   * @returns {string} コンテキスト名
   */
  getCurrentContext() {
    if (typeof chrome !== 'undefined') {
      if (chrome.runtime && chrome.runtime.getManifest) {
        // Extension context
        if (chrome.tabs) {
          return MESSAGE_TARGETS.BACKGROUND;
        }
        if (window.location && window.location.href.includes('popup.html')) {
          return MESSAGE_TARGETS.POPUP;
        }
        if (window.location && window.location.href.includes('options.html')) {
          return MESSAGE_TARGETS.OPTIONS;
        }
        if (window.location && window.location.href.includes('offscreen.html')) {
          return MESSAGE_TARGETS.OFFSCREEN;
        }
      } else {
        // Content script context
        return MESSAGE_TARGETS.CONTENT;
      }
    }

    return 'unknown';
  }

  /**
   * 保留中のリクエストをクリーンアップ
   * @param {number} maxAge 最大保持時間（ms）
   */
  cleanupPendingRequests(maxAge = 300000) { // 5分
    const now = Date.now();

    for (const [requestId, request] of this.pendingRequests.entries()) {
      if (now - request.timestamp > maxAge) {
        clearTimeout(request.timeoutId);
        request.reject(new Error('リクエストタイムアウト（クリーンアップ）'));
        this.pendingRequests.delete(requestId);

        Logger.warn('保留中リクエストをクリーンアップ', { requestId });
      }
    }
  }

  /**
   * メッセージング統計を取得
   * @returns {Object} 統計情報
   */
  getStatistics() {
    return {
      pendingRequests: this.pendingRequests.size,
      registeredHandlers: this.messageHandlers.size,
      handlerActions: Array.from(this.messageHandlers.keys()),
      isInitialized: this.isInitialized,
      currentContext: this.getCurrentContext()
    };
  }

  /**
   * サービスをクリーンアップ
   */
  cleanup() {
    // 保留中のリクエストをすべてキャンセル
    for (const [requestId, request] of this.pendingRequests.entries()) {
      clearTimeout(request.timeoutId);
      request.reject(new Error('メッセージングサービス停止'));
    }

    this.pendingRequests.clear();
    this.messageHandlers.clear();
    this.isInitialized = false;

    Logger.info('メッセージングサービスクリーンアップ完了');
  }

  /**
   * スリープ関数
   * @param {number} ms 待機時間（ミリ秒）
   * @returns {Promise} Promise
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

/**
 * グローバルメッセージングサービスインスタンス
 */
export const messagingService = new MessagingService();

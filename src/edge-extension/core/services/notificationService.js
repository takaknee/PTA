/*
 * AI業務支援ツール Edge拡張機能 - 通知サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';
import { UI_CONSTANTS } from '../../shared/constants/ui.js';

/**
 * 通知管理サービス
 */
export class NotificationService {
  /**
   * コンストラクタ
   */
  constructor() {
    this.logger = new Logger('NotificationService');
    this.validator = new Validator();
    this.activeNotifications = new Map();
    this.notificationQueue = [];
    this.isProcessingQueue = false;
  }

  /**
   * 成功通知を表示
   * @param {string} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<string>} 通知ID
   */
  async showSuccess(message, options = {}) {
    return this.showNotification({
      type: 'success',
      message,
      title: options.title || '成功',
      duration: options.duration || UI_CONSTANTS.NOTIFICATION_DURATION.SUCCESS,
      ...options
    });
  }

  /**
   * エラー通知を表示
   * @param {string} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<string>} 通知ID
   */
  async showError(message, options = {}) {
    return this.showNotification({
      type: 'error',
      message,
      title: options.title || 'エラー',
      duration: options.duration || UI_CONSTANTS.NOTIFICATION_DURATION.ERROR,
      persistent: options.persistent !== undefined ? options.persistent : true,
      ...options
    });
  }

  /**
   * 警告通知を表示
   * @param {string} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<string>} 通知ID
   */
  async showWarning(message, options = {}) {
    return this.showNotification({
      type: 'warning',
      message,
      title: options.title || '警告',
      duration: options.duration || UI_CONSTANTS.NOTIFICATION_DURATION.WARNING,
      ...options
    });
  }

  /**
   * 情報通知を表示
   * @param {string} message - メッセージ
   * @param {Object} options - オプション
   * @returns {Promise<string>} 通知ID
   */
  async showInfo(message, options = {}) {
    return this.showNotification({
      type: 'info',
      message,
      title: options.title || '情報',
      duration: options.duration || UI_CONSTANTS.NOTIFICATION_DURATION.INFO,
      ...options
    });
  }

  /**
   * 通知を表示
   * @param {Object} notificationData - 通知データ
   * @returns {Promise<string>} 通知ID
   */
  async showNotification(notificationData) {
    try {
      this.logger.debug('通知表示処理を開始します', { type: notificationData.type });

      // 通知データの妥当性検証
      const validationResult = this.validateNotification(notificationData);
      if (!validationResult.isValid) {
        throw new Error(`通知データが無効です: ${validationResult.errors.join(', ')}`);
      }

      const notification = {
        id: this.generateNotificationId(),
        type: notificationData.type,
        title: notificationData.title,
        message: notificationData.message,
        duration: notificationData.duration,
        persistent: notificationData.persistent || false,
        actions: notificationData.actions || [],
        createdAt: new Date().toISOString(),
        ...notificationData
      };

      // 通知をキューに追加
      this.notificationQueue.push(notification);

      // キュー処理を開始
      this.processNotificationQueue();

      this.logger.debug('通知表示処理が完了しました', { id: notification.id });
      return notification.id;
    } catch (error) {
      this.logger.error('通知表示処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 通知を非表示
   * @param {string} notificationId - 通知ID
   * @returns {Promise<boolean>} 成功フラグ
   */
  async hideNotification(notificationId) {
    try {
      this.logger.debug('通知非表示処理を開始します', { id: notificationId });

      if (!this.validator.isString(notificationId) || notificationId.trim().length === 0) {
        throw new Error('通知IDが無効です');
      }

      // アクティブな通知から削除
      if (this.activeNotifications.has(notificationId)) {
        const notification = this.activeNotifications.get(notificationId);

        // Chrome通知を削除
        if (notification.chromeNotificationId) {
          await this.clearChromeNotification(notification.chromeNotificationId);
        }

        // UI通知を削除
        if (notification.uiElement) {
          this.removeUINotification(notification.uiElement);
        }

        // タイマーをクリア
        if (notification.timer) {
          clearTimeout(notification.timer);
        }

        this.activeNotifications.delete(notificationId);
      }

      this.logger.debug('通知非表示処理が完了しました', { id: notificationId });
      return true;
    } catch (error) {
      this.logger.error('通知非表示処理中にエラーが発生しました', error);
      return false;
    }
  }

  /**
   * 全ての通知を非表示
   * @returns {Promise<boolean>} 成功フラグ
   */
  async hideAllNotifications() {
    try {
      this.logger.debug('全通知非表示処理を開始します');

      const notificationIds = Array.from(this.activeNotifications.keys());
      const hidePromises = notificationIds.map(id => this.hideNotification(id));

      await Promise.all(hidePromises);

      // キューもクリア
      this.notificationQueue.length = 0;

      this.logger.debug('全通知非表示処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('全通知非表示処理中にエラーが発生しました', error);
      return false;
    }
  }

  /**
   * プログレス通知を表示
   * @param {string} message - メッセージ
   * @param {number} progress - 進捗（0-100）
   * @param {Object} options - オプション
   * @returns {Promise<string>} 通知ID
   */
  async showProgress(message, progress = 0, options = {}) {
    return this.showNotification({
      type: 'progress',
      message,
      progress: Math.max(0, Math.min(100, progress)),
      title: options.title || '処理中',
      persistent: true,
      ...options
    });
  }

  /**
   * プログレス通知を更新
   * @param {string} notificationId - 通知ID
   * @param {number} progress - 進捗（0-100）
   * @param {string} [message] - メッセージ
   * @returns {Promise<boolean>} 成功フラグ
   */
  async updateProgress(notificationId, progress, message = null) {
    try {
      if (!this.activeNotifications.has(notificationId)) {
        return false;
      }

      const notification = this.activeNotifications.get(notificationId);
      notification.progress = Math.max(0, Math.min(100, progress));

      if (message) {
        notification.message = message;
      }

      // UI通知を更新
      if (notification.uiElement) {
        this.updateUINotification(notification.uiElement, notification);
      }

      return true;
    } catch (error) {
      this.logger.error('プログレス通知更新中にエラーが発生しました', error);
      return false;
    }
  }

  /**
   * 通知キューを処理
   * @private
   */
  async processNotificationQueue() {
    if (this.isProcessingQueue || this.notificationQueue.length === 0) {
      return;
    }

    this.isProcessingQueue = true;

    try {
      while (this.notificationQueue.length > 0) {
        const notification = this.notificationQueue.shift();
        await this.displayNotification(notification);

        // 短い間隔で連続表示を避ける
        await this.sleep(100);
      }
    } catch (error) {
      this.logger.error('通知キュー処理中にエラーが発生しました', error);
    } finally {
      this.isProcessingQueue = false;
    }
  }

  /**
   * 通知を実際に表示
   * @param {Object} notification - 通知データ
   * @private
   */
  async displayNotification(notification) {
    try {
      // Chrome通知を表示
      if (this.shouldUseChromeNotification(notification)) {
        await this.showChromeNotification(notification);
      }

      // UI通知を表示
      if (this.shouldUseUINotification(notification)) {
        await this.showUINotification(notification);
      }

      // アクティブ通知として登録
      this.activeNotifications.set(notification.id, notification);

      // 自動非表示タイマーを設定
      if (!notification.persistent && notification.duration > 0) {
        notification.timer = setTimeout(() => {
          this.hideNotification(notification.id);
        }, notification.duration);
      }
    } catch (error) {
      this.logger.error('通知表示中にエラーが発生しました', error);
    }
  }

  /**
   * Chrome通知を表示すべきかチェック
   * @param {Object} notification - 通知データ
   * @returns {boolean} 表示フラグ
   * @private
   */
  shouldUseChromeNotification(notification) {
    return notification.type === 'error' ||
      notification.type === 'warning' ||
      notification.persistent;
  }

  /**
   * UI通知を表示すべきかチェック
   * @param {Object} notification - 通知データ
   * @returns {boolean} 表示フラグ
   * @private
   */
  shouldUseUINotification(notification) {
    return true; // 基本的に全ての通知でUI表示
  }

  /**
   * Chrome通知を表示
   * @param {Object} notification - 通知データ
   * @private
   */
  async showChromeNotification(notification) {
    try {
      if (!chrome?.notifications) {
        return;
      }

      const chromeNotificationId = await chrome.notifications.create({
        type: 'basic',
        iconUrl: chrome.runtime.getURL('assets/icons/icon48.png'),
        title: notification.title,
        message: notification.message,
        contextMessage: this.getContextMessage(notification.type),
        isClickable: true
      });

      notification.chromeNotificationId = chromeNotificationId;
    } catch (error) {
      this.logger.error('Chrome通知表示中にエラーが発生しました', error);
    }
  }

  /**
   * UI通知を表示
   * @param {Object} notification - 通知データ
   * @private
   */
  async showUINotification(notification) {
    try {
      // UI通知要素を作成
      const uiElement = this.createUINotificationElement(notification);
      notification.uiElement = uiElement;

      // DOM に追加
      this.addUINotificationToDOM(uiElement);
    } catch (error) {
      this.logger.error('UI通知表示中にエラーが発生しました', error);
    }
  }

  /**
   * UI通知要素を作成
   * @param {Object} notification - 通知データ
   * @returns {HTMLElement} UI要素
   * @private
   */
  createUINotificationElement(notification) {
    const element = document.createElement('div');
    element.className = `ai-notification ai-notification-${notification.type}`;
    element.setAttribute('data-notification-id', notification.id);

    const iconClass = this.getNotificationIcon(notification.type);
    const progressHtml = notification.type === 'progress' ?
      `<div class="ai-notification-progress">
                <div class="ai-notification-progress-bar" style="width: ${notification.progress}%"></div>
            </div>` : '';

    element.innerHTML = `
            <div class="ai-notification-content">
                <div class="ai-notification-header">
                    <i class="ai-notification-icon ${iconClass}"></i>
                    <span class="ai-notification-title">${notification.title}</span>
                    <button class="ai-notification-close" type="button">&times;</button>
                </div>
                <div class="ai-notification-message">${notification.message}</div>
                ${progressHtml}
            </div>
        `;

    // 閉じるボタンのイベント
    const closeButton = element.querySelector('.ai-notification-close');
    closeButton.addEventListener('click', () => {
      this.hideNotification(notification.id);
    });

    return element;
  }

  /**
   * 通知データの妥当性を検証
   * @param {Object} notification - 通知データ
   * @returns {Object} 検証結果
   * @private
   */
  validateNotification(notification) {
    const errors = [];

    try {
      // 必須フィールドの確認
      if (!UI_CONSTANTS.NOTIFICATION_TYPES.includes(notification.type)) {
        errors.push('通知タイプが無効です');
      }

      if (!this.validator.isString(notification.message) || notification.message.trim().length === 0) {
        errors.push('メッセージが無効です');
      }

      if (notification.title && (!this.validator.isString(notification.title) || notification.title.trim().length === 0)) {
        errors.push('タイトルが無効です');
      }

      // 数値フィールドの確認
      if (notification.duration !== undefined && !this.validator.isNonNegativeInteger(notification.duration)) {
        errors.push('表示時間が無効です');
      }

      if (notification.progress !== undefined && !this.validator.isNumber(notification.progress)) {
        errors.push('進捗値が無効です');
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('通知妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['通知の妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * 通知IDを生成
   * @returns {string} 通知ID
   * @private
   */
  generateNotificationId() {
    return `notification_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * コンテキストメッセージを取得
   * @param {string} type - 通知タイプ
   * @returns {string} コンテキストメッセージ
   * @private
   */
  getContextMessage(type) {
    const messages = {
      success: 'AI業務支援ツール - 成功',
      error: 'AI業務支援ツール - エラー',
      warning: 'AI業務支援ツール - 警告',
      info: 'AI業務支援ツール - 情報'
    };
    return messages[type] || 'AI業務支援ツール';
  }

  /**
   * 通知アイコンを取得
   * @param {string} type - 通知タイプ
   * @returns {string} アイコンクラス
   * @private
   */
  getNotificationIcon(type) {
    const icons = {
      success: 'fas fa-check-circle',
      error: 'fas fa-exclamation-circle',
      warning: 'fas fa-exclamation-triangle',
      info: 'fas fa-info-circle',
      progress: 'fas fa-spinner fa-spin'
    };
    return icons[type] || 'fas fa-bell';
  }

  /**
   * Chrome通知をクリア
   * @param {string} chromeNotificationId - Chrome通知ID
   * @private
   */
  async clearChromeNotification(chromeNotificationId) {
    try {
      if (chrome?.notifications) {
        await chrome.notifications.clear(chromeNotificationId);
      }
    } catch (error) {
      this.logger.error('Chrome通知クリア中にエラーが発生しました', error);
    }
  }

  /**
   * UI通知を削除
   * @param {HTMLElement} element - UI要素
   * @private
   */
  removeUINotification(element) {
    try {
      if (element && element.parentNode) {
        element.remove();
      }
    } catch (error) {
      this.logger.error('UI通知削除中にエラーが発生しました', error);
    }
  }

  /**
   * UI通知を更新
   * @param {HTMLElement} element - UI要素
   * @param {Object} notification - 通知データ
   * @private
   */
  updateUINotification(element, notification) {
    try {
      const messageElement = element.querySelector('.ai-notification-message');
      if (messageElement) {
        messageElement.textContent = notification.message;
      }

      const progressBar = element.querySelector('.ai-notification-progress-bar');
      if (progressBar && notification.progress !== undefined) {
        progressBar.style.width = `${notification.progress}%`;
      }
    } catch (error) {
      this.logger.error('UI通知更新中にエラーが発生しました', error);
    }
  }

  /**
   * UI通知をDOMに追加
   * @param {HTMLElement} element - UI要素
   * @private
   */
  addUINotificationToDOM(element) {
    try {
      // 通知コンテナを取得または作成
      let container = document.getElementById('ai-notifications-container');
      if (!container) {
        container = document.createElement('div');
        container.id = 'ai-notifications-container';
        container.className = 'ai-notifications-container';
        document.body.appendChild(container);
      }

      container.appendChild(element);
    } catch (error) {
      this.logger.error('UI通知DOM追加中にエラーが発生しました', error);
    }
  }

  /**
   * スリープ
   * @param {number} ms - ミリ秒
   * @returns {Promise<void>}
   * @private
   */
  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

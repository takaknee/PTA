/*
 * AI業務支援ツール Edge拡張機能 - サービス層エクスポート
 * Copyright (c) 2024 AI Business Support Team
 */

// コアサービス
export { SettingsService } from './settingsService.js';
export { HistoryService } from './historyService.js';
export { StatisticsService } from './statisticsService.js';
export { ApiService } from './apiService.js';
export { PromptService } from './promptService.js';
export { MessagingService } from './messagingService.js';
export { NotificationService } from './notificationService.js';

/**
 * サービス管理クラス
 * シングルトンパターンでサービスインスタンスを管理
 */
export class ServiceManager {
  constructor() {
    this._services = new Map();
    this._initialized = false;
  }

  /**
   * サービスマネージャーを初期化
   */
  async initialize() {
    if (this._initialized) {
      return;
    }

    try {
      // 基本サービスを初期化
      this._services.set('settings', new SettingsService());
      this._services.set('history', new HistoryService());
      this._services.set('statistics', new StatisticsService());
      this._services.set('api', new ApiService());
      this._services.set('prompt', new PromptService());
      this._services.set('messaging', new MessagingService());
      this._services.set('notification', new NotificationService());

      this._initialized = true;
      console.log('ServiceManager が正常に初期化されました');
    } catch (error) {
      console.error('ServiceManager の初期化中にエラーが発生しました:', error);
      throw error;
    }
  }

  /**
   * サービスを取得
   * @param {string} serviceName - サービス名
   * @returns {Object} サービスインスタンス
   */
  getService(serviceName) {
    if (!this._initialized) {
      throw new Error('ServiceManager が初期化されていません');
    }

    const service = this._services.get(serviceName);
    if (!service) {
      throw new Error(`サービスが見つかりません: ${serviceName}`);
    }

    return service;
  }

  /**
   * 設定サービスを取得
   * @returns {SettingsService} 設定サービス
   */
  get settings() {
    return this.getService('settings');
  }

  /**
   * 履歴サービスを取得
   * @returns {HistoryService} 履歴サービス
   */
  get history() {
    return this.getService('history');
  }

  /**
   * 統計サービスを取得
   * @returns {StatisticsService} 統計サービス
   */
  get statistics() {
    return this.getService('statistics');
  }

  /**
   * APIサービスを取得
   * @returns {ApiService} APIサービス
   */
  get api() {
    return this.getService('api');
  }

  /**
   * プロンプトサービスを取得
   * @returns {PromptService} プロンプトサービス
   */
  get prompt() {
    return this.getService('prompt');
  }

  /**
   * メッセージングサービスを取得
   * @returns {MessagingService} メッセージングサービス
   */
  get messaging() {
    return this.getService('messaging');
  }

  /**
   * 通知サービスを取得
   * @returns {NotificationService} 通知サービス
   */
  get notification() {
    return this.getService('notification');
  }

  /**
   * 全サービスを停止
   */
  async shutdown() {
    try {
      // メッセージングサービスのハンドラーをクリア
      if (this._services.has('messaging')) {
        this._services.get('messaging').clearAllHandlers();
      }

      // 通知をクリア
      if (this._services.has('notification')) {
        await this._services.get('notification').hideAllNotifications();
      }

      this._services.clear();
      this._initialized = false;
      console.log('ServiceManager が正常に停止されました');
    } catch (error) {
      console.error('ServiceManager の停止中にエラーが発生しました:', error);
    }
  }

  /**
   * サービスの状態を取得
   * @returns {Object} サービス状態
   */
  getStatus() {
    return {
      initialized: this._initialized,
      serviceCount: this._services.size,
      services: Array.from(this._services.keys())
    };
  }
}

// シングルトンインスタンス
let serviceManagerInstance = null;

/**
 * ServiceManager のシングルトンインスタンスを取得
 * @returns {ServiceManager} サービスマネージャー
 */
export function getServiceManager() {
  if (!serviceManagerInstance) {
    serviceManagerInstance = new ServiceManager();
  }
  return serviceManagerInstance;
}

/*
 * AI業務支援ツール Edge拡張機能 - サービス層エクスポート
 * Copyright (c) 2024 AI Business Support Team
 */

// Core services
export { ApiService } from './apiService.js';
export { SettingsService } from './settingsService.js';
export { HistoryService } from './historyService.js';
export { StatisticsService } from './statisticsService.js';
export { MessagingService, messagingService } from './messagingService.js';
export { PromptService } from './promptService.js';

// Service factory for creating service instances
export class ServiceFactory {
  constructor() {
    this.services = new Map();
    this.initialized = false;
  }

  /**
   * サービスファクトリを初期化
   */
  async initialize() {
    if (this.initialized) {
      return;
    }

    try {
      // Core services の初期化
      this.services.set('api', new ApiService());
      this.services.set('settings', new SettingsService());
      this.services.set('history', new HistoryService());
      this.services.set('statistics', new StatisticsService());
      this.services.set('messaging', messagingService);
      this.services.set('prompt', new PromptService());

      // Messaging service の初期化
      messagingService.initialize();

      this.initialized = true;
      console.log('ServiceFactory initialized successfully');

    } catch (error) {
      console.error('ServiceFactory initialization failed:', error);
      throw error;
    }
  }

  /**
   * サービスインスタンスを取得
   * @param {string} serviceName サービス名
   * @returns {Object} サービスインスタンス
   */
  getService(serviceName) {
    if (!this.initialized) {
      throw new Error('ServiceFactory is not initialized. Call initialize() first.');
    }

    const service = this.services.get(serviceName);
    if (!service) {
      throw new Error(`Service '${serviceName}' not found`);
    }

    return service;
  }

  /**
   * すべてのサービスインスタンスを取得
   * @returns {Object} サービスインスタンス群
   */
  getAllServices() {
    if (!this.initialized) {
      throw new Error('ServiceFactory is not initialized. Call initialize() first.');
    }

    return {
      api: this.getService('api'),
      settings: this.getService('settings'),
      history: this.getService('history'),
      statistics: this.getService('statistics'),
      messaging: this.getService('messaging'),
      prompt: this.getService('prompt')
    };
  }

  /**
   * サービスファクトリをクリーンアップ
   */
  cleanup() {
    if (this.services.has('messaging')) {
      this.services.get('messaging').cleanup();
    }

    this.services.clear();
    this.initialized = false;
    console.log('ServiceFactory cleaned up');
  }
}

// Global service factory instance
export const serviceFactory = new ServiceFactory();

// Convenience function to get services
export async function getServices() {
  if (!serviceFactory.initialized) {
    await serviceFactory.initialize();
  }
  return serviceFactory.getAllServices();
}

// Individual service getters for convenience
export async function getApiService() {
  if (!serviceFactory.initialized) {
    await serviceFactory.initialize();
  }
  return serviceFactory.getService('api');
}

export async function getSettingsService() {
  if (!serviceFactory.initialized) {
    await serviceFactory.initialize();
  }
  return serviceFactory.getService('settings');
}

export async function getHistoryService() {
  if (!serviceFactory.initialized) {
    await serviceFactory.initialize();
  }
  return serviceFactory.getService('history');
}

export async function getStatisticsService() {
  if (!serviceFactory.initialized) {
    await serviceFactory.initialize();
  }
  return serviceFactory.getService('statistics');
}

export async function getMessagingService() {
  if (!serviceFactory.initialized) {
    await serviceFactory.initialize();
  }
  return serviceFactory.getService('messaging');
}

export async function getPromptService() {
  if (!serviceFactory.initialized) {
    await serviceFactory.initialize();
  }
  return serviceFactory.getService('prompt');
}

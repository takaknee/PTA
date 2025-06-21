/*
 * AI業務支援ツール Edge拡張機能 - 設定エンティティ
 * Copyright (c) 2024 AI Business Support Team
 */

import { DEFAULT_SETTINGS } from '../../shared/constants/index.js';
import { Validator } from '../../shared/utils/index.js';

/**
 * API設定エンティティ
 */
export class ApiSettings {
  constructor(data = {}) {
    this.provider = data.provider || DEFAULT_SETTINGS.api.provider;
    this.model = data.model || DEFAULT_SETTINGS.api.model;
    this.apiKey = data.apiKey || DEFAULT_SETTINGS.api.apiKey;
    this.azureEndpoint = data.azureEndpoint || DEFAULT_SETTINGS.api.azureEndpoint;
    this.azureDeploymentName = data.azureDeploymentName || DEFAULT_SETTINGS.api.azureDeploymentName;
    this.maxTokens = data.maxTokens || DEFAULT_SETTINGS.api.maxTokens;
    this.temperature = data.temperature || DEFAULT_SETTINGS.api.temperature;
    this.requestTimeout = data.requestTimeout || DEFAULT_SETTINGS.api.requestTimeout;
  }

  /**
   * 設定の検証
   * @returns {Object} 検証結果
   */
  validate() {
    const validations = [
      () => Validator.validateApiKey(this.apiKey)
    ];

    if (this.provider === 'azure') {
      validations.push(
        () => Validator.validateAzureEndpoint(this.azureEndpoint),
        () => Validator.validateDeploymentName(this.azureDeploymentName)
      );
    }

    return Validator.collectMultipleValidationResults(validations);
  }

  /**
   * 設定の完了状態確認
   * @returns {boolean} 完了フラグ
   */
  isComplete() {
    if (!this.apiKey) return false;

    if (this.provider === 'azure') {
      return !!(this.azureEndpoint && this.azureDeploymentName);
    }

    return true;
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      provider: this.provider,
      model: this.model,
      apiKey: this.apiKey,
      azureEndpoint: this.azureEndpoint,
      azureDeploymentName: this.azureDeploymentName,
      maxTokens: this.maxTokens,
      temperature: this.temperature,
      requestTimeout: this.requestTimeout
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {ApiSettings} 設定インスタンス
   */
  static fromJSON(json) {
    return new ApiSettings(json);
  }

  /**
   * 設定のクローン
   * @returns {ApiSettings} クローン
   */
  clone() {
    return new ApiSettings(this.toJSON());
  }
}

/**
 * UI設定エンティティ
 */
export class UISettings {
  constructor(data = {}) {
    this.theme = data.theme || DEFAULT_SETTINGS.ui.theme;
    this.showNotifications = data.showNotifications !== undefined
      ? data.showNotifications
      : DEFAULT_SETTINGS.ui.showNotifications;
    this.notificationDuration = data.notificationDuration || DEFAULT_SETTINGS.ui.notificationDuration;
    this.buttonPosition = data.buttonPosition || { ...DEFAULT_SETTINGS.ui.buttonPosition };
    this.expandByDefault = data.expandByDefault !== undefined
      ? data.expandByDefault
      : DEFAULT_SETTINGS.ui.expandByDefault;
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      theme: this.theme,
      showNotifications: this.showNotifications,
      notificationDuration: this.notificationDuration,
      buttonPosition: { ...this.buttonPosition },
      expandByDefault: this.expandByDefault
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {UISettings} 設定インスタンス
   */
  static fromJSON(json) {
    return new UISettings(json);
  }
}

/**
 * 機能設定エンティティ
 */
export class FeatureSettings {
  constructor(data = {}) {
    this.autoDetect = data.autoDetect !== undefined
      ? data.autoDetect
      : DEFAULT_SETTINGS.features.autoDetect;
    this.saveHistory = data.saveHistory !== undefined
      ? data.saveHistory
      : DEFAULT_SETTINGS.features.saveHistory;
    this.enableStatistics = data.enableStatistics !== undefined
      ? data.enableStatistics
      : DEFAULT_SETTINGS.features.enableStatistics;
    this.enableCache = data.enableCache !== undefined
      ? data.enableCache
      : DEFAULT_SETTINGS.features.enableCache;
    this.enableDiagnostics = data.enableDiagnostics !== undefined
      ? data.enableDiagnostics
      : DEFAULT_SETTINGS.features.enableDiagnostics;
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      autoDetect: this.autoDetect,
      saveHistory: this.saveHistory,
      enableStatistics: this.enableStatistics,
      enableCache: this.enableCache,
      enableDiagnostics: this.enableDiagnostics
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {FeatureSettings} 設定インスタンス
   */
  static fromJSON(json) {
    return new FeatureSettings(json);
  }
}

/**
 * 高度な設定エンティティ
 */
export class AdvancedSettings {
  constructor(data = {}) {
    this.debugMode = data.debugMode !== undefined
      ? data.debugMode
      : DEFAULT_SETTINGS.advanced.debugMode;
    this.logLevel = data.logLevel || DEFAULT_SETTINGS.advanced.logLevel;
    this.enableTelemetry = data.enableTelemetry !== undefined
      ? data.enableTelemetry
      : DEFAULT_SETTINGS.advanced.enableTelemetry;
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      debugMode: this.debugMode,
      logLevel: this.logLevel,
      enableTelemetry: this.enableTelemetry
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {AdvancedSettings} 設定インスタンス
   */
  static fromJSON(json) {
    return new AdvancedSettings(json);
  }
}

/**
 * 拡張機能設定エンティティ（統合）
 */
export class ExtensionSettings {
  constructor(data = {}) {
    this.api = data.api ? ApiSettings.fromJSON(data.api) : new ApiSettings();
    this.ui = data.ui ? UISettings.fromJSON(data.ui) : new UISettings();
    this.features = data.features ? FeatureSettings.fromJSON(data.features) : new FeatureSettings();
    this.advanced = data.advanced ? AdvancedSettings.fromJSON(data.advanced) : new AdvancedSettings();
    this.version = data.version || '1.0.0';
    this.lastUpdated = data.lastUpdated || new Date().toISOString();
  }

  /**
   * 設定の検証
   * @returns {Object} 検証結果
   */
  validate() {
    return this.api.validate();
  }

  /**
   * 設定の完了状態確認
   * @returns {boolean} 完了フラグ
   */
  isComplete() {
    return this.api.isComplete();
  }

  /**
   * 設定の更新
   * @param {Object} updates - 更新データ
   */
  update(updates) {
    if (updates.api) {
      Object.assign(this.api, updates.api);
    }
    if (updates.ui) {
      Object.assign(this.ui, updates.ui);
    }
    if (updates.features) {
      Object.assign(this.features, updates.features);
    }
    if (updates.advanced) {
      Object.assign(this.advanced, updates.advanced);
    }

    this.lastUpdated = new Date().toISOString();
  }

  /**
   * 設定のリセット
   * @param {string[]} sections - リセット対象セクション
   */
  reset(sections = ['api', 'ui', 'features', 'advanced']) {
    if (sections.includes('api')) {
      this.api = new ApiSettings();
    }
    if (sections.includes('ui')) {
      this.ui = new UISettings();
    }
    if (sections.includes('features')) {
      this.features = new FeatureSettings();
    }
    if (sections.includes('advanced')) {
      this.advanced = new AdvancedSettings();
    }

    this.lastUpdated = new Date().toISOString();
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      api: this.api.toJSON(),
      ui: this.ui.toJSON(),
      features: this.features.toJSON(),
      advanced: this.advanced.toJSON(),
      version: this.version,
      lastUpdated: this.lastUpdated
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {ExtensionSettings} 設定インスタンス
   */
  static fromJSON(json) {
    return new ExtensionSettings(json);
  }

  /**
   * デフォルト設定の取得
   * @returns {ExtensionSettings} デフォルト設定
   */
  static getDefault() {
    return new ExtensionSettings();
  }
}

/*
 * AI業務支援ツール Edge拡張機能 - 設定リポジトリ
 * Copyright (c) 2024 AI Business Support Team
 */

import { STORAGE_KEYS } from '../../shared/constants/index.js';
import { ExtensionSettings } from '../../core/entities/index.js';
import { BaseStorageRepository } from './storageRepository.js';
import { createLogger } from '../../shared/utils/index.js';

const logger = createLogger('SettingsRepository');

/**
 * 設定リポジトリ
 */
export class SettingsRepository {
  constructor(storageRepository = null) {
    this.storage = storageRepository || new BaseStorageRepository();
    this.cacheKey = STORAGE_KEYS.SETTINGS;
    this.cache = null;
    this.cacheTime = 0;
    this.cacheDuration = 5 * 60 * 1000; // 5分
  }

  /**
   * 設定の取得
   * @param {boolean} useCache - キャッシュ使用フラグ
   * @returns {Promise<ExtensionSettings>} 設定オブジェクト
   */
  async get(useCache = true) {
    try {
      // キャッシュチェック
      if (useCache && this.isValid()) {
        logger.debug('キャッシュから設定を返却');
        return this.cache;
      }

      logger.debug('ストレージから設定を取得');
      const data = await this.storage.getJSON(this.cacheKey, {});

      const settings = ExtensionSettings.fromJSON(data);

      // キャッシュを更新
      this.cache = settings;
      this.cacheTime = Date.now();

      logger.debug('設定取得完了', { hasApiKey: !!settings.api.apiKey });
      return settings;

    } catch (error) {
      logger.error('設定取得エラー', { error: error.message });

      // フォールバック: デフォルト設定を返す
      const defaultSettings = ExtensionSettings.getDefault();
      this.cache = defaultSettings;
      this.cacheTime = Date.now();

      return defaultSettings;
    }
  }

  /**
   * 設定の保存
   * @param {ExtensionSettings} settings - 設定オブジェクト
   * @returns {Promise<void>}
   */
  async save(settings) {
    try {
      logger.debug('設定保存開始');

      if (!(settings instanceof ExtensionSettings)) {
        throw new Error('無効な設定オブジェクトです');
      }

      // 保存前にタイムスタンプを更新
      settings.lastUpdated = new Date().toISOString();

      await this.storage.setJSON(this.cacheKey, settings.toJSON());

      // キャッシュを更新
      this.cache = settings;
      this.cacheTime = Date.now();

      logger.debug('設定保存完了');

    } catch (error) {
      logger.error('設定保存エラー', { error: error.message });
      throw error;
    }
  }

  /**
   * 設定の部分更新
   * @param {Object} updates - 更新データ
   * @returns {Promise<ExtensionSettings>} 更新後の設定
   */
  async update(updates) {
    try {
      logger.debug('設定部分更新開始', { updates: Object.keys(updates) });

      const currentSettings = await this.get();
      currentSettings.update(updates);

      await this.save(currentSettings);

      logger.debug('設定部分更新完了');
      return currentSettings;

    } catch (error) {
      logger.error('設定部分更新エラー', { error: error.message });
      throw error;
    }
  }

  /**
   * 設定のリセット
   * @param {string[]} sections - リセット対象セクション
   * @returns {Promise<ExtensionSettings>} リセット後の設定
   */
  async reset(sections = ['api', 'ui', 'features', 'advanced']) {
    try {
      logger.debug('設定リセット開始', { sections });

      const currentSettings = await this.get();
      currentSettings.reset(sections);

      await this.save(currentSettings);

      logger.debug('設定リセット完了');
      return currentSettings;

    } catch (error) {
      logger.error('設定リセットエラー', { error: error.message });
      throw error;
    }
  }

  /**
   * 設定の削除
   * @returns {Promise<void>}
   */
  async delete() {
    try {
      logger.debug('設定削除開始');

      await this.storage.remove(this.cacheKey);

      // キャッシュをクリア
      this.cache = null;
      this.cacheTime = 0;

      logger.debug('設定削除完了');

    } catch (error) {
      logger.error('設定削除エラー', { error: error.message });
      throw error;
    }
  }

  /**
   * 設定の存在確認
   * @returns {Promise<boolean>} 存在フラグ
   */
  async exists() {
    try {
      return await this.storage.exists(this.cacheKey);
    } catch (error) {
      logger.warn('設定存在確認エラー', { error: error.message });
      return false;
    }
  }

  /**
   * 設定の検証
   * @returns {Promise<Object>} 検証結果
   */
  async validate() {
    try {
      const settings = await this.get();
      return settings.validate();
    } catch (error) {
      logger.error('設定検証エラー', { error: error.message });
      return {
        valid: false,
        errors: [{
          code: 'SYSTEM_ERROR',
          message: error.message,
          field: null
        }]
      };
    }
  }

  /**
   * API設定の取得
   * @returns {Promise<ApiSettings>} API設定
   */
  async getApiSettings() {
    const settings = await this.get();
    return settings.api;
  }

  /**
   * UI設定の取得
   * @returns {Promise<UISettings>} UI設定
   */
  async getUISettings() {
    const settings = await this.get();
    return settings.ui;
  }

  /**
   * 機能設定の取得
   * @returns {Promise<FeatureSettings>} 機能設定
   */
  async getFeatureSettings() {
    const settings = await this.get();
    return settings.features;
  }

  /**
   * 設定のエクスポート
   * @returns {Promise<Object>} エクスポートデータ
   */
  async export() {
    try {
      const settings = await this.get();

      // APIキーを除外したエクスポート用データ
      const exportData = settings.toJSON();
      exportData.api.apiKey = '[HIDDEN]';

      return {
        version: '1.0.0',
        exportedAt: new Date().toISOString(),
        settings: exportData
      };

    } catch (error) {
      logger.error('設定エクスポートエラー', { error: error.message });
      throw error;
    }
  }

  /**
   * 設定のインポート
   * @param {Object} importData - インポートデータ
   * @param {boolean} mergeWithCurrent - 現在の設定とマージするか
   * @returns {Promise<ExtensionSettings>} インポート後の設定
   */
  async import(importData, mergeWithCurrent = true) {
    try {
      logger.debug('設定インポート開始');

      if (!importData.settings) {
        throw new Error('無効なインポートデータです');
      }

      let settings;
      if (mergeWithCurrent) {
        settings = await this.get();
        settings.update(importData.settings);
      } else {
        settings = ExtensionSettings.fromJSON(importData.settings);
      }

      await this.save(settings);

      logger.debug('設定インポート完了');
      return settings;

    } catch (error) {
      logger.error('設定インポートエラー', { error: error.message });
      throw error;
    }
  }

  /**
   * キャッシュの有効性確認
   * @returns {boolean} 有効フラグ
   */
  isValid() {
    return this.cache &&
      this.cacheTime > 0 &&
      (Date.now() - this.cacheTime) < this.cacheDuration;
  }

  /**
   * キャッシュのクリア
   */
  clearCache() {
    this.cache = null;
    this.cacheTime = 0;
    logger.debug('設定キャッシュをクリアしました');
  }

  /**
   * キャッシュ期間の設定
   * @param {number} duration - キャッシュ期間（ミリ秒）
   */
  setCacheDuration(duration) {
    this.cacheDuration = duration;
    logger.debug('キャッシュ期間を更新', { duration });
  }
}

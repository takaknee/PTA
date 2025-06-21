/*
 * AI業務支援ツール Edge拡張機能 - 設定管理サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { ApiSettings } from '../lib/models.js';
import { DEFAULT_SETTINGS, STORAGE_KEYS } from '../lib/constants.js';
import { Logger, deepMerge, safeJsonParse, safeJsonStringify } from '../lib/utils.js';

/**
 * 設定管理サービス
 */
export class SettingsService {
  constructor() {
    this.cache = null;
    this.listeners = new Set();
  }

  /**
   * 設定を取得
   * @param {boolean} useCache キャッシュを使用するか
   * @returns {Promise<ApiSettings>} 設定
   */
  async getSettings(useCache = true) {
    try {
      // キャッシュがある場合はキャッシュを返す
      if (useCache && this.cache) {
        Logger.debug('設定をキャッシュから取得');
        return this.cache;
      }

      Logger.debug('設定をストレージから取得');

      // Chrome storage から設定を取得
      const result = await chrome.storage.sync.get(STORAGE_KEYS.SETTINGS);
      const storedSettings = result[STORAGE_KEYS.SETTINGS] || {};

      // デフォルト設定とマージ
      const mergedSettings = deepMerge(DEFAULT_SETTINGS, storedSettings);

      // ApiSettingsオブジェクトを作成
      const settings = new ApiSettings(mergedSettings);

      // キャッシュに保存
      this.cache = settings;

      Logger.debug('設定取得完了', { provider: settings.provider, model: settings.model });
      return settings;

    } catch (error) {
      Logger.error('設定取得エラー', error);

      // エラー時はデフォルト設定を返す
      const defaultSettings = new ApiSettings(DEFAULT_SETTINGS);
      this.cache = defaultSettings;
      return defaultSettings;
    }
  }

  /**
   * 設定を保存
   * @param {ApiSettings|Object} newSettings 新しい設定
   * @returns {Promise<boolean>} 保存成功可否
   */
  async saveSettings(newSettings) {
    try {
      Logger.info('設定保存開始');

      // ApiSettingsオブジェクトでない場合は変換
      const settings = newSettings instanceof ApiSettings
        ? newSettings
        : new ApiSettings(newSettings);

      // 設定の妥当性を検証
      const validation = settings.validate();
      if (!validation.isValid) {
        throw new Error(`設定検証エラー: ${validation.getErrorMessage()}`);
      }

      // Chrome storage に保存
      await chrome.storage.sync.set({
        [STORAGE_KEYS.SETTINGS]: settings.toJSON()
      });

      // キャッシュを更新
      this.cache = settings;

      // リスナーに変更を通知
      this.notifyListeners(settings);

      Logger.info('設定保存完了', { provider: settings.provider, model: settings.model });
      return true;

    } catch (error) {
      Logger.error('設定保存エラー', error);
      return false;
    }
  }

  /**
   * 設定を部分的に更新
   * @param {Object} partialSettings 部分的な設定
   * @returns {Promise<boolean>} 更新成功可否
   */
  async updateSettings(partialSettings) {
    try {
      // 現在の設定を取得
      const currentSettings = await this.getSettings();

      // 部分的な設定をマージ
      const updatedData = deepMerge(currentSettings.toJSON(), partialSettings);

      // 保存
      return await this.saveSettings(updatedData);

    } catch (error) {
      Logger.error('設定更新エラー', error);
      return false;
    }
  }

  /**
   * 設定をリセット
   * @returns {Promise<boolean>} リセット成功可否
   */
  async resetSettings() {
    try {
      Logger.info('設定リセット開始');

      // Chrome storage をクリア
      await chrome.storage.sync.remove(STORAGE_KEYS.SETTINGS);

      // キャッシュをクリア
      this.cache = null;

      // デフォルト設定を取得（キャッシュを使わない）
      const defaultSettings = await this.getSettings(false);

      // リスナーに変更を通知
      this.notifyListeners(defaultSettings);

      Logger.info('設定リセット完了');
      return true;

    } catch (error) {
      Logger.error('設定リセットエラー', error);
      return false;
    }
  }

  /**
   * 設定をエクスポート
   * @returns {Promise<string>} エクスポートされたJSON文字列
   */
  async exportSettings() {
    try {
      const settings = await this.getSettings();
      const exportData = {
        version: '1.0',
        timestamp: new Date().toISOString(),
        settings: settings.toJSON()
      };

      return safeJsonStringify(exportData, '{}');

    } catch (error) {
      Logger.error('設定エクスポートエラー', error);
      return '{}';
    }
  }

  /**
   * 設定をインポート
   * @param {string} jsonString インポートするJSON文字列
   * @returns {Promise<boolean>} インポート成功可否
   */
  async importSettings(jsonString) {
    try {
      Logger.info('設定インポート開始');

      const importData = safeJsonParse(jsonString);
      if (!importData || !importData.settings) {
        throw new Error('無効なインポートデータ');
      }

      // バージョンチェック（将来の拡張用）
      if (importData.version && importData.version !== '1.0') {
        Logger.warn('異なるバージョンの設定ファイル', importData.version);
      }

      // 設定を保存
      const success = await this.saveSettings(importData.settings);

      if (success) {
        Logger.info('設定インポート完了');
      } else {
        throw new Error('設定保存に失敗');
      }

      return success;

    } catch (error) {
      Logger.error('設定インポートエラー', error);
      return false;
    }
  }

  /**
   * 設定変更リスナーを追加
   * @param {Function} listener リスナー関数
   */
  addChangeListener(listener) {
    if (typeof listener === 'function') {
      this.listeners.add(listener);
      Logger.debug('設定変更リスナー追加', { count: this.listeners.size });
    }
  }

  /**
   * 設定変更リスナーを削除
   * @param {Function} listener リスナー関数
   */
  removeChangeListener(listener) {
    this.listeners.delete(listener);
    Logger.debug('設定変更リスナー削除', { count: this.listeners.size });
  }

  /**
   * リスナーに変更を通知
   * @param {ApiSettings} settings 新しい設定
   */
  notifyListeners(settings) {
    this.listeners.forEach(listener => {
      try {
        listener(settings);
      } catch (error) {
        Logger.error('リスナー通知エラー', error);
      }
    });
  }

  /**
   * キャッシュをクリア
   */
  clearCache() {
    this.cache = null;
    Logger.debug('設定キャッシュクリア');
  }

  /**
   * 設定の妥当性をチェック
   * @returns {Promise<Object>} 検証結果
   */
  async validateCurrentSettings() {
    try {
      const settings = await this.getSettings();
      const validation = settings.validate();

      return {
        isValid: validation.isValid,
        errors: validation.errors,
        settings: settings.toJSON()
      };

    } catch (error) {
      Logger.error('設定検証エラー', error);
      return {
        isValid: false,
        errors: [error.message],
        settings: null
      };
    }
  }

  /**
   * APIキーの部分的表示用文字列を生成
   * @param {string} apiKey APIキー
   * @returns {string} 部分的に隠されたAPIキー
   */
  maskApiKey(apiKey) {
    if (!apiKey || apiKey.length < 8) {
      return '*'.repeat(apiKey ? apiKey.length : 0);
    }

    const prefix = apiKey.substring(0, 4);
    const suffix = apiKey.substring(apiKey.length - 4);
    const middle = '*'.repeat(apiKey.length - 8);

    return `${prefix}${middle}${suffix}`;
  }

  /**
   * Azure エンドポイントの妥当性チェック
   * @param {string} endpoint エンドポイント
   * @returns {Object} 検証結果
   */
  validateAzureEndpoint(endpoint) {
    const result = {
      isValid: false,
      error: null,
      resourceName: null
    };

    try {
      if (!endpoint) {
        result.error = 'エンドポイントが空です';
        return result;
      }

      const url = new URL(endpoint);

      // Azure OpenAI のエンドポイント形式チェック
      if (!url.hostname.includes('.openai.azure.com')) {
        result.error = 'Azure OpenAI のエンドポイント形式ではありません';
        return result;
      }

      // HTTPSチェック
      if (url.protocol !== 'https:') {
        result.error = 'HTTPSプロトコルを使用してください';
        return result;
      }

      // リソース名を抽出
      result.resourceName = url.hostname.split('.')[0];
      result.isValid = true;

    } catch (error) {
      result.error = '無効なURL形式です';
    }

    return result;
  }

  /**
   * モデル互換性チェック
   * @param {string} provider プロバイダー
   * @param {string} model モデル名
   * @returns {boolean} 互換性があるかどうか
   */
  isModelCompatible(provider, model) {
    const compatibilityMap = {
      openai: ['gpt-4', 'gpt-4o', 'gpt-4o-mini', 'gpt-3.5-turbo'],
      azure: ['gpt-4', 'gpt-4o', 'gpt-4o-mini', 'gpt-35-turbo']
    };

    return compatibilityMap[provider]?.includes(model) || false;
  }
}

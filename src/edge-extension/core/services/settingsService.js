/*
 * AI業務支援ツール Edge拡張機能 - 設定サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { SettingsRepository } from '../repositories/settingsRepository.js';
import { Settings } from '../entities/settings.js';
import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';
import { APP_CONSTANTS } from '../../shared/constants/app.js';

/**
 * 設定管理サービス
 */
export class SettingsService {
  /**
   * コンストラクタ
   */
  constructor() {
    this.repository = new SettingsRepository();
    this.logger = new Logger('SettingsService');
    this.validator = new Validator();
  }

  /**
   * 設定を取得
   * @returns {Promise<Settings>} 設定オブジェクト
   */
  async getSettings() {
    try {
      this.logger.info('設定取得処理を開始します');

      const settingsData = await this.repository.getSettings();
      const settings = Settings.fromJSON(settingsData);

      this.logger.info('設定取得処理が完了しました');
      return settings;
    } catch (error) {
      this.logger.error('設定取得処理中にエラーが発生しました', error);

      // デフォルト設定を返す
      this.logger.info('デフォルト設定を返します');
      return Settings.createDefault();
    }
  }

  /**
   * 設定を保存
   * @param {Settings} settings - 設定オブジェクト
   * @returns {Promise<boolean>} 保存成功フラグ
   */
  async saveSettings(settings) {
    try {
      this.logger.info('設定保存処理を開始します');

      // 設定の妥当性検証
      const validationResult = this.validateSettings(settings);
      if (!validationResult.isValid) {
        throw new Error(`設定の妥当性検証に失敗しました: ${validationResult.errors.join(', ')}`);
      }

      // 設定を保存
      await this.repository.saveSettings(settings.toJSON());

      this.logger.info('設定保存処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('設定保存処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * API設定を更新
   * @param {Object} apiSettings - API設定
   * @returns {Promise<boolean>} 更新成功フラグ
   */
  async updateApiSettings(apiSettings) {
    try {
      this.logger.info('API設定更新処理を開始します');

      const currentSettings = await this.getSettings();
      const updatedSettings = currentSettings.updateApiSettings(apiSettings);

      await this.saveSettings(updatedSettings);

      this.logger.info('API設定更新処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('API設定更新処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * UI設定を更新
   * @param {Object} uiSettings - UI設定
   * @returns {Promise<boolean>} 更新成功フラグ
   */
  async updateUiSettings(uiSettings) {
    try {
      this.logger.info('UI設定更新処理を開始します');

      const currentSettings = await this.getSettings();
      const updatedSettings = currentSettings.updateUiSettings(uiSettings);

      await this.saveSettings(updatedSettings);

      this.logger.info('UI設定更新処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('UI設定更新処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 設定を初期化
   * @returns {Promise<boolean>} 初期化成功フラグ
   */
  async resetSettings() {
    try {
      this.logger.info('設定初期化処理を開始します');

      const defaultSettings = Settings.createDefault();
      await this.saveSettings(defaultSettings);

      this.logger.info('設定初期化処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('設定初期化処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 設定の妥当性を検証
   * @param {Settings} settings - 設定オブジェクト
   * @returns {Object} 検証結果
   */
  validateSettings(settings) {
    const errors = [];

    try {
      // API設定の検証
      if (!settings.api) {
        errors.push('API設定が設定されていません');
      } else {
        // APIキーの検証
        if (!this.validator.isString(settings.api.apiKey) ||
          settings.api.apiKey.length < APP_CONSTANTS.MIN_API_KEY_LENGTH) {
          errors.push('APIキーが無効です');
        }

        // プロバイダーの検証
        if (!APP_CONSTANTS.SUPPORTED_PROVIDERS.includes(settings.api.provider)) {
          errors.push('サポートされていないAPIプロバイダーです');
        }

        // モデル名の検証
        if (!this.validator.isString(settings.api.model) ||
          settings.api.model.trim().length === 0) {
          errors.push('モデル名が無効です');
        }

        // Azure固有設定の検証
        if (settings.api.provider === 'azure') {
          if (!this.validator.isUrl(settings.api.azureEndpoint)) {
            errors.push('AzureエンドポイントURLが無効です');
          }
          if (!this.validator.isString(settings.api.azureDeploymentName) ||
            settings.api.azureDeploymentName.trim().length === 0) {
            errors.push('Azureデプロイメント名が無効です');
          }
        }

        // トークン数の検証
        if (!this.validator.isPositiveInteger(settings.api.maxTokens) ||
          settings.api.maxTokens > APP_CONSTANTS.MAX_TOKENS_LIMIT) {
          errors.push('最大トークン数が無効です');
        }

        // 温度パラメータの検証
        if (!this.validator.isNumber(settings.api.temperature) ||
          settings.api.temperature < 0 || settings.api.temperature > 2) {
          errors.push('温度パラメータが無効です（0-2の範囲で設定してください）');
        }
      }

      // UI設定の検証
      if (settings.ui) {
        // テーマの検証
        if (settings.ui.theme &&
          !APP_CONSTANTS.SUPPORTED_THEMES.includes(settings.ui.theme)) {
          errors.push('サポートされていないテーマです');
        }
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('設定妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['設定の妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * 設定の変更を監視
   * @param {Function} callback - 変更時のコールバック関数
   */
  onSettingsChanged(callback) {
    try {
      this.repository.onSettingsChanged(callback);
    } catch (error) {
      this.logger.error('設定変更監視の登録中にエラーが発生しました', error);
      throw error;
    }
  }
}

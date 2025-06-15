/*
 * AI業務支援ツール Edge拡張機能 - 設定管理ユースケース
 * Copyright (c) 2024 AI Business Support Team
 */

import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';

/**
 * 設定管理ユースケース
 * 設定の読み込み、保存、検証、テストを統合管理
 */
export class ManageSettingsUseCase {
  /**
   * コンストラクタ
   * @param {Object} services - サービス群
   */
  constructor(services) {
    this.settingsService = services.settings;
    this.apiService = services.api;
    this.notificationService = services.notification;
    this.messagingService = services.messaging;

    this.logger = new Logger('ManageSettingsUseCase');
    this.validator = new Validator();
  }

  /**
   * 設定を取得
   * @returns {Promise<Object>} 設定データ
   */
  async getSettings() {
    try {
      this.logger.info('設定取得処理を開始します');

      const settings = await this.settingsService.getSettings();

      this.logger.info('設定取得処理が完了しました');
      return {
        success: true,
        data: settings.toJSON()
      };
    } catch (error) {
      this.logger.error('設定取得処理中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * API設定を更新
   * @param {Object} apiSettings - API設定
   * @param {Object} options - オプション
   * @returns {Promise<Object>} 更新結果
   */
  async updateApiSettings(apiSettings, options = {}) {
    let notificationId = null;

    try {
      this.logger.info('API設定更新処理を開始します', { provider: apiSettings.provider });

      // 進捗通知を表示
      notificationId = await this.notificationService.showProgress(
        'API設定を更新しています...', 0, { title: '設定更新' }
      );

      // 設定の妥当性検証
      await this.notificationService.updateProgress(notificationId, 20, '設定を検証しています...');
      const validationResult = this.validateApiSettings(apiSettings);
      if (!validationResult.isValid) {
        throw new Error(`API設定が無効です: ${validationResult.errors.join(', ')}`);
      }

      // 接続テスト（オプション）
      if (options.testConnection !== false) {
        await this.notificationService.updateProgress(notificationId, 40, 'API接続をテストしています...');
        const testResult = await this.testApiConnection(apiSettings);
        if (!testResult.success) {
          if (options.skipTestOnFailure !== true) {
            throw new Error(`API接続テストに失敗しました: ${testResult.message}`);
          } else {
            await this.notificationService.showWarning(
              `API接続テストに失敗しましたが、設定を保存します: ${testResult.message}`
            );
          }
        }
      }

      // 設定を保存
      await this.notificationService.updateProgress(notificationId, 80, '設定を保存しています...');
      await this.settingsService.updateApiSettings(apiSettings);

      // 他のコンポーネントに設定変更を通知
      await this.notifySettingsChanged('api', apiSettings);

      // 完了通知
      await this.notificationService.hideNotification(notificationId);
      await this.notificationService.showSuccess('API設定が正常に更新されました');

      this.logger.info('API設定更新処理が完了しました');
      return {
        success: true,
        message: 'API設定が正常に更新されました'
      };

    } catch (error) {
      if (notificationId) {
        await this.notificationService.hideNotification(notificationId);
      }
      await this.notificationService.showError(`API設定の更新に失敗しました: ${error.message}`);

      this.logger.error('API設定更新処理中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * UI設定を更新
   * @param {Object} uiSettings - UI設定
   * @returns {Promise<Object>} 更新結果
   */
  async updateUiSettings(uiSettings) {
    try {
      this.logger.info('UI設定更新処理を開始します');

      // 設定の妥当性検証
      const validationResult = this.validateUiSettings(uiSettings);
      if (!validationResult.isValid) {
        throw new Error(`UI設定が無効です: ${validationResult.errors.join(', ')}`);
      }

      // 設定を保存
      await this.settingsService.updateUiSettings(uiSettings);

      // 他のコンポーネントに設定変更を通知
      await this.notifySettingsChanged('ui', uiSettings);

      await this.notificationService.showSuccess('UI設定が正常に更新されました');

      this.logger.info('UI設定更新処理が完了しました');
      return {
        success: true,
        message: 'UI設定が正常に更新されました'
      };

    } catch (error) {
      await this.notificationService.showError(`UI設定の更新に失敗しました: ${error.message}`);

      this.logger.error('UI設定更新処理中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * 設定を一括更新
   * @param {Object} allSettings - 全設定
   * @param {Object} options - オプション
   * @returns {Promise<Object>} 更新結果
   */
  async updateAllSettings(allSettings, options = {}) {
    let notificationId = null;

    try {
      this.logger.info('設定一括更新処理を開始します');

      notificationId = await this.notificationService.showProgress(
        '設定を一括更新しています...', 0, { title: '設定一括更新' }
      );

      const results = [];

      // API設定の更新
      if (allSettings.api) {
        await this.notificationService.updateProgress(notificationId, 30, 'API設定を更新中...');
        const apiResult = await this.updateApiSettings(allSettings.api, {
          ...options,
          showNotifications: false
        });
        results.push({ type: 'api', ...apiResult });
      }

      // UI設定の更新
      if (allSettings.ui) {
        await this.notificationService.updateProgress(notificationId, 70, 'UI設定を更新中...');
        const uiResult = await this.updateUiSettings(allSettings.ui);
        results.push({ type: 'ui', ...uiResult });
      }

      const failedUpdates = results.filter(r => !r.success);

      await this.notificationService.hideNotification(notificationId);

      if (failedUpdates.length === 0) {
        await this.notificationService.showSuccess('全ての設定が正常に更新されました');
        this.logger.info('設定一括更新処理が完了しました');
        return {
          success: true,
          message: '全ての設定が正常に更新されました',
          results
        };
      } else {
        const errorMessage = `一部の設定更新に失敗しました: ${failedUpdates.map(r => r.error).join(', ')}`;
        await this.notificationService.showWarning(errorMessage);
        return {
          success: false,
          error: errorMessage,
          results
        };
      }

    } catch (error) {
      if (notificationId) {
        await this.notificationService.hideNotification(notificationId);
      }
      await this.notificationService.showError(`設定一括更新に失敗しました: ${error.message}`);

      this.logger.error('設定一括更新処理中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * 設定を初期化
   * @param {Object} options - オプション
   * @returns {Promise<Object>} 初期化結果
   */
  async resetSettings(options = {}) {
    try {
      this.logger.info('設定初期化処理を開始します');

      // 確認が必要な場合
      if (options.requireConfirmation !== false) {
        const confirmed = await this.showConfirmDialog(
          '設定をリセットしますか？',
          '全ての設定がデフォルト値に戻ります。この操作は元に戻せません。'
        );
        if (!confirmed) {
          return {
            success: false,
            error: 'ユーザーによりキャンセルされました'
          };
        }
      }

      // 設定を初期化
      await this.settingsService.resetSettings();

      // 他のコンポーネントに設定変更を通知
      await this.notifySettingsChanged('reset', {});

      await this.notificationService.showSuccess('設定が正常に初期化されました');

      this.logger.info('設定初期化処理が完了しました');
      return {
        success: true,
        message: '設定が正常に初期化されました'
      };

    } catch (error) {
      await this.notificationService.showError(`設定初期化に失敗しました: ${error.message}`);

      this.logger.error('設定初期化処理中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * API接続をテスト
   * @param {Object} apiSettings - API設定
   * @returns {Promise<Object>} テスト結果
   */
  async testApiConnection(apiSettings) {
    try {
      this.logger.info('API接続テストを開始します', { provider: apiSettings.provider });

      const testResult = await this.apiService.testConnection(apiSettings);

      this.logger.info('API接続テストが完了しました', { success: testResult.success });
      return testResult;

    } catch (error) {
      this.logger.error('API接続テスト中にエラーが発生しました', error);
      return {
        success: false,
        message: error.message,
        latency: 0
      };
    }
  }

  /**
   * 利用可能なモデル一覧を取得
   * @param {Object} apiSettings - API設定
   * @returns {Promise<Object>} モデル一覧
   */
  async getAvailableModels(apiSettings) {
    try {
      this.logger.info('利用可能モデル一覧取得を開始します', { provider: apiSettings.provider });

      const models = await this.apiService.getAvailableModels(apiSettings);

      this.logger.info('利用可能モデル一覧取得が完了しました', { modelCount: models.length });
      return {
        success: true,
        data: models
      };

    } catch (error) {
      this.logger.error('利用可能モデル一覧取得中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message,
        data: []
      };
    }
  }

  /**
   * 設定をエクスポート
   * @param {Object} options - エクスポートオプション
   * @returns {Promise<Object>} エクスポートデータ
   */
  async exportSettings(options = {}) {
    try {
      this.logger.info('設定エクスポート処理を開始します');

      const settings = await this.settingsService.getSettings();
      const exportData = {
        version: '1.0',
        exportedAt: new Date().toISOString(),
        settings: settings.toJSON(),
        metadata: {
          userAgent: navigator.userAgent,
          extensionVersion: chrome.runtime.getManifest().version
        }
      };

      // 機密情報を除去（オプション）
      if (options.excludeSensitive !== false) {
        if (exportData.settings.api?.apiKey) {
          exportData.settings.api.apiKey = '[REDACTED]';
        }
      }

      this.logger.info('設定エクスポート処理が完了しました');
      return {
        success: true,
        data: exportData
      };

    } catch (error) {
      this.logger.error('設定エクスポート処理中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * 設定をインポート
   * @param {Object} importData - インポートデータ
   * @param {Object} options - インポートオプション
   * @returns {Promise<Object>} インポート結果
   */
  async importSettings(importData, options = {}) {
    try {
      this.logger.info('設定インポート処理を開始します');

      // インポートデータの妥当性検証
      const validationResult = this.validateImportData(importData);
      if (!validationResult.isValid) {
        throw new Error(`インポートデータが無効です: ${validationResult.errors.join(', ')}`);
      }

      // 設定を更新
      const updateResult = await this.updateAllSettings(importData.settings, {
        ...options,
        testConnection: options.testConnection !== false
      });

      if (updateResult.success) {
        await this.notificationService.showSuccess('設定が正常にインポートされました');
      }

      this.logger.info('設定インポート処理が完了しました');
      return updateResult;

    } catch (error) {
      await this.notificationService.showError(`設定インポートに失敗しました: ${error.message}`);

      this.logger.error('設定インポート処理中にエラーが発生しました', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * API設定の妥当性を検証
   * @param {Object} apiSettings - API設定
   * @returns {Object} 検証結果
   * @private
   */
  validateApiSettings(apiSettings) {
    // settingsService の validateSettings メソッドを活用
    const tempSettings = { api: apiSettings, ui: {} };
    return this.settingsService.validateSettings(tempSettings);
  }

  /**
   * UI設定の妥当性を検証
   * @param {Object} uiSettings - UI設定
   * @returns {Object} 検証結果
   * @private
   */
  validateUiSettings(uiSettings) {
    const errors = [];

    try {
      // テーマの検証
      if (uiSettings.theme && !['light', 'dark', 'auto'].includes(uiSettings.theme)) {
        errors.push('無効なテーマが指定されています');
      }

      // 各種フラグの検証
      const booleanFields = ['autoDetect', 'showNotifications', 'saveHistory'];
      for (const field of booleanFields) {
        if (uiSettings[field] !== undefined && typeof uiSettings[field] !== 'boolean') {
          errors.push(`${field}の値が無効です（boolean値である必要があります）`);
        }
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('UI設定妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['UI設定の妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * インポートデータの妥当性を検証
   * @param {Object} importData - インポートデータ
   * @returns {Object} 検証結果
   * @private
   */
  validateImportData(importData) {
    const errors = [];

    try {
      // 基本構造の確認
      if (!this.validator.isObject(importData)) {
        errors.push('インポートデータの形式が無効です');
        return { isValid: false, errors };
      }

      if (!importData.settings) {
        errors.push('設定データが含まれていません');
        return { isValid: false, errors };
      }

      // バージョンの確認
      if (importData.version && importData.version !== '1.0') {
        errors.push('サポートされていないバージョンです');
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('インポートデータ妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['インポートデータの妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * 設定変更を他のコンポーネントに通知
   * @param {string} settingType - 設定タイプ
   * @param {Object} settingData - 設定データ
   * @private
   */
  async notifySettingsChanged(settingType, settingData) {
    try {
      await this.messagingService.broadcast('settingsChanged', {
        type: settingType,
        data: settingData,
        timestamp: new Date().toISOString()
      });
    } catch (error) {
      this.logger.error('設定変更通知中にエラーが発生しました', error);
      // 通知の失敗は処理を止めない
    }
  }

  /**
   * 確認ダイアログを表示
   * @param {string} title - タイトル
   * @param {string} message - メッセージ
   * @returns {Promise<boolean>} 確認結果
   * @private
   */
  async showConfirmDialog(title, message) {
    try {
      // 実装はプレゼンテーション層で行う
      // ここでは仮の実装
      return confirm(`${title}\n\n${message}`);
    } catch (error) {
      this.logger.error('確認ダイアログ表示中にエラーが発生しました', error);
      return false;
    }
  }
}

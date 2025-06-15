/*
 * AI業務支援ツール Edge拡張機能 - テキスト処理ユースケース
 * Copyright (c) 2024 AI Business Support Team
 */

import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';

/**
 * テキスト処理ユースケース
 * メインのAI処理機能を統合管理
 */
export class ProcessTextUseCase {
  /**
   * コンストラクタ
   * @param {Object} services - サービス群
   */
  constructor(services) {
    this.apiService = services.api;
    this.promptService = services.prompt;
    this.historyService = services.history;
    this.statisticsService = services.statistics;
    this.settingsService = services.settings;
    this.notificationService = services.notification;

    this.logger = new Logger('ProcessTextUseCase');
    this.validator = new Validator();
  }

  /**
   * テキストを処理（メイン処理）
   * @param {Object} request - 処理リクエスト
   * @param {string} request.text - 処理対象テキスト
   * @param {string} [request.templateId] - プロンプトテンプレートID
   * @param {Object} [request.templateParams] - テンプレートパラメータ
   * @param {string} [request.customPrompt] - カスタムプロンプト
   * @param {Object} [request.options] - 処理オプション
   * @returns {Promise<Object>} 処理結果
   */
  async execute(request) {
    let notificationId = null;
    const startTime = Date.now();

    try {
      this.logger.info('テキスト処理を開始します', {
        textLength: request.text?.length,
        templateId: request.templateId,
        hasCustomPrompt: !!request.customPrompt
      });

      // 入力の妥当性検証
      const validationResult = this.validateRequest(request);
      if (!validationResult.isValid) {
        throw new Error(`リクエストが無効です: ${validationResult.errors.join(', ')}`);
      }

      // 進捗通知を表示
      notificationId = await this.notificationService.showProgress(
        'AI処理を開始しています...', 0, { title: 'テキスト処理' }
      );

      // 設定を取得
      await this.notificationService.updateProgress(notificationId, 10, '設定を読み込んでいます...');
      const settings = await this.settingsService.getSettings();

      // プロンプトを生成
      await this.notificationService.updateProgress(notificationId, 20, 'プロンプトを生成しています...');
      const promptData = await this.generatePrompt(request);

      // 統計記録（リクエスト開始）
      const requestStats = this.createRequestStats(request, settings, promptData);
      await this.statisticsService.recordRequest(requestStats);

      // API呼び出し
      await this.notificationService.updateProgress(notificationId, 40, 'AI APIを呼び出しています...');
      const apiResponse = await this.apiService.callAI(
        promptData.prompt,
        settings.api,
        {
          systemPrompt: promptData.systemPrompt,
          ...request.options
        }
      );

      if (!apiResponse.success) {
        throw new Error(`AI処理に失敗しました: ${apiResponse.error}`);
      }

      // 結果を処理
      await this.notificationService.updateProgress(notificationId, 80, '結果を処理しています...');
      const result = await this.processResult(request, promptData, apiResponse, settings);

      // 履歴に保存
      if (settings.ui?.saveHistory !== false) {
        await this.saveToHistory(request, promptData, result);
      }

      // 統計記録（成功）
      await this.statisticsService.recordSuccess(requestStats, {
        tokensUsed: apiResponse.metadata?.tokensUsed || 0,
        responseTime: apiResponse.metadata?.duration || 0
      });

      // 完了通知
      await this.notificationService.hideNotification(notificationId);
      await this.notificationService.showSuccess('AI処理が完了しました');

      const duration = Date.now() - startTime;
      this.logger.info('テキスト処理が完了しました', { duration: `${duration}ms` });

      return {
        success: true,
        data: result,
        metadata: {
          duration,
          tokensUsed: apiResponse.metadata?.tokensUsed || 0,
          timestamp: new Date().toISOString()
        }
      };

    } catch (error) {
      // エラー処理
      if (notificationId) {
        await this.notificationService.hideNotification(notificationId);
      }

      await this.notificationService.showError(`処理中にエラーが発生しました: ${error.message}`);

      // 統計記録（失敗）
      if (requestStats) {
        await this.statisticsService.recordFailure(requestStats, {
          errorType: error.name || 'Error',
          errorMessage: error.message
        });
      }

      const duration = Date.now() - startTime;
      this.logger.error('テキスト処理中にエラーが発生しました', error, { duration: `${duration}ms` });

      return {
        success: false,
        error: error.message,
        metadata: {
          duration,
          timestamp: new Date().toISOString()
        }
      };
    }
  }

  /**
   * クイック処理（簡易版）
   * @param {string} text - 処理対象テキスト
   * @param {string} templateId - テンプレートID
   * @param {Object} templateParams - テンプレートパラメータ
   * @returns {Promise<Object>} 処理結果
   */
  async quickProcess(text, templateId, templateParams = {}) {
    return this.execute({
      text,
      templateId,
      templateParams,
      options: { saveToHistory: true }
    });
  }

  /**
   * カスタム処理
   * @param {string} text - 処理対象テキスト
   * @param {string} customPrompt - カスタムプロンプト
   * @param {Object} options - オプション
   * @returns {Promise<Object>} 処理結果
   */
  async customProcess(text, customPrompt, options = {}) {
    return this.execute({
      text,
      customPrompt,
      options
    });
  }

  /**
   * バッチ処理
   * @param {Array} textItems - テキスト項目の配列
   * @param {string} templateId - テンプレートID
   * @param {Object} options - オプション
   * @returns {Promise<Object>} 処理結果
   */
  async batchProcess(textItems, templateId, options = {}) {
    try {
      this.logger.info('バッチ処理を開始します', { itemCount: textItems.length });

      const notificationId = await this.notificationService.showProgress(
        `${textItems.length}件のテキストを処理しています...`, 0, { title: 'バッチ処理' }
      );

      const results = [];
      const batchSize = options.batchSize || 5; // 同時実行数制限

      for (let i = 0; i < textItems.length; i += batchSize) {
        const batch = textItems.slice(i, i + batchSize);
        const progress = Math.round((i / textItems.length) * 90);

        await this.notificationService.updateProgress(
          notificationId,
          progress,
          `${i + 1}-${Math.min(i + batchSize, textItems.length)}件目を処理中...`
        );

        const batchPromises = batch.map(item =>
          this.quickProcess(item.text, templateId, item.params || {})
            .catch(error => ({ success: false, error: error.message, item }))
        );

        const batchResults = await Promise.all(batchPromises);
        results.push(...batchResults);

        // レート制限対策の待機
        if (i + batchSize < textItems.length) {
          await this.sleep(1000);
        }
      }

      await this.notificationService.updateProgress(notificationId, 100, '処理が完了しました');
      await this.notificationService.hideNotification(notificationId);

      const successCount = results.filter(r => r.success).length;
      const failCount = results.length - successCount;

      if (failCount === 0) {
        await this.notificationService.showSuccess(`バッチ処理が完了しました（${successCount}件成功）`);
      } else {
        await this.notificationService.showWarning(
          `バッチ処理が完了しました（${successCount}件成功、${failCount}件失敗）`
        );
      }

      this.logger.info('バッチ処理が完了しました', { successCount, failCount });

      return {
        success: true,
        data: {
          results,
          summary: { successCount, failCount, total: results.length }
        }
      };

    } catch (error) {
      this.logger.error('バッチ処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * リクエストの妥当性を検証
   * @param {Object} request - リクエスト
   * @returns {Object} 検証結果
   * @private
   */
  validateRequest(request) {
    const errors = [];

    try {
      // テキストの検証
      if (!this.validator.isString(request.text) || request.text.trim().length === 0) {
        errors.push('処理対象テキストが無効です');
      }

      // プロンプト設定の検証
      const hasTemplate = request.templateId && this.validator.isString(request.templateId);
      const hasCustomPrompt = request.customPrompt && this.validator.isString(request.customPrompt);

      if (!hasTemplate && !hasCustomPrompt) {
        errors.push('テンプレートIDまたはカスタムプロンプトのいずれかが必要です');
      }

      if (hasTemplate && hasCustomPrompt) {
        errors.push('テンプレートIDとカスタムプロンプトは同時に指定できません');
      }

      // テンプレートパラメータの検証
      if (request.templateParams && !this.validator.isObject(request.templateParams)) {
        errors.push('テンプレートパラメータの形式が無効です');
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('リクエスト妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['リクエストの妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * プロンプトを生成
   * @param {Object} request - リクエスト
   * @returns {Promise<Object>} プロンプトデータ
   * @private
   */
  async generatePrompt(request) {
    try {
      if (request.templateId) {
        // テンプレートからプロンプトを生成
        const params = { content: request.text, ...request.templateParams };
        return await this.promptService.generatePrompt(request.templateId, params);
      } else {
        // カスタムプロンプトを使用
        return this.promptService.createCustomPrompt(request.customPrompt);
      }
    } catch (error) {
      this.logger.error('プロンプト生成中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 結果を処理
   * @param {Object} request - リクエスト
   * @param {Object} promptData - プロンプトデータ
   * @param {Object} apiResponse - API応答
   * @param {Object} settings - 設定
   * @returns {Object} 処理済み結果
   * @private
   */
  async processResult(request, promptData, apiResponse, settings) {
    try {
      return {
        originalText: request.text,
        processedText: apiResponse.data,
        promptUsed: promptData.prompt,
        systemPromptUsed: promptData.systemPrompt,
        templateUsed: promptData.template?.id || null,
        processingTime: apiResponse.metadata?.duration || 0,
        tokensUsed: apiResponse.metadata?.tokensUsed || 0,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      this.logger.error('結果処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴に保存
   * @param {Object} request - リクエスト
   * @param {Object} promptData - プロンプトデータ
   * @param {Object} result - 処理結果
   * @private
   */
  async saveToHistory(request, promptData, result) {
    try {
      const historyItem = {
        type: 'text_processing',
        title: this.generateHistoryTitle(request, promptData),
        content: request.text,
        result: result.processedText,
        service: 'text_processor',
        metadata: {
          templateId: request.templateId,
          templateParams: request.templateParams,
          customPrompt: request.customPrompt,
          tokensUsed: result.tokensUsed,
          processingTime: result.processingTime
        }
      };

      await this.historyService.addHistoryItem(historyItem);
    } catch (error) {
      this.logger.error('履歴保存中にエラーが発生しました', error);
      // 履歴保存の失敗は処理を止めない
    }
  }

  /**
   * 履歴タイトルを生成
   * @param {Object} request - リクエスト
   * @param {Object} promptData - プロンプトデータ
   * @returns {string} タイトル
   * @private
   */
  generateHistoryTitle(request, promptData) {
    if (promptData.template) {
      return `${promptData.template.name} - ${request.text.substring(0, 30)}...`;
    } else {
      return `カスタム処理 - ${request.text.substring(0, 30)}...`;
    }
  }

  /**
   * リクエスト統計を作成
   * @param {Object} request - リクエスト
   * @param {Object} settings - 設定
   * @param {Object} promptData - プロンプトデータ
   * @returns {Object} リクエスト統計
   * @private
   */
  createRequestStats(request, settings, promptData) {
    return {
      model: settings.api.model,
      provider: settings.api.provider,
      type: 'text_processing',
      inputLength: request.text.length,
      promptLength: promptData.prompt.length,
      templateId: request.templateId || 'custom'
    };
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

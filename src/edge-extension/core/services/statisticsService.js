/*
 * AI業務支援ツール Edge拡張機能 - 統計サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { StatisticsRepository } from '../repositories/statisticsRepository.js';
import { Statistics } from '../entities/statistics.js';
import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';

/**
 * 統計管理サービス
 */
export class StatisticsService {
  /**
   * コンストラクタ
   */
  constructor() {
    this.repository = new StatisticsRepository();
    this.logger = new Logger('StatisticsService');
    this.validator = new Validator();
  }

  /**
   * 統計情報を取得
   * @returns {Promise<Statistics>} 統計オブジェクト
   */
  async getStatistics() {
    try {
      this.logger.info('統計情報取得処理を開始します');

      const statisticsData = await this.repository.getStatistics();
      const statistics = Statistics.fromJSON(statisticsData);

      this.logger.info('統計情報取得処理が完了しました');
      return statistics;
    } catch (error) {
      this.logger.error('統計情報取得処理中にエラーが発生しました', error);

      // デフォルト統計を返す
      this.logger.info('デフォルト統計を返します');
      return Statistics.createDefault();
    }
  }

  /**
   * リクエスト実行時の統計を更新
   * @param {Object} requestData - リクエストデータ
   * @returns {Promise<boolean>} 更新成功フラグ
   */
  async recordRequest(requestData) {
    try {
      this.logger.debug('リクエスト統計記録処理を開始します');

      // リクエストデータの妥当性検証
      const validationResult = this.validateRequestData(requestData);
      if (!validationResult.isValid) {
        throw new Error(`リクエストデータの妥当性検証に失敗しました: ${validationResult.errors.join(', ')}`);
      }

      const statistics = await this.getStatistics();
      const updatedStatistics = statistics.recordRequest(requestData);

      await this.repository.saveStatistics(updatedStatistics.toJSON());

      this.logger.debug('リクエスト統計記録処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('リクエスト統計記録処理中にエラーが発生しました', error);
      // 統計記録の失敗は処理を止めない
      return false;
    }
  }

  /**
   * 成功したリクエストの統計を更新
   * @param {Object} requestData - リクエストデータ
   * @param {Object} responseData - レスポンスデータ
   * @returns {Promise<boolean>} 更新成功フラグ
   */
  async recordSuccess(requestData, responseData) {
    try {
      this.logger.debug('成功統計記録処理を開始します');

      const statistics = await this.getStatistics();
      const updatedStatistics = statistics.recordSuccess(requestData, responseData);

      await this.repository.saveStatistics(updatedStatistics.toJSON());

      this.logger.debug('成功統計記録処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('成功統計記録処理中にエラーが発生しました', error);
      return false;
    }
  }

  /**
   * 失敗したリクエストの統計を更新
   * @param {Object} requestData - リクエストデータ
   * @param {Object} errorData - エラーデータ
   * @returns {Promise<boolean>} 更新成功フラグ
   */
  async recordFailure(requestData, errorData) {
    try {
      this.logger.debug('失敗統計記録処理を開始します');

      const statistics = await this.getStatistics();
      const updatedStatistics = statistics.recordFailure(requestData, errorData);

      await this.repository.saveStatistics(updatedStatistics.toJSON());

      this.logger.debug('失敗統計記録処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('失敗統計記録処理中にエラーが発生しました', error);
      return false;
    }
  }

  /**
   * 使用量統計を取得
   * @param {Object} options - 取得オプション
   * @param {string} [options.period] - 期間フィルター ('daily'|'weekly'|'monthly')
   * @param {string} [options.model] - モデルフィルター
   * @returns {Promise<Object>} 使用量統計
   */
  async getUsageStatistics(options = {}) {
    try {
      this.logger.info('使用量統計取得処理を開始します', options);

      const statistics = await this.getStatistics();
      const usageStats = statistics.getUsageStatistics(options);

      this.logger.info('使用量統計取得処理が完了しました');
      return usageStats;
    } catch (error) {
      this.logger.error('使用量統計取得処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * パフォーマンス統計を取得
   * @param {Object} options - 取得オプション
   * @returns {Promise<Object>} パフォーマンス統計
   */
  async getPerformanceStatistics(options = {}) {
    try {
      this.logger.info('パフォーマンス統計取得処理を開始します', options);

      const statistics = await this.getStatistics();
      const performanceStats = statistics.getPerformanceStatistics(options);

      this.logger.info('パフォーマンス統計取得処理が完了しました');
      return performanceStats;
    } catch (error) {
      this.logger.error('パフォーマンス統計取得処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * エラー統計を取得
   * @param {Object} options - 取得オプション
   * @returns {Promise<Object>} エラー統計
   */
  async getErrorStatistics(options = {}) {
    try {
      this.logger.info('エラー統計取得処理を開始します', options);

      const statistics = await this.getStatistics();
      const errorStats = statistics.getErrorStatistics(options);

      this.logger.info('エラー統計取得処理が完了しました');
      return errorStats;
    } catch (error) {
      this.logger.error('エラー統計取得処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 統計を初期化
   * @returns {Promise<boolean>} 初期化成功フラグ
   */
  async resetStatistics() {
    try {
      this.logger.info('統計初期化処理を開始します');

      const defaultStatistics = Statistics.createDefault();
      await this.repository.saveStatistics(defaultStatistics.toJSON());

      this.logger.info('統計初期化処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('統計初期化処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 古い統計データを削除
   * @param {number} retentionDays - 保持日数
   * @returns {Promise<boolean>} 削除成功フラグ
   */
  async cleanupOldStatistics(retentionDays = 90) {
    try {
      this.logger.info('古い統計データ削除処理を開始します', { retentionDays });

      if (!this.validator.isPositiveInteger(retentionDays)) {
        throw new Error('保持日数が無効です');
      }

      const statistics = await this.getStatistics();
      const cleanedStatistics = statistics.cleanup(retentionDays);

      await this.repository.saveStatistics(cleanedStatistics.toJSON());

      this.logger.info('古い統計データ削除処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('古い統計データ削除処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 統計レポートを生成
   * @param {Object} options - レポートオプション
   * @param {string} [options.format] - レポート形式 ('summary'|'detailed')
   * @param {string} [options.period] - レポート期間
   * @returns {Promise<Object>} 統計レポート
   */
  async generateReport(options = {}) {
    try {
      this.logger.info('統計レポート生成処理を開始します', options);

      const statistics = await this.getStatistics();
      const usageStats = await this.getUsageStatistics(options);
      const performanceStats = await this.getPerformanceStatistics(options);
      const errorStats = await this.getErrorStatistics(options);

      const report = {
        generatedAt: new Date().toISOString(),
        period: options.period || 'all',
        format: options.format || 'summary',
        summary: {
          totalRequests: statistics.totalRequests,
          successRate: statistics.getSuccessRate(),
          averageResponseTime: performanceStats.averageResponseTime,
          totalTokensUsed: statistics.totalTokensUsed
        },
        usage: usageStats,
        performance: performanceStats,
        errors: errorStats
      };

      if (options.format === 'detailed') {
        report.rawStatistics = statistics.toJSON();
      }

      this.logger.info('統計レポート生成処理が完了しました');
      return report;
    } catch (error) {
      this.logger.error('統計レポート生成処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * リクエストデータの妥当性を検証
   * @param {Object} requestData - リクエストデータ
   * @returns {Object} 検証結果
   * @private
   */
  validateRequestData(requestData) {
    const errors = [];

    try {
      // 必須フィールドの検証
      if (!this.validator.isString(requestData.model) || requestData.model.trim().length === 0) {
        errors.push('モデル名が無効です');
      }

      if (!this.validator.isString(requestData.provider) || requestData.provider.trim().length === 0) {
        errors.push('プロバイダー名が無効です');
      }

      if (!this.validator.isString(requestData.type) || requestData.type.trim().length === 0) {
        errors.push('リクエストタイプが無効です');
      }

      // 数値フィールドの検証
      if (requestData.inputTokens !== undefined && !this.validator.isNonNegativeInteger(requestData.inputTokens)) {
        errors.push('入力トークン数が無効です');
      }

      if (requestData.outputTokens !== undefined && !this.validator.isNonNegativeInteger(requestData.outputTokens)) {
        errors.push('出力トークン数が無効です');
      }

      if (requestData.responseTime !== undefined && !this.validator.isNonNegativeNumber(requestData.responseTime)) {
        errors.push('レスポンス時間が無効です');
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('リクエストデータ妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['リクエストデータの妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * 統計の変更を監視
   * @param {Function} callback - 変更時のコールバック関数
   */
  onStatisticsChanged(callback) {
    try {
      this.repository.onStatisticsChanged(callback);
    } catch (error) {
      this.logger.error('統計変更監視の登録中にエラーが発生しました', error);
      throw error;
    }
  }
}

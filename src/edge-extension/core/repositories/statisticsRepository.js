/*
 * AI業務支援ツール Edge拡張機能 - 統計リポジトリ
 * Copyright (c) 2024 AI Business Support Team
 */

import { StorageRepository } from '../../infrastructure/storage/storageRepository.js';
import { Logger } from '../../shared/utils/logger.js';

/**
 * 統計データリポジトリ
 */
export class StatisticsRepository {
  /**
   * コンストラクタ
   */
  constructor() {
    this.storageRepository = new StorageRepository();
    this.logger = new Logger('StatisticsRepository');
    this.STORAGE_KEY = 'ai_assistant_statistics';
  }

  /**
   * 統計データを取得
   * @returns {Promise<Object>} 統計データ
   */
  async getStatistics() {
    try {
      this.logger.debug('統計データ取得を開始します');

      const statisticsData = await this.storageRepository.get(this.STORAGE_KEY);

      if (!statisticsData) {
        this.logger.debug('統計データが存在しません。デフォルト統計を返します');
        return this.createDefaultStatistics();
      }

      this.logger.debug('統計データ取得が完了しました');
      return statisticsData;
    } catch (error) {
      this.logger.error('統計データ取得中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 統計データを保存
   * @param {Object} statistics - 統計データ
   * @returns {Promise<void>}
   */
  async saveStatistics(statistics) {
    try {
      this.logger.debug('統計データ保存を開始します');

      // 最終更新日時を設定
      statistics.lastUpdated = new Date().toISOString();

      await this.storageRepository.set(this.STORAGE_KEY, statistics);

      this.logger.debug('統計データ保存が完了しました');
    } catch (error) {
      this.logger.error('統計データ保存中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 統計データを初期化
   * @returns {Promise<void>}
   */
  async resetStatistics() {
    try {
      this.logger.debug('統計データ初期化を開始します');

      const defaultStatistics = this.createDefaultStatistics();
      await this.storageRepository.set(this.STORAGE_KEY, defaultStatistics);

      this.logger.debug('統計データ初期化が完了しました');
    } catch (error) {
      this.logger.error('統計データ初期化中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 統計データを削除
   * @returns {Promise<void>}
   */
  async deleteStatistics() {
    try {
      this.logger.debug('統計データ削除を開始します');

      await this.storageRepository.remove(this.STORAGE_KEY);

      this.logger.debug('統計データ削除が完了しました');
    } catch (error) {
      this.logger.error('統計データ削除中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 統計データのバックアップを作成
   * @returns {Promise<Object>} バックアップデータ
   */
  async backupStatistics() {
    try {
      this.logger.debug('統計データバックアップを開始します');

      const statistics = await this.getStatistics();
      const backupData = {
        ...statistics,
        backupCreatedAt: new Date().toISOString()
      };

      this.logger.debug('統計データバックアップが完了しました');
      return backupData;
    } catch (error) {
      this.logger.error('統計データバックアップ中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 統計データを復元
   * @param {Object} backupData - バックアップデータ
   * @returns {Promise<void>}
   */
  async restoreStatistics(backupData) {
    try {
      this.logger.debug('統計データ復元を開始します');

      // バックアップデータの妥当性検証
      if (!this.validateBackupData(backupData)) {
        throw new Error('バックアップデータの形式が無効です');
      }

      // バックアップ固有のフィールドを除去
      const { backupCreatedAt, ...statisticsData } = backupData;

      await this.saveStatistics(statisticsData);

      this.logger.debug('統計データ復元が完了しました');
    } catch (error) {
      this.logger.error('統計データ復元中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * デフォルト統計データを作成
   * @returns {Object} デフォルト統計データ
   * @private
   */
  createDefaultStatistics() {
    return {
      totalRequests: 0,
      successfulRequests: 0,
      failedRequests: 0,
      totalTokensUsed: 0,
      totalInputTokens: 0,
      totalOutputTokens: 0,
      usageByModel: {},
      usageByProvider: {},
      usageByType: {},
      dailyUsage: {},
      responseTimeStats: {
        total: 0,
        count: 0,
        min: null,
        max: null,
        average: 0
      },
      errorStats: {
        byType: {},
        byProvider: {},
        byModel: {},
        recentErrors: []
      },
      createdAt: new Date().toISOString(),
      lastUpdated: new Date().toISOString()
    };
  }

  /**
   * バックアップデータの妥当性を検証
   * @param {Object} backupData - バックアップデータ
   * @returns {boolean} 妥当性
   * @private
   */
  validateBackupData(backupData) {
    try {
      // 必須フィールドの存在確認
      const requiredFields = [
        'totalRequests',
        'successfulRequests',
        'failedRequests',
        'totalTokensUsed',
        'usageByModel',
        'usageByProvider',
        'createdAt'
      ];

      for (const field of requiredFields) {
        if (!(field in backupData)) {
          this.logger.warn(`必須フィールドが不足しています: ${field}`);
          return false;
        }
      }

      // 数値フィールドの型確認
      const numericFields = [
        'totalRequests',
        'successfulRequests',
        'failedRequests',
        'totalTokensUsed'
      ];

      for (const field of numericFields) {
        if (typeof backupData[field] !== 'number' || backupData[field] < 0) {
          this.logger.warn(`数値フィールドの値が無効です: ${field}`);
          return false;
        }
      }

      // オブジェクトフィールドの型確認
      const objectFields = ['usageByModel', 'usageByProvider'];
      for (const field of objectFields) {
        if (typeof backupData[field] !== 'object' || backupData[field] === null) {
          this.logger.warn(`オブジェクトフィールドの値が無効です: ${field}`);
          return false;
        }
      }

      return true;
    } catch (error) {
      this.logger.error('バックアップデータ妥当性検証中にエラーが発生しました', error);
      return false;
    }
  }

  /**
   * 統計の変更を監視
   * @param {Function} callback - 変更時のコールバック関数
   */
  onStatisticsChanged(callback) {
    try {
      this.storageRepository.onChanged(this.STORAGE_KEY, callback);
    } catch (error) {
      this.logger.error('統計変更監視の登録中にエラーが発生しました', error);
      throw error;
    }
  }
}

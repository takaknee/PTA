/*
 * AI業務支援ツール Edge拡張機能 - 統計情報サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { Statistics } from '../lib/models.js';
import { STORAGE_KEYS } from '../lib/constants.js';
import { Logger, formatDateTime, safeJsonParse, safeJsonStringify } from '../lib/utils.js';

/**
 * 統計情報管理サービス
 */
export class StatisticsService {
  constructor() {
    this.cache = null;
    this.listeners = new Set();
  }

  /**
   * 統計情報を取得
   * @param {boolean} useCache キャッシュを使用するか
   * @returns {Promise<Statistics>} 統計情報
   */
  async getStatistics(useCache = true) {
    try {
      // キャッシュがある場合はキャッシュを返す
      if (useCache && this.cache) {
        Logger.debug('統計情報をキャッシュから取得');
        return this.cache;
      }

      Logger.debug('統計情報をストレージから取得');

      // Chrome storage から統計情報を取得
      const result = await chrome.storage.local.get(STORAGE_KEYS.STATISTICS);
      const storedStats = result[STORAGE_KEYS.STATISTICS] || {};

      // Statisticsオブジェクトを作成
      const statistics = new Statistics(storedStats);

      // キャッシュに保存
      this.cache = statistics;

      Logger.debug('統計情報取得完了', {
        totalRequests: statistics.totalRequests,
        successRate: statistics.getSuccessRate()
      });

      return statistics;

    } catch (error) {
      Logger.error('統計情報取得エラー', error);

      // エラー時は新しい統計情報オブジェクトを返す
      const newStats = new Statistics();
      this.cache = newStats;
      return newStats;
    }
  }

  /**
   * リクエストを記録
   * @param {string} type リクエストタイプ
   * @param {boolean} success 成功かどうか
   * @param {number} responseTime レスポンス時間（ms）
   * @param {Object} metadata 追加メタデータ
   * @returns {Promise<boolean>} 記録成功可否
   */
  async recordRequest(type, success, responseTime, metadata = {}) {
    try {
      Logger.debug('リクエスト記録開始', { type, success, responseTime });

      // 現在の統計情報を取得
      const statistics = await this.getStatistics(false);

      // リクエストを記録
      statistics.recordRequest(type, success, responseTime);

      // 追加メタデータがある場合は記録
      if (metadata.model) {
        if (!statistics.requestsByModel) {
          statistics.requestsByModel = {};
        }
        statistics.requestsByModel[metadata.model] = (statistics.requestsByModel[metadata.model] || 0) + 1;
      }

      if (metadata.provider) {
        if (!statistics.requestsByProvider) {
          statistics.requestsByProvider = {};
        }
        statistics.requestsByProvider[metadata.provider] = (statistics.requestsByProvider[metadata.provider] || 0) + 1;
      }

      // ストレージに保存
      await chrome.storage.local.set({
        [STORAGE_KEYS.STATISTICS]: statistics.toJSON()
      });

      // キャッシュを更新
      this.cache = statistics;

      // リスナーに変更を通知
      this.notifyListeners('request', { type, success, responseTime, metadata });

      Logger.debug('リクエスト記録完了', { totalRequests: statistics.totalRequests });
      return true;

    } catch (error) {
      Logger.error('リクエスト記録エラー', error);
      return false;
    }
  }

  /**
   * 統計情報をリセット
   * @returns {Promise<boolean>} リセット成功可否
   */
  async resetStatistics() {
    try {
      Logger.info('統計情報リセット開始');

      // ストレージから削除
      await chrome.storage.local.remove(STORAGE_KEYS.STATISTICS);

      // 新しい統計情報オブジェクトを作成
      const newStats = new Statistics();
      this.cache = newStats;

      // リスナーに変更を通知
      this.notifyListeners('reset');

      Logger.info('統計情報リセット完了');
      return true;

    } catch (error) {
      Logger.error('統計情報リセットエラー', error);
      return false;
    }
  }

  /**
   * 詳細な統計レポートを生成
   * @returns {Promise<Object>} 統計レポート
   */
  async generateReport() {
    try {
      const statistics = await this.getStatistics();

      const report = {
        // 基本統計
        summary: {
          totalRequests: statistics.totalRequests,
          successfulRequests: statistics.successfulRequests,
          failedRequests: statistics.failedRequests,
          successRate: statistics.getSuccessRate(),
          averageResponseTime: Math.round(statistics.averageResponseTime),
          lastRequestTime: statistics.lastRequestTime ? formatDateTime(statistics.lastRequestTime) : null
        },

        // リクエストタイプ別統計
        requestsByType: this.calculatePercentages(statistics.requestsByType, statistics.totalRequests),

        // 日別統計（過去30日）
        dailyStats: this.generateDailyStats(statistics.requestsByDay),

        // 週別統計
        weeklyStats: this.generateWeeklyStats(statistics.requestsByDay),

        // パフォーマンス統計
        performance: {
          averageResponseTime: Math.round(statistics.averageResponseTime),
          responseTimeCategory: this.categorizeResponseTime(statistics.averageResponseTime)
        },

        // 品質指標
        qualityMetrics: {
          reliability: statistics.getSuccessRate(),
          reliabilityCategory: this.categorizeReliability(statistics.getSuccessRate())
        }
      };

      // プロバイダー別・モデル別統計（存在する場合）
      if (statistics.requestsByProvider) {
        report.requestsByProvider = this.calculatePercentages(statistics.requestsByProvider, statistics.totalRequests);
      }

      if (statistics.requestsByModel) {
        report.requestsByModel = this.calculatePercentages(statistics.requestsByModel, statistics.totalRequests);
      }

      return report;

    } catch (error) {
      Logger.error('統計レポート生成エラー', error);
      return {
        summary: {
          totalRequests: 0,
          successfulRequests: 0,
          failedRequests: 0,
          successRate: 0,
          averageResponseTime: 0,
          lastRequestTime: null
        },
        requestsByType: {},
        dailyStats: [],
        weeklyStats: [],
        performance: {
          averageResponseTime: 0,
          responseTimeCategory: 'unknown'
        },
        qualityMetrics: {
          reliability: 0,
          reliabilityCategory: 'unknown'
        }
      };
    }
  }

  /**
   * 統計情報をエクスポート
   * @param {string} format エクスポート形式（json, csv）
   * @returns {Promise<string>} エクスポートされたデータ
   */
  async exportStatistics(format = 'json') {
    try {
      if (format === 'json') {
        const statistics = await this.getStatistics();
        const exportData = {
          version: '1.0',
          timestamp: new Date().toISOString(),
          statistics: statistics.toJSON()
        };
        return safeJsonStringify(exportData, '{}');
      }

      if (format === 'csv') {
        const report = await this.generateReport();
        return this.convertReportToCsv(report);
      }

      throw new Error(`サポートされていない形式: ${format}`);

    } catch (error) {
      Logger.error('統計エクスポートエラー', error);
      return format === 'csv' ? '' : '{}';
    }
  }

  /**
   * 統計変更リスナーを追加
   * @param {Function} listener リスナー関数
   */
  addChangeListener(listener) {
    if (typeof listener === 'function') {
      this.listeners.add(listener);
      Logger.debug('統計変更リスナー追加', { count: this.listeners.size });
    }
  }

  /**
   * 統計変更リスナーを削除
   * @param {Function} listener リスナー関数
   */
  removeChangeListener(listener) {
    this.listeners.delete(listener);
    Logger.debug('統計変更リスナー削除', { count: this.listeners.size });
  }

  /**
   * リスナーに変更を通知
   * @param {string} action アクション
   * @param {*} data データ
   */
  notifyListeners(action, data = null) {
    this.listeners.forEach(listener => {
      try {
        listener(action, data);
      } catch (error) {
        Logger.error('統計リスナー通知エラー', error);
      }
    });
  }

  /**
   * キャッシュをクリア
   */
  clearCache() {
    this.cache = null;
    Logger.debug('統計キャッシュクリア');
  }

  /**
   * パーセンテージを計算
   * @param {Object} data データオブジェクト
   * @param {number} total 合計値
   * @returns {Object} パーセンテージ付きデータ
   */
  calculatePercentages(data, total) {
    const result = {};

    for (const [key, value] of Object.entries(data)) {
      result[key] = {
        count: value,
        percentage: total > 0 ? Math.round((value / total) * 100) : 0
      };
    }

    return result;
  }

  /**
   * 日別統計を生成
   * @param {Object} requestsByDay 日別リクエスト数
   * @returns {Array} 日別統計（過去30日）
   */
  generateDailyStats(requestsByDay) {
    const stats = [];
    const today = new Date();

    for (let i = 29; i >= 0; i--) {
      const date = new Date(today);
      date.setDate(date.getDate() - i);
      const dateStr = date.toISOString().split('T')[0];

      stats.push({
        date: dateStr,
        requests: requestsByDay[dateStr] || 0,
        dayOfWeek: date.getDay(),
        dayName: date.toLocaleDateString('ja-JP', { weekday: 'short' })
      });
    }

    return stats;
  }

  /**
   * 週別統計を生成
   * @param {Object} requestsByDay 日別リクエスト数
   * @returns {Array} 週別統計（過去8週間）
   */
  generateWeeklyStats(requestsByDay) {
    const stats = [];
    const today = new Date();

    for (let i = 7; i >= 0; i--) {
      const weekStart = new Date(today);
      weekStart.setDate(weekStart.getDate() - (weekStart.getDay() + i * 7));

      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);

      let weeklyRequests = 0;
      for (let j = 0; j < 7; j++) {
        const date = new Date(weekStart);
        date.setDate(date.getDate() + j);
        const dateStr = date.toISOString().split('T')[0];
        weeklyRequests += requestsByDay[dateStr] || 0;
      }

      stats.push({
        weekStart: weekStart.toISOString().split('T')[0],
        weekEnd: weekEnd.toISOString().split('T')[0],
        requests: weeklyRequests
      });
    }

    return stats;
  }

  /**
   * レスポンス時間をカテゴリ化
   * @param {number} responseTime レスポンス時間（ms）
   * @returns {string} カテゴリ
   */
  categorizeResponseTime(responseTime) {
    if (responseTime < 1000) return 'excellent';
    if (responseTime < 3000) return 'good';
    if (responseTime < 5000) return 'average';
    if (responseTime < 10000) return 'slow';
    return 'very_slow';
  }

  /**
   * 信頼性をカテゴリ化
   * @param {number} successRate 成功率（0-100）
   * @returns {string} カテゴリ
   */
  categorizeReliability(successRate) {
    if (successRate >= 99) return 'excellent';
    if (successRate >= 95) return 'good';
    if (successRate >= 90) return 'average';
    if (successRate >= 80) return 'poor';
    return 'very_poor';
  }

  /**
   * レポートをCSV形式に変換
   * @param {Object} report 統計レポート
   * @returns {string} CSV文字列
   */
  convertReportToCsv(report) {
    const csvRows = [];

    // サマリー
    csvRows.push('項目,値');
    csvRows.push(`総リクエスト数,${report.summary.totalRequests}`);
    csvRows.push(`成功リクエスト数,${report.summary.successfulRequests}`);
    csvRows.push(`失敗リクエスト数,${report.summary.failedRequests}`);
    csvRows.push(`成功率,${report.summary.successRate}%`);
    csvRows.push(`平均レスポンス時間,${report.summary.averageResponseTime}ms`);
    csvRows.push('');

    // リクエストタイプ別
    csvRows.push('リクエストタイプ,件数,割合');
    for (const [type, data] of Object.entries(report.requestsByType)) {
      csvRows.push(`${type},${data.count},${data.percentage}%`);
    }
    csvRows.push('');

    // 日別統計
    csvRows.push('日付,リクエスト数');
    report.dailyStats.forEach(stat => {
      csvRows.push(`${stat.date},${stat.requests}`);
    });

    return csvRows.join('\n');
  }
}

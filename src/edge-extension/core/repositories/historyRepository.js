/*
 * AI業務支援ツール Edge拡張機能 - 履歴リポジトリ
 * Copyright (c) 2024 AI Business Support Team
 */

import { StorageRepository } from '../../infrastructure/storage/storageRepository.js';
import { Logger } from '../../shared/utils/logger.js';

/**
 * 履歴データリポジトリ
 */
export class HistoryRepository {
  /**
   * コンストラクタ
   */
  constructor() {
    this.storageRepository = new StorageRepository();
    this.logger = new Logger('HistoryRepository');
    this.STORAGE_KEY = 'ai_assistant_history';
  }

  /**
   * 履歴を取得
   * @param {Object} options - 取得オプション
   * @returns {Promise<Object>} 履歴データ
   */
  async getHistory(options = {}) {
    try {
      this.logger.debug('履歴データ取得を開始します', options);

      const historyData = await this.storageRepository.get(this.STORAGE_KEY);

      if (!historyData || !historyData.items) {
        this.logger.debug('履歴データが存在しません。空の履歴を返します');
        return { items: [], lastUpdated: new Date().toISOString() };
      }

      // フィルタリング処理
      let filteredItems = historyData.items;

      // タイプフィルター
      if (options.type) {
        filteredItems = filteredItems.filter(item => item.type === options.type);
      }

      // 期間フィルター
      if (options.dateFrom) {
        const fromDate = new Date(options.dateFrom);
        filteredItems = filteredItems.filter(item => new Date(item.timestamp) >= fromDate);
      }

      if (options.dateTo) {
        const toDate = new Date(options.dateTo);
        filteredItems = filteredItems.filter(item => new Date(item.timestamp) <= toDate);
      }

      // 件数制限
      if (options.limit && options.limit > 0) {
        filteredItems = filteredItems
          .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
          .slice(0, options.limit);
      }

      this.logger.debug(`履歴データ取得が完了しました（${filteredItems.length}件）`);
      return { ...historyData, items: filteredItems };
    } catch (error) {
      this.logger.error('履歴データ取得中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目を追加
   * @param {Object} item - 履歴項目
   * @returns {Promise<void>}
   */
  async addHistoryItem(item) {
    try {
      this.logger.debug('履歴項目追加を開始します', { id: item.id });

      const historyData = await this.getHistory();

      // 新しい項目を先頭に追加
      historyData.items.unshift(item);
      historyData.lastUpdated = new Date().toISOString();

      await this.storageRepository.set(this.STORAGE_KEY, historyData);

      this.logger.debug('履歴項目追加が完了しました', { id: item.id });
    } catch (error) {
      this.logger.error('履歴項目追加中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目を取得
   * @param {string} id - 履歴項目ID
   * @returns {Promise<Object|null>} 履歴項目
   */
  async getHistoryItem(id) {
    try {
      this.logger.debug('履歴項目取得を開始します', { id });

      const historyData = await this.getHistory();
      const item = historyData.items.find(item => item.id === id);

      this.logger.debug('履歴項目取得が完了しました', { id, found: !!item });
      return item || null;
    } catch (error) {
      this.logger.error('履歴項目取得中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目を更新
   * @param {string} id - 履歴項目ID
   * @param {Object} updatedItem - 更新された履歴項目
   * @returns {Promise<void>}
   */
  async updateHistoryItem(id, updatedItem) {
    try {
      this.logger.debug('履歴項目更新を開始します', { id });

      const historyData = await this.getHistory();
      const itemIndex = historyData.items.findIndex(item => item.id === id);

      if (itemIndex === -1) {
        throw new Error(`履歴項目が見つかりません: ${id}`);
      }

      historyData.items[itemIndex] = updatedItem;
      historyData.lastUpdated = new Date().toISOString();

      await this.storageRepository.set(this.STORAGE_KEY, historyData);

      this.logger.debug('履歴項目更新が完了しました', { id });
    } catch (error) {
      this.logger.error('履歴項目更新中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目を削除
   * @param {string} id - 履歴項目ID
   * @returns {Promise<void>}
   */
  async deleteHistoryItem(id) {
    try {
      this.logger.debug('履歴項目削除を開始します', { id });

      const historyData = await this.getHistory();
      const originalLength = historyData.items.length;

      historyData.items = historyData.items.filter(item => item.id !== id);

      if (historyData.items.length === originalLength) {
        throw new Error(`履歴項目が見つかりません: ${id}`);
      }

      historyData.lastUpdated = new Date().toISOString();

      await this.storageRepository.set(this.STORAGE_KEY, historyData);

      this.logger.debug('履歴項目削除が完了しました', { id });
    } catch (error) {
      this.logger.error('履歴項目削除中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴を全削除
   * @returns {Promise<void>}
   */
  async clearHistory() {
    try {
      this.logger.debug('履歴全削除を開始します');

      const emptyHistory = {
        items: [],
        lastUpdated: new Date().toISOString()
      };

      await this.storageRepository.set(this.STORAGE_KEY, emptyHistory);

      this.logger.debug('履歴全削除が完了しました');
    } catch (error) {
      this.logger.error('履歴全削除中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴を検索
   * @param {string} query - 検索クエリ
   * @param {Object} options - 検索オプション
   * @returns {Promise<Array>} 検索結果
   */
  async searchHistory(query, options = {}) {
    try {
      this.logger.debug('履歴検索を開始します', { query, options });

      const historyData = await this.getHistory();
      const queryLower = query.toLowerCase();

      const results = historyData.items.filter(item => {
        const titleMatch = item.title.toLowerCase().includes(queryLower);
        const contentMatch = item.content.toLowerCase().includes(queryLower);
        const resultMatch = item.result.toLowerCase().includes(queryLower);

        return titleMatch || contentMatch || resultMatch;
      });

      // 関連度でソート（タイトル一致を優先）
      results.sort((a, b) => {
        const aScore = this.calculateRelevanceScore(a, queryLower);
        const bScore = this.calculateRelevanceScore(b, queryLower);
        return bScore - aScore;
      });

      // 件数制限
      const limitedResults = options.limit ? results.slice(0, options.limit) : results;

      this.logger.debug(`履歴検索が完了しました（${limitedResults.length}件）`);
      return limitedResults;
    } catch (error) {
      this.logger.error('履歴検索中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴統計を取得
   * @param {Object} options - 統計オプション
   * @returns {Promise<Object>} 履歴統計
   */
  async getHistoryStatistics(options = {}) {
    try {
      this.logger.debug('履歴統計取得を開始します', options);

      const historyData = await this.getHistory();
      const items = historyData.items;

      // 基本統計
      const statistics = {
        totalItems: items.length,
        typeDistribution: {},
        serviceDistribution: {},
        dailyActivity: {},
        averageContentLength: 0,
        averageResultLength: 0
      };

      if (items.length === 0) {
        return statistics;
      }

      // タイプ別分布
      items.forEach(item => {
        statistics.typeDistribution[item.type] =
          (statistics.typeDistribution[item.type] || 0) + 1;
      });

      // サービス別分布
      items.forEach(item => {
        statistics.serviceDistribution[item.service] =
          (statistics.serviceDistribution[item.service] || 0) + 1;
      });

      // 日別アクティビティ
      items.forEach(item => {
        const date = new Date(item.timestamp).toISOString().split('T')[0];
        statistics.dailyActivity[date] =
          (statistics.dailyActivity[date] || 0) + 1;
      });

      // 平均文字数
      const totalContentLength = items.reduce((sum, item) => sum + (item.content?.length || 0), 0);
      const totalResultLength = items.reduce((sum, item) => sum + (item.result?.length || 0), 0);

      statistics.averageContentLength = Math.round(totalContentLength / items.length);
      statistics.averageResultLength = Math.round(totalResultLength / items.length);

      this.logger.debug('履歴統計取得が完了しました');
      return statistics;
    } catch (error) {
      this.logger.error('履歴統計取得中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 関連度スコアを計算
   * @param {Object} item - 履歴項目
   * @param {string} query - 検索クエリ（小文字）
   * @returns {number} スコア
   * @private
   */
  calculateRelevanceScore(item, query) {
    let score = 0;

    // タイトル一致（高優先度）
    if (item.title.toLowerCase().includes(query)) {
      score += 100;
    }

    // コンテンツ一致（中優先度）
    if (item.content.toLowerCase().includes(query)) {
      score += 50;
    }

    // 結果一致（低優先度）
    if (item.result.toLowerCase().includes(query)) {
      score += 25;
    }

    // 最新の項目に追加スコア
    const daysSinceCreation = (Date.now() - new Date(item.timestamp).getTime()) / (1000 * 60 * 60 * 24);
    score += Math.max(0, 10 - daysSinceCreation);

    return score;
  }

  /**
   * 履歴の変更を監視
   * @param {Function} callback - 変更時のコールバック関数
   */
  onHistoryChanged(callback) {
    try {
      this.storageRepository.onChanged(this.STORAGE_KEY, callback);
    } catch (error) {
      this.logger.error('履歴変更監視の登録中にエラーが発生しました', error);
      throw error;
    }
  }
}

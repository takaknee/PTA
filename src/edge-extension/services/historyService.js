/*
 * AI業務支援ツール Edge拡張機能 - 履歴管理サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { HistoryItem } from '../lib/models.js';
import { STORAGE_KEYS, SYSTEM_CONFIG } from '../lib/constants.js';
import { Logger, formatDateTime, safeJsonParse, safeJsonStringify } from '../lib/utils.js';

/**
 * 履歴管理サービス
 */
export class HistoryService {
  constructor() {
    this.cache = null;
    this.listeners = new Set();
  }

  /**
   * 履歴を取得
   * @param {Object} options 取得オプション
   * @returns {Promise<HistoryItem[]>} 履歴項目配列
   */
  async getHistory(options = {}) {
    try {
      const {
        limit = SYSTEM_CONFIG.MAX_HISTORY_ITEMS,
        offset = 0,
        type = null,
        searchQuery = null,
        sortBy = 'timestamp',
        sortOrder = 'desc',
        useCache = true
      } = options;

      // キャッシュがある場合はキャッシュから取得
      if (useCache && this.cache) {
        Logger.debug('履歴をキャッシュから取得');
        return this.filterAndSortHistory(this.cache, { limit, offset, type, searchQuery, sortBy, sortOrder });
      }

      Logger.debug('履歴をストレージから取得');

      // Chrome storage から履歴を取得
      const result = await chrome.storage.local.get(STORAGE_KEYS.HISTORY);
      const storedHistory = result[STORAGE_KEYS.HISTORY] || [];

      // HistoryItemオブジェクトに変換
      const historyItems = storedHistory.map(item => new HistoryItem(item));

      // キャッシュに保存
      this.cache = historyItems;

      Logger.debug('履歴取得完了', { count: historyItems.length });

      // フィルタリングとソート
      return this.filterAndSortHistory(historyItems, { limit, offset, type, searchQuery, sortBy, sortOrder });

    } catch (error) {
      Logger.error('履歴取得エラー', error);
      return [];
    }
  }

  /**
   * 履歴にアイテムを追加
   * @param {HistoryItem|Object} item 履歴項目
   * @returns {Promise<boolean>} 追加成功可否
   */
  async addHistoryItem(item) {
    try {
      Logger.info('履歴追加開始', { type: item.type, title: item.title?.substring(0, 50) });

      // HistoryItemオブジェクトでない場合は変換
      const historyItem = item instanceof HistoryItem ? item : new HistoryItem(item);

      // 現在の履歴を取得
      const currentHistory = await this.getHistory({ useCache: false });

      // 新しいアイテムを先頭に追加
      const updatedHistory = [historyItem, ...currentHistory];

      // 最大件数を超えた場合は古いものを削除
      if (updatedHistory.length > SYSTEM_CONFIG.MAX_HISTORY_ITEMS) {
        updatedHistory.splice(SYSTEM_CONFIG.MAX_HISTORY_ITEMS);
        Logger.debug('古い履歴項目を削除', {
          削除件数: updatedHistory.length - SYSTEM_CONFIG.MAX_HISTORY_ITEMS
        });
      }

      // ストレージに保存
      await chrome.storage.local.set({
        [STORAGE_KEYS.HISTORY]: updatedHistory.map(item => item.toJSON())
      });

      // キャッシュを更新
      this.cache = updatedHistory;

      // リスナーに変更を通知
      this.notifyListeners('add', historyItem);

      Logger.info('履歴追加完了', { id: historyItem.id });
      return true;

    } catch (error) {
      Logger.error('履歴追加エラー', error);
      return false;
    }
  }

  /**
   * 履歴項目を更新
   * @param {string} id 項目ID
   * @param {Object} updates 更新データ
   * @returns {Promise<boolean>} 更新成功可否
   */
  async updateHistoryItem(id, updates) {
    try {
      Logger.info('履歴更新開始', { id });

      const currentHistory = await this.getHistory({ useCache: false });
      const itemIndex = currentHistory.findIndex(item => item.id === id);

      if (itemIndex === -1) {
        throw new Error('履歴項目が見つかりません');
      }

      // アイテムを更新
      const currentItem = currentHistory[itemIndex];
      const updatedItem = new HistoryItem({
        ...currentItem.toJSON(),
        ...updates,
        id: currentItem.id // IDは変更不可
      });

      currentHistory[itemIndex] = updatedItem;

      // ストレージに保存
      await chrome.storage.local.set({
        [STORAGE_KEYS.HISTORY]: currentHistory.map(item => item.toJSON())
      });

      // キャッシュを更新
      this.cache = currentHistory;

      // リスナーに変更を通知
      this.notifyListeners('update', updatedItem);

      Logger.info('履歴更新完了', { id });
      return true;

    } catch (error) {
      Logger.error('履歴更新エラー', error);
      return false;
    }
  }

  /**
   * 履歴項目を削除
   * @param {string} id 項目ID
   * @returns {Promise<boolean>} 削除成功可否
   */
  async deleteHistoryItem(id) {
    try {
      Logger.info('履歴削除開始', { id });

      const currentHistory = await this.getHistory({ useCache: false });
      const filteredHistory = currentHistory.filter(item => item.id !== id);

      if (filteredHistory.length === currentHistory.length) {
        throw new Error('削除対象の履歴項目が見つかりません');
      }

      // ストレージに保存
      await chrome.storage.local.set({
        [STORAGE_KEYS.HISTORY]: filteredHistory.map(item => item.toJSON())
      });

      // キャッシュを更新
      this.cache = filteredHistory;

      // リスナーに変更を通知
      this.notifyListeners('delete', { id });

      Logger.info('履歴削除完了', { id });
      return true;

    } catch (error) {
      Logger.error('履歴削除エラー', error);
      return false;
    }
  }

  /**
   * 履歴を全削除
   * @returns {Promise<boolean>} 削除成功可否
   */
  async clearHistory() {
    try {
      Logger.info('履歴全削除開始');

      // ストレージから削除
      await chrome.storage.local.remove(STORAGE_KEYS.HISTORY);

      // キャッシュをクリア
      this.cache = null;

      // リスナーに変更を通知
      this.notifyListeners('clear');

      Logger.info('履歴全削除完了');
      return true;

    } catch (error) {
      Logger.error('履歴全削除エラー', error);
      return false;
    }
  }

  /**
   * 履歴を検索
   * @param {string} query 検索クエリ
   * @param {Object} options 検索オプション
   * @returns {Promise<HistoryItem[]>} 検索結果
   */
  async searchHistory(query, options = {}) {
    try {
      const {
        type = null,
        limit = 50,
        sortBy = 'timestamp',
        sortOrder = 'desc'
      } = options;

      const allHistory = await this.getHistory({ useCache: true });

      // 検索クエリでフィルタリング
      const searchResults = allHistory.filter(item => {
        const searchText = item.getSearchableText();
        const queryLower = query.toLowerCase();

        // タイプフィルタ
        if (type && item.type !== type) {
          return false;
        }

        // テキスト検索
        return searchText.includes(queryLower);
      });

      // ソート
      const sortedResults = this.sortHistory(searchResults, sortBy, sortOrder);

      // 制限
      return sortedResults.slice(0, limit);

    } catch (error) {
      Logger.error('履歴検索エラー', error);
      return [];
    }
  }

  /**
   * 履歴をエクスポート
   * @param {Object} options エクスポートオプション
   * @returns {Promise<string>} エクスポートされたJSON文字列
   */
  async exportHistory(options = {}) {
    try {
      const {
        includeContent = true,
        includeResults = true,
        format = 'json'
      } = options;

      const history = await this.getHistory({ useCache: true });

      const exportData = {
        version: '1.0',
        timestamp: new Date().toISOString(),
        count: history.length,
        history: history.map(item => {
          const data = item.toJSON();

          // オプションに基づいてデータを調整
          if (!includeContent) {
            delete data.content;
          }
          if (!includeResults) {
            delete data.result;
          }

          return data;
        })
      };

      if (format === 'csv') {
        return this.convertHistoryToCsv(exportData.history);
      }

      return safeJsonStringify(exportData, '{}');

    } catch (error) {
      Logger.error('履歴エクスポートエラー', error);
      return format === 'csv' ? '' : '{}';
    }
  }

  /**
   * 履歴をインポート
   * @param {string} data インポートするデータ
   * @param {Object} options インポートオプション
   * @returns {Promise<boolean>} インポート成功可否
   */
  async importHistory(data, options = {}) {
    try {
      const {
        merge = true,
        overwrite = false
      } = options;

      Logger.info('履歴インポート開始');

      const importData = safeJsonParse(data);
      if (!importData || !importData.history || !Array.isArray(importData.history)) {
        throw new Error('無効なインポートデータ');
      }

      // インポートデータをHistoryItemオブジェクトに変換
      const importedItems = importData.history.map(item => new HistoryItem(item));

      let finalHistory;
      if (overwrite) {
        // 上書きモード
        finalHistory = importedItems;
      } else if (merge) {
        // マージモード
        const currentHistory = await this.getHistory({ useCache: false });
        const mergedHistory = [...importedItems, ...currentHistory];

        // 重複を除去（IDベース）
        const uniqueHistory = [];
        const seenIds = new Set();

        for (const item of mergedHistory) {
          if (!seenIds.has(item.id)) {
            uniqueHistory.push(item);
            seenIds.add(item.id);
          }
        }

        finalHistory = uniqueHistory;
      } else {
        // 既存履歴を保持
        finalHistory = await this.getHistory({ useCache: false });
      }

      // 最大件数制限
      if (finalHistory.length > SYSTEM_CONFIG.MAX_HISTORY_ITEMS) {
        finalHistory = finalHistory.slice(0, SYSTEM_CONFIG.MAX_HISTORY_ITEMS);
      }

      // ストレージに保存
      await chrome.storage.local.set({
        [STORAGE_KEYS.HISTORY]: finalHistory.map(item => item.toJSON())
      });

      // キャッシュを更新
      this.cache = finalHistory;

      // リスナーに変更を通知
      this.notifyListeners('import', { count: importedItems.length });

      Logger.info('履歴インポート完了', { count: importedItems.length });
      return true;

    } catch (error) {
      Logger.error('履歴インポートエラー', error);
      return false;
    }
  }

  /**
   * 履歴変更リスナーを追加
   * @param {Function} listener リスナー関数
   */
  addChangeListener(listener) {
    if (typeof listener === 'function') {
      this.listeners.add(listener);
      Logger.debug('履歴変更リスナー追加', { count: this.listeners.size });
    }
  }

  /**
   * 履歴変更リスナーを削除
   * @param {Function} listener リスナー関数
   */
  removeChangeListener(listener) {
    this.listeners.delete(listener);
    Logger.debug('履歴変更リスナー削除', { count: this.listeners.size });
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
        Logger.error('履歴リスナー通知エラー', error);
      }
    });
  }

  /**
   * キャッシュをクリア
   */
  clearCache() {
    this.cache = null;
    Logger.debug('履歴キャッシュクリア');
  }

  /**
   * 履歴統計を取得
   * @returns {Promise<Object>} 統計情報
   */
  async getStatistics() {
    try {
      const history = await this.getHistory({ useCache: true });

      const stats = {
        totalItems: history.length,
        itemsByType: {},
        itemsByDay: {},
        recentActivity: []
      };

      // タイプ別集計
      history.forEach(item => {
        stats.itemsByType[item.type] = (stats.itemsByType[item.type] || 0) + 1;
      });

      // 日別集計
      history.forEach(item => {
        const date = new Date(item.timestamp).toISOString().split('T')[0];
        stats.itemsByDay[date] = (stats.itemsByDay[date] || 0) + 1;
      });

      // 最近のアクティビティ（過去7日間）
      const sevenDaysAgo = Date.now() - (7 * 24 * 60 * 60 * 1000);
      stats.recentActivity = history
        .filter(item => item.timestamp > sevenDaysAgo)
        .slice(0, 10);

      return stats;

    } catch (error) {
      Logger.error('履歴統計取得エラー', error);
      return {
        totalItems: 0,
        itemsByType: {},
        itemsByDay: {},
        recentActivity: []
      };
    }
  }

  /**
   * 履歴のフィルタリングとソート
   * @param {HistoryItem[]} history 履歴配列
   * @param {Object} options オプション
   * @returns {HistoryItem[]} フィルタリング・ソート済み履歴
   */
  filterAndSortHistory(history, options) {
    let filteredHistory = [...history];

    // タイプフィルタ
    if (options.type) {
      filteredHistory = filteredHistory.filter(item => item.type === options.type);
    }

    // 検索クエリフィルタ
    if (options.searchQuery) {
      const query = options.searchQuery.toLowerCase();
      filteredHistory = filteredHistory.filter(item => {
        return item.getSearchableText().includes(query);
      });
    }

    // ソート
    filteredHistory = this.sortHistory(filteredHistory, options.sortBy, options.sortOrder);

    // ページネーション
    const start = options.offset || 0;
    const end = start + (options.limit || filteredHistory.length);

    return filteredHistory.slice(start, end);
  }

  /**
   * 履歴のソート
   * @param {HistoryItem[]} history 履歴配列
   * @param {string} sortBy ソートキー
   * @param {string} sortOrder ソート順
   * @returns {HistoryItem[]} ソート済み履歴
   */
  sortHistory(history, sortBy, sortOrder) {
    const sortedHistory = [...history];

    sortedHistory.sort((a, b) => {
      let aValue, bValue;

      switch (sortBy) {
        case 'timestamp':
          aValue = a.timestamp;
          bValue = b.timestamp;
          break;
        case 'title':
          aValue = a.title.toLowerCase();
          bValue = b.title.toLowerCase();
          break;
        case 'type':
          aValue = a.type;
          bValue = b.type;
          break;
        default:
          aValue = a.timestamp;
          bValue = b.timestamp;
      }

      if (aValue < bValue) return sortOrder === 'asc' ? -1 : 1;
      if (aValue > bValue) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });

    return sortedHistory;
  }

  /**
   * 履歴をCSV形式に変換
   * @param {Object[]} history 履歴データ
   * @returns {string} CSV文字列
   */
  convertHistoryToCsv(history) {
    if (!history || history.length === 0) {
      return '';
    }

    const headers = ['ID', 'タイプ', 'タイトル', '日時', 'タグ'];
    const csvRows = [headers.join(',')];

    history.forEach(item => {
      const row = [
        `"${item.id}"`,
        `"${item.type}"`,
        `"${item.title.replace(/"/g, '""')}"`,
        `"${formatDateTime(item.timestamp)}"`,
        `"${item.tags.join(', ')}"`
      ];
      csvRows.push(row.join(','));
    });

    return csvRows.join('\n');
  }
}

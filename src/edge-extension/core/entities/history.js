/*
 * AI業務支援ツール Edge拡張機能 - 履歴エンティティ
 * Copyright (c) 2024 AI Business Support Team
 */

import { HISTORY_TYPES } from '../../shared/constants/index.js';

/**
 * 履歴項目エンティティ
 */
export class HistoryItem {
  constructor(data = {}) {
    this.id = data.id || this.generateId();
    this.type = data.type || HISTORY_TYPES.PAGE_ANALYSIS;
    this.timestamp = data.timestamp || new Date().toISOString();
    this.title = data.title || '';
    this.content = data.content || '';
    this.result = data.result || '';
    this.service = data.service || 'unknown';
    this.metadata = data.metadata || {};
    this.status = data.status || 'completed';
    this.duration = data.duration || 0;
    this.tokenUsage = data.tokenUsage || { input: 0, output: 0, total: 0 };
  }

  /**
   * 一意IDの生成
   * @returns {string} 一意ID
   */
  generateId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * タイトルの自動生成
   * @returns {string} 生成されたタイトル
   */
  generateTitle() {
    if (this.title) return this.title;

    const typeMap = {
      [HISTORY_TYPES.EMAIL_ANALYSIS]: 'メール解析',
      [HISTORY_TYPES.PAGE_ANALYSIS]: 'ページ解析',
      [HISTORY_TYPES.SELECTION_ANALYSIS]: 'テキスト解析',
      [HISTORY_TYPES.EMAIL_COMPOSITION]: 'メール作成',
      [HISTORY_TYPES.EMAIL_REPLY]: 'メール返信',
      [HISTORY_TYPES.TRANSLATION]: '翻訳',
      [HISTORY_TYPES.URL_EXTRACTION]: 'URL抽出',
      [HISTORY_TYPES.CONTENT_SUMMARY]: 'コンテンツ要約'
    };

    const baseTitle = typeMap[this.type] || '処理';
    const date = new Date(this.timestamp);
    const timeStr = date.toLocaleTimeString('ja-JP', {
      hour: '2-digit',
      minute: '2-digit'
    });

    return `${baseTitle} ${timeStr}`;
  }

  /**
   * サマリーテキストの生成
   * @param {number} maxLength - 最大長
   * @returns {string} サマリー
   */
  getSummary(maxLength = 100) {
    let text = this.result || this.content;
    if (text.length <= maxLength) return text;

    return text.substring(0, maxLength - 3) + '...';
  }

  /**
   * 実行時間の取得（人間が読める形式）
   * @returns {string} 実行時間文字列
   */
  getFormattedDuration() {
    if (this.duration < 1000) {
      return `${this.duration}ms`;
    } else if (this.duration < 60000) {
      return `${(this.duration / 1000).toFixed(1)}秒`;
    } else {
      const minutes = Math.floor(this.duration / 60000);
      const seconds = Math.floor((this.duration % 60000) / 1000);
      return `${minutes}分${seconds}秒`;
    }
  }

  /**
   * 成功判定
   * @returns {boolean} 成功フラグ
   */
  isSuccess() {
    return this.status === 'completed' && this.result;
  }

  /**
   * エラー判定
   * @returns {boolean} エラーフラグ
   */
  isError() {
    return this.status === 'error';
  }

  /**
   * 検索用テキストの生成
   * @returns {string} 検索用テキスト
   */
  getSearchText() {
    return [
      this.title,
      this.content,
      this.result,
      this.metadata.url || '',
      this.metadata.emailSubject || ''
    ].join(' ').toLowerCase();
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      id: this.id,
      type: this.type,
      timestamp: this.timestamp,
      title: this.title,
      content: this.content,
      result: this.result,
      service: this.service,
      metadata: { ...this.metadata },
      status: this.status,
      duration: this.duration,
      tokenUsage: { ...this.tokenUsage }
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {HistoryItem} 履歴項目インスタンス
   */
  static fromJSON(json) {
    return new HistoryItem(json);
  }

  /**
   * 履歴項目の作成
   * @param {string} type - 項目タイプ
   * @param {string} content - コンテンツ
   * @param {Object} options - オプション
   * @returns {HistoryItem} 履歴項目インスタンス
   */
  static create(type, content, options = {}) {
    return new HistoryItem({
      type,
      content,
      title: options.title,
      service: options.service,
      metadata: options.metadata || {}
    });
  }
}

/**
 * 履歴コレクションエンティティ
 */
export class HistoryCollection {
  constructor(items = []) {
    this.items = items.map(item =>
      item instanceof HistoryItem ? item : HistoryItem.fromJSON(item)
    );
    this.maxItems = 100;
  }

  /**
   * 項目の追加
   * @param {HistoryItem} item - 履歴項目
   */
  add(item) {
    this.items.unshift(item);

    // 最大件数制限
    if (this.items.length > this.maxItems) {
      this.items = this.items.slice(0, this.maxItems);
    }
  }

  /**
   * 項目の削除
   * @param {string} id - 項目ID
   * @returns {boolean} 削除成功フラグ
   */
  remove(id) {
    const index = this.items.findIndex(item => item.id === id);
    if (index !== -1) {
      this.items.splice(index, 1);
      return true;
    }
    return false;
  }

  /**
   * 項目の取得
   * @param {string} id - 項目ID
   * @returns {HistoryItem|null} 履歴項目
   */
  get(id) {
    return this.items.find(item => item.id === id) || null;
  }

  /**
   * 全項目のクリア
   */
  clear() {
    this.items = [];
  }

  /**
   * 項目数の取得
   * @returns {number} 項目数
   */
  count() {
    return this.items.length;
  }

  /**
   * 空判定
   * @returns {boolean} 空フラグ
   */
  isEmpty() {
    return this.items.length === 0;
  }

  /**
   * タイプ別フィルタ
   * @param {string} type - 項目タイプ
   * @returns {HistoryItem[]} フィルタされた項目
   */
  filterByType(type) {
    return this.items.filter(item => item.type === type);
  }

  /**
   * 期間別フィルタ
   * @param {Date} startDate - 開始日
   * @param {Date} endDate - 終了日
   * @returns {HistoryItem[]} フィルタされた項目
   */
  filterByDateRange(startDate, endDate) {
    return this.items.filter(item => {
      const itemDate = new Date(item.timestamp);
      return itemDate >= startDate && itemDate <= endDate;
    });
  }

  /**
   * キーワード検索
   * @param {string} keyword - 検索キーワード
   * @returns {HistoryItem[]} 検索結果
   */
  search(keyword) {
    if (!keyword) return this.items;

    const lowerKeyword = keyword.toLowerCase();
    return this.items.filter(item =>
      item.getSearchText().includes(lowerKeyword)
    );
  }

  /**
   * ページネーション
   * @param {number} page - ページ番号（1から開始）
   * @param {number} pageSize - ページサイズ
   * @returns {Object} ページネーション結果
   */
  paginate(page = 1, pageSize = 20) {
    const totalItems = this.items.length;
    const totalPages = Math.ceil(totalItems / pageSize);
    const start = (page - 1) * pageSize;
    const end = start + pageSize;

    return {
      items: this.items.slice(start, end),
      currentPage: page,
      totalPages,
      totalItems,
      hasNext: page < totalPages,
      hasPrev: page > 1
    };
  }

  /**
   * 統計情報の取得
   * @returns {Object} 統計情報
   */
  getStatistics() {
    const stats = {
      total: this.items.length,
      byType: {},
      byStatus: {},
      totalDuration: 0,
      totalTokens: 0
    };

    this.items.forEach(item => {
      // タイプ別統計
      stats.byType[item.type] = (stats.byType[item.type] || 0) + 1;

      // ステータス別統計
      stats.byStatus[item.status] = (stats.byStatus[item.status] || 0) + 1;

      // 総実行時間
      stats.totalDuration += item.duration;

      // 総トークン数
      stats.totalTokens += item.tokenUsage.total;
    });

    return stats;
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      items: this.items.map(item => item.toJSON()),
      maxItems: this.maxItems
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {HistoryCollection} 履歴コレクション
   */
  static fromJSON(json) {
    const collection = new HistoryCollection(json.items || []);
    if (json.maxItems) {
      collection.maxItems = json.maxItems;
    }
    return collection;
  }
}

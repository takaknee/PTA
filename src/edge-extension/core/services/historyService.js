/*
 * AI業務支援ツール Edge拡張機能 - 履歴サービス
 * Copyright (c) 2024 AI Business Support Team
 */

import { HistoryRepository } from '../repositories/historyRepository.js';
import { History, HistoryItem } from '../entities/history.js';
import { Logger } from '../../shared/utils/logger.js';
import { Validator } from '../../shared/utils/validator.js';
import { APP_CONSTANTS } from '../../shared/constants/app.js';

/**
 * 履歴管理サービス
 */
export class HistoryService {
  /**
   * コンストラクタ
   */
  constructor() {
    this.repository = new HistoryRepository();
    this.logger = new Logger('HistoryService');
    this.validator = new Validator();
  }

  /**
   * 履歴を取得
   * @param {Object} options - 取得オプション
   * @param {number} [options.limit] - 取得件数制限
   * @param {string} [options.type] - 履歴タイプフィルター
   * @param {string} [options.dateFrom] - 期間フィルター（開始）
   * @param {string} [options.dateTo] - 期間フィルター（終了）
   * @returns {Promise<History>} 履歴オブジェクト
   */
  async getHistory(options = {}) {
    try {
      this.logger.info('履歴取得処理を開始します', options);

      const historyData = await this.repository.getHistory(options);
      const history = History.fromJSON(historyData);

      this.logger.info(`履歴取得処理が完了しました（${history.items.length}件）`);
      return history;
    } catch (error) {
      this.logger.error('履歴取得処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目を追加
   * @param {Object} itemData - 履歴項目データ
   * @returns {Promise<HistoryItem>} 追加された履歴項目
   */
  async addHistoryItem(itemData) {
    try {
      this.logger.info('履歴項目追加処理を開始します');

      // 履歴項目の妥当性検証
      const validationResult = this.validateHistoryItem(itemData);
      if (!validationResult.isValid) {
        throw new Error(`履歴項目の妥当性検証に失敗しました: ${validationResult.errors.join(', ')}`);
      }

      // 履歴項目を作成
      const historyItem = HistoryItem.create(itemData);

      // 履歴項目を保存
      await this.repository.addHistoryItem(historyItem.toJSON());

      // 履歴のサイズ制限をチェック
      await this.enforceHistoryLimit();

      this.logger.info('履歴項目追加処理が完了しました', { id: historyItem.id });
      return historyItem;
    } catch (error) {
      this.logger.error('履歴項目追加処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目を更新
   * @param {string} id - 履歴項目ID
   * @param {Object} updateData - 更新データ
   * @returns {Promise<HistoryItem>} 更新された履歴項目
   */
  async updateHistoryItem(id, updateData) {
    try {
      this.logger.info('履歴項目更新処理を開始します', { id });

      if (!this.validator.isString(id) || id.trim().length === 0) {
        throw new Error('無効な履歴項目IDが指定されました');
      }

      // 既存の履歴項目を取得
      const existingItem = await this.repository.getHistoryItem(id);
      if (!existingItem) {
        throw new Error('指定された履歴項目が見つかりません');
      }

      // 履歴項目を更新
      const historyItem = HistoryItem.fromJSON(existingItem);
      const updatedItem = historyItem.update(updateData);

      // 更新を保存
      await this.repository.updateHistoryItem(id, updatedItem.toJSON());

      this.logger.info('履歴項目更新処理が完了しました', { id });
      return updatedItem;
    } catch (error) {
      this.logger.error('履歴項目更新処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目を削除
   * @param {string} id - 履歴項目ID
   * @returns {Promise<boolean>} 削除成功フラグ
   */
  async deleteHistoryItem(id) {
    try {
      this.logger.info('履歴項目削除処理を開始します', { id });

      if (!this.validator.isString(id) || id.trim().length === 0) {
        throw new Error('無効な履歴項目IDが指定されました');
      }

      await this.repository.deleteHistoryItem(id);

      this.logger.info('履歴項目削除処理が完了しました', { id });
      return true;
    } catch (error) {
      this.logger.error('履歴項目削除処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴を全削除
   * @returns {Promise<boolean>} 削除成功フラグ
   */
  async clearHistory() {
    try {
      this.logger.info('履歴全削除処理を開始します');

      await this.repository.clearHistory();

      this.logger.info('履歴全削除処理が完了しました');
      return true;
    } catch (error) {
      this.logger.error('履歴全削除処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴を検索
   * @param {string} query - 検索クエリ
   * @param {Object} options - 検索オプション
   * @returns {Promise<HistoryItem[]>} 検索結果
   */
  async searchHistory(query, options = {}) {
    try {
      this.logger.info('履歴検索処理を開始します', { query, options });

      if (!this.validator.isString(query) || query.trim().length === 0) {
        throw new Error('検索クエリが無効です');
      }

      const results = await this.repository.searchHistory(query, options);
      const historyItems = results.map(item => HistoryItem.fromJSON(item));

      this.logger.info(`履歴検索処理が完了しました（${historyItems.length}件）`);
      return historyItems;
    } catch (error) {
      this.logger.error('履歴検索処理中にエラーが発生しました', error);
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
      this.logger.info('履歴統計取得処理を開始します', options);

      const statistics = await this.repository.getHistoryStatistics(options);

      this.logger.info('履歴統計取得処理が完了しました');
      return statistics;
    } catch (error) {
      this.logger.error('履歴統計取得処理中にエラーが発生しました', error);
      throw error;
    }
  }

  /**
   * 履歴項目の妥当性を検証
   * @param {Object} itemData - 履歴項目データ
   * @returns {Object} 検証結果
   */
  validateHistoryItem(itemData) {
    const errors = [];

    try {
      // 必須フィールドの検証
      if (!this.validator.isString(itemData.type) || itemData.type.trim().length === 0) {
        errors.push('処理タイプが無効です');
      }

      if (!this.validator.isString(itemData.title) || itemData.title.trim().length === 0) {
        errors.push('タイトルが無効です');
      }

      if (!this.validator.isString(itemData.content)) {
        errors.push('コンテンツが無効です');
      }

      if (!this.validator.isString(itemData.result)) {
        errors.push('処理結果が無効です');
      }

      if (!this.validator.isString(itemData.service) || itemData.service.trim().length === 0) {
        errors.push('サービス名が無効です');
      }

      // データサイズの検証
      if (itemData.content && itemData.content.length > APP_CONSTANTS.MAX_HISTORY_CONTENT_LENGTH) {
        errors.push('コンテンツのサイズが上限を超えています');
      }

      if (itemData.result && itemData.result.length > APP_CONSTANTS.MAX_HISTORY_RESULT_LENGTH) {
        errors.push('処理結果のサイズが上限を超えています');
      }

      // メタデータの検証
      if (itemData.metadata && !this.validator.isObject(itemData.metadata)) {
        errors.push('メタデータの形式が無効です');
      }

      return {
        isValid: errors.length === 0,
        errors: errors
      };
    } catch (error) {
      this.logger.error('履歴項目妥当性検証中にエラーが発生しました', error);
      return {
        isValid: false,
        errors: ['履歴項目の妥当性検証中にエラーが発生しました']
      };
    }
  }

  /**
   * 履歴のサイズ制限を強制
   * @private
   */
  async enforceHistoryLimit() {
    try {
      const history = await this.getHistory();

      if (history.items.length > APP_CONSTANTS.MAX_HISTORY_ITEMS) {
        const itemsToDelete = history.items.length - APP_CONSTANTS.MAX_HISTORY_ITEMS;
        const oldestItems = history.items
          .sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp))
          .slice(0, itemsToDelete);

        for (const item of oldestItems) {
          await this.repository.deleteHistoryItem(item.id);
        }

        this.logger.info(`履歴サイズ制限により${itemsToDelete}件の古い項目を削除しました`);
      }
    } catch (error) {
      this.logger.error('履歴サイズ制限強制中にエラーが発生しました', error);
      // エラーが発生しても処理を続行
    }
  }

  /**
   * 履歴の変更を監視
   * @param {Function} callback - 変更時のコールバック関数
   */
  onHistoryChanged(callback) {
    try {
      this.repository.onHistoryChanged(callback);
    } catch (error) {
      this.logger.error('履歴変更監視の登録中にエラーが発生しました', error);
      throw error;
    }
  }
}

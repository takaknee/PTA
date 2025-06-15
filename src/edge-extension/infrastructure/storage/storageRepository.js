/*
 * AI業務支援ツール Edge拡張機能 - ストレージリポジトリ
 * Copyright (c) 2024 AI Business Support Team
 */

import { STORAGE_KEYS, ERROR_CODES } from '../../shared/constants/index.js';
import { createLogger } from '../../shared/utils/index.js';

const logger = createLogger('StorageRepository');

/**
 * ストレージエラークラス
 */
export class StorageError extends Error {
  constructor(code, message, cause = null) {
    super(message);
    this.name = 'StorageError';
    this.code = code;
    this.cause = cause;
  }
}

/**
 * ストレージリポジトリベースクラス
 */
export class BaseStorageRepository {
  constructor(storageArea = chrome.storage.local) {
    this.storage = storageArea;
  }

  /**
   * データの取得
   * @param {string} key - ストレージキー
   * @param {*} defaultValue - デフォルト値
   * @returns {Promise<*>} 取得したデータ
   */
  async get(key, defaultValue = null) {
    try {
      logger.debug('データ取得開始', { key });

      const result = await this.storage.get(key);
      const value = result[key];

      if (value === undefined) {
        logger.debug('データが見つからない、デフォルト値を返却', { key, defaultValue });
        return defaultValue;
      }

      logger.debug('データ取得完了', { key, hasValue: value !== null });
      return value;

    } catch (error) {
      logger.error('データ取得エラー', { key, error: error.message });
      throw new StorageError(
        ERROR_CODES.STORAGE_ACCESS_DENIED,
        `データの取得に失敗しました: ${key}`,
        error
      );
    }
  }

  /**
   * データの保存
   * @param {string} key - ストレージキー
   * @param {*} value - 保存するデータ
   * @returns {Promise<void>}
   */
  async set(key, value) {
    try {
      logger.debug('データ保存開始', { key });

      await this.storage.set({ [key]: value });

      logger.debug('データ保存完了', { key });

    } catch (error) {
      logger.error('データ保存エラー', { key, error: error.message });

      if (error.message.includes('QUOTA_EXCEEDED')) {
        throw new StorageError(
          ERROR_CODES.STORAGE_QUOTA_EXCEEDED,
          'ストレージ容量が不足しています',
          error
        );
      }

      throw new StorageError(
        ERROR_CODES.STORAGE_ACCESS_DENIED,
        `データの保存に失敗しました: ${key}`,
        error
      );
    }
  }

  /**
   * データの削除
   * @param {string} key - ストレージキー
   * @returns {Promise<void>}
   */
  async remove(key) {
    try {
      logger.debug('データ削除開始', { key });

      await this.storage.remove(key);

      logger.debug('データ削除完了', { key });

    } catch (error) {
      logger.error('データ削除エラー', { key, error: error.message });
      throw new StorageError(
        ERROR_CODES.STORAGE_ACCESS_DENIED,
        `データの削除に失敗しました: ${key}`,
        error
      );
    }
  }

  /**
   * 複数データの取得
   * @param {string[]} keys - ストレージキー配列
   * @returns {Promise<Object>} 取得したデータオブジェクト
   */
  async getMultiple(keys) {
    try {
      logger.debug('複数データ取得開始', { keys });

      const result = await this.storage.get(keys);

      logger.debug('複数データ取得完了', { keys, resultKeys: Object.keys(result) });
      return result;

    } catch (error) {
      logger.error('複数データ取得エラー', { keys, error: error.message });
      throw new StorageError(
        ERROR_CODES.STORAGE_ACCESS_DENIED,
        'データの取得に失敗しました',
        error
      );
    }
  }

  /**
   * 複数データの保存
   * @param {Object} data - 保存するデータオブジェクト
   * @returns {Promise<void>}
   */
  async setMultiple(data) {
    try {
      logger.debug('複数データ保存開始', { keys: Object.keys(data) });

      await this.storage.set(data);

      logger.debug('複数データ保存完了', { keys: Object.keys(data) });

    } catch (error) {
      logger.error('複数データ保存エラー', { keys: Object.keys(data), error: error.message });

      if (error.message.includes('QUOTA_EXCEEDED')) {
        throw new StorageError(
          ERROR_CODES.STORAGE_QUOTA_EXCEEDED,
          'ストレージ容量が不足しています',
          error
        );
      }

      throw new StorageError(
        ERROR_CODES.STORAGE_ACCESS_DENIED,
        'データの保存に失敗しました',
        error
      );
    }
  }

  /**
   * 複数データの削除
   * @param {string[]} keys - ストレージキー配列
   * @returns {Promise<void>}
   */
  async removeMultiple(keys) {
    try {
      logger.debug('複数データ削除開始', { keys });

      await this.storage.remove(keys);

      logger.debug('複数データ削除完了', { keys });

    } catch (error) {
      logger.error('複数データ削除エラー', { keys, error: error.message });
      throw new StorageError(
        ERROR_CODES.STORAGE_ACCESS_DENIED,
        'データの削除に失敗しました',
        error
      );
    }
  }

  /**
   * ストレージの使用量取得
   * @returns {Promise<Object>} 使用量情報
   */
  async getStorageUsage() {
    try {
      if (this.storage.getBytesInUse) {
        const bytesInUse = await this.storage.getBytesInUse();
        return { bytesInUse };
      }

      return { bytesInUse: null };

    } catch (error) {
      logger.warn('ストレージ使用量取得エラー', { error: error.message });
      return { bytesInUse: null };
    }
  }

  /**
   * ストレージのクリア
   * @returns {Promise<void>}
   */
  async clear() {
    try {
      logger.debug('ストレージクリア開始');

      await this.storage.clear();

      logger.debug('ストレージクリア完了');

    } catch (error) {
      logger.error('ストレージクリアエラー', { error: error.message });
      throw new StorageError(
        ERROR_CODES.STORAGE_ACCESS_DENIED,
        'ストレージのクリアに失敗しました',
        error
      );
    }
  }

  /**
   * データの存在確認
   * @param {string} key - ストレージキー
   * @returns {Promise<boolean>} 存在フラグ
   */
  async exists(key) {
    try {
      const result = await this.storage.get(key);
      return key in result;
    } catch (error) {
      logger.warn('データ存在確認エラー', { key, error: error.message });
      return false;
    }
  }

  /**
   * JSONデータの安全な取得
   * @param {string} key - ストレージキー
   * @param {*} defaultValue - デフォルト値
   * @returns {Promise<*>} パースされたデータ
   */
  async getJSON(key, defaultValue = null) {
    try {
      const value = await this.get(key, defaultValue);

      if (value === defaultValue) {
        return defaultValue;
      }

      // すでにオブジェクトの場合はそのまま返す
      if (typeof value === 'object' && value !== null) {
        return value;
      }

      // 文字列の場合はパースを試行
      if (typeof value === 'string') {
        try {
          return JSON.parse(value);
        } catch (parseError) {
          logger.warn('JSON パースエラー、生の値を返却', { key, parseError: parseError.message });
          return value;
        }
      }

      return value;

    } catch (error) {
      logger.error('JSON データ取得エラー', { key, error: error.message });
      throw error;
    }
  }

  /**
   * JSONデータの安全な保存
   * @param {string} key - ストレージキー
   * @param {*} value - 保存するデータ
   * @returns {Promise<void>}
   */
  async setJSON(key, value) {
    try {
      // オブジェクトの場合はそのまま保存（Chrome storage は JSON を自動処理）
      await this.set(key, value);

    } catch (error) {
      logger.error('JSON データ保存エラー', { key, error: error.message });
      throw error;
    }
  }
}

/**
 * ローカルストレージリポジトリ
 */
export class LocalStorageRepository extends BaseStorageRepository {
  constructor() {
    super(chrome.storage.local);
  }
}

/**
 * 同期ストレージリポジトリ
 */
export class SyncStorageRepository extends BaseStorageRepository {
  constructor() {
    super(chrome.storage.sync);
  }

  /**
   * 同期ストレージ特有の制限チェック
   * @param {*} value - 保存するデータ
   * @returns {boolean} 制限内フラグ
   */
  checkSyncLimits(value) {
    const serialized = JSON.stringify(value);
    const size = new Blob([serialized]).size;

    // Chrome sync storage limits
    const MAX_ITEM_SIZE = 8192; // 8KB per item
    const MAX_TOTAL_SIZE = 102400; // 100KB total

    return size <= MAX_ITEM_SIZE;
  }

  /**
   * 同期対応データ保存
   * @param {string} key - ストレージキー
   * @param {*} value - 保存するデータ
   * @returns {Promise<void>}
   */
  async set(key, value) {
    if (!this.checkSyncLimits(value)) {
      throw new StorageError(
        ERROR_CODES.STORAGE_QUOTA_EXCEEDED,
        'データサイズが同期ストレージの制限を超えています',
      );
    }

    return super.set(key, value);
  }
}

/**
 * ストレージリポジトリファクトリ
 */
export class StorageRepositoryFactory {
  /**
   * ローカルストレージリポジトリの取得
   * @returns {LocalStorageRepository} ローカルストレージリポジトリ
   */
  static createLocal() {
    return new LocalStorageRepository();
  }

  /**
   * 同期ストレージリポジトリの取得
   * @returns {SyncStorageRepository} 同期ストレージリポジトリ
   */
  static createSync() {
    return new SyncStorageRepository();
  }

  /**
   * デフォルトストレージリポジトリの取得
   * @returns {LocalStorageRepository} デフォルトリポジトリ
   */
  static createDefault() {
    return this.createLocal();
  }
}

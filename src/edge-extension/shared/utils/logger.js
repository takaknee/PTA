/*
 * AI業務支援ツール Edge拡張機能 - ログユーティリティ
 * Copyright (c) 2024 AI Business Support Team
 */

import { DEBUG_CONFIG } from '../constants/index.js';

/**
 * ログレベル列挙
 */
export const LogLevel = Object.freeze({
  ERROR: 0,
  WARN: 1,
  INFO: 2,
  DEBUG: 3
});

/**
 * ログエントリ
 */
export class LogEntry {
  constructor(level, message, data = null, source = null) {
    this.timestamp = new Date().toISOString();
    this.level = level;
    this.message = message;
    this.data = data;
    this.source = source;
    this.id = this.generateId();
  }

  generateId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  toString() {
    const levelName = Object.keys(LogLevel)[this.level];
    const sourceStr = this.source ? ` [${this.source}]` : '';
    const dataStr = this.data ? ` ${JSON.stringify(this.data)}` : '';
    return `${this.timestamp} ${levelName}${sourceStr}: ${this.message}${dataStr}`;
  }
}

/**
 * ログユーティリティクラス
 */
export class Logger {
  constructor(source = null, level = LogLevel.INFO) {
    this.source = source;
    this.level = level;
    this.entries = [];
    this.maxEntries = DEBUG_CONFIG.MAX_LOG_ENTRIES;
    this.enabled = DEBUG_CONFIG.ENABLED;
  }

  /**
   * ログレベルの設定
   * @param {number} level - ログレベル
   */
  setLevel(level) {
    this.level = level;
  }

  /**
   * ログの有効/無効切り替え
   * @param {boolean} enabled - 有効フラグ
   */
  setEnabled(enabled) {
    this.enabled = enabled;
  }

  /**
   * ログエントリの追加
   * @param {number} level - ログレベル
   * @param {string} message - メッセージ
   * @param {*} data - 追加データ
   */
  log(level, message, data = null) {
    if (!this.enabled || level > this.level) {
      return;
    }

    const entry = new LogEntry(level, message, data, this.source);
    this.entries.push(entry);

    // コンソール出力
    this.outputToConsole(entry);

    // エントリ数制限
    if (this.entries.length > this.maxEntries) {
      this.entries = this.entries.slice(-this.maxEntries);
    }
  }

  /**
   * エラーログ
   * @param {string} message - メッセージ
   * @param {*} data - 追加データ
   */
  error(message, data = null) {
    this.log(LogLevel.ERROR, message, data);
  }

  /**
   * 警告ログ
   * @param {string} message - メッセージ
   * @param {*} data - 追加データ
   */
  warn(message, data = null) {
    this.log(LogLevel.WARN, message, data);
  }

  /**
   * 情報ログ
   * @param {string} message - メッセージ
   * @param {*} data - 追加データ
   */
  info(message, data = null) {
    this.log(LogLevel.INFO, message, data);
  }

  /**
   * デバッグログ
   * @param {string} message - メッセージ
   * @param {*} data - 追加データ
   */
  debug(message, data = null) {
    this.log(LogLevel.DEBUG, message, data);
  }

  /**
   * コンソールへの出力
   * @param {LogEntry} entry - ログエントリ
   */
  outputToConsole(entry) {
    const logMessage = entry.toString();

    switch (entry.level) {
      case LogLevel.ERROR:
        console.error(logMessage);
        break;
      case LogLevel.WARN:
        console.warn(logMessage);
        break;
      case LogLevel.INFO:
        console.info(logMessage);
        break;
      case LogLevel.DEBUG:
        console.debug(logMessage);
        break;
      default:
        console.log(logMessage);
    }
  }

  /**
   * パフォーマンス測定の開始
   * @param {string} label - ラベル
   */
  timeStart(label) {
    if (this.enabled && this.level >= LogLevel.DEBUG) {
      console.time(label);
    }
  }

  /**
   * パフォーマンス測定の終了
   * @param {string} label - ラベル
   */
  timeEnd(label) {
    if (this.enabled && this.level >= LogLevel.DEBUG) {
      console.timeEnd(label);
    }
  }

  /**
   * 条件付きログ
   * @param {boolean} condition - 条件
   * @param {number} level - ログレベル
   * @param {string} message - メッセージ
   * @param {*} data - 追加データ
   */
  logIf(condition, level, message, data = null) {
    if (condition) {
      this.log(level, message, data);
    }
  }

  /**
   * ログエントリの取得
   * @param {number} level - フィルタするログレベル（省略可）
   * @param {number} limit - 取得件数制限（省略可）
   * @returns {LogEntry[]} ログエントリ配列
   */
  getEntries(level = null, limit = null) {
    let entries = this.entries;

    if (level !== null) {
      entries = entries.filter(entry => entry.level === level);
    }

    if (limit !== null) {
      entries = entries.slice(-limit);
    }

    return entries;
  }

  /**
   * ログエントリのクリア
   */
  clear() {
    this.entries = [];
  }

  /**
   * ログのエクスポート
   * @returns {string} ログの文字列表現
   */
  export() {
    return this.entries.map(entry => entry.toString()).join('\n');
  }

  /**
   * ログの統計情報取得
   * @returns {Object} 統計情報
   */
  getStatistics() {
    const stats = {
      total: this.entries.length,
      byLevel: {}
    };

    Object.values(LogLevel).forEach(level => {
      stats.byLevel[level] = this.entries.filter(entry => entry.level === level).length;
    });

    return stats;
  }
}

/**
 * グローバルロガーインスタンス
 */
export const globalLogger = new Logger('Global');

/**
 * ソース別ロガーファクトリ
 * @param {string} source - ソース名
 * @returns {Logger} ロガーインスタンス
 */
export function createLogger(source) {
  return new Logger(source);
}

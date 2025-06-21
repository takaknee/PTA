/*
 * PTA AI業務支援ツール - ロガーモジュール
 * Copyright (c) 2024 PTA Development Team
 */

import { ENV_CONSTANTS } from './constants.js';

// 統一セキュリティサニタイザーのインポート
let sanitizer = null;
try {
    // 動的インポートでサニタイザーを取得（循環依存回避）
    import('./unified-security-sanitizer.js').then(module => {
        sanitizer = module.getSecuritySanitizer();
    });
} catch (error) {
    console.warn('統一セキュリティサニタイザーの読み込みに失敗しました:', error);
}

/**
 * ログメッセージの安全な処理
 * HTML/特殊文字を含む可能性のあるデータをサニタイズ
 */
function sanitizeLogData(data) {
    if (!sanitizer || !data) return data;
    
    try {
        if (typeof data === 'string') {
            return sanitizer.escapeUserInput(data);
        } else if (typeof data === 'object') {
            const sanitized = {};
            for (const [key, value] of Object.entries(data)) {
                if (typeof value === 'string') {
                    sanitized[key] = sanitizer.escapeUserInput(value);
                } else {
                    sanitized[key] = value;
                }
            }
            return sanitized;
        }
        return data;
    } catch (error) {
        console.warn('ログデータのサニタイゼーションに失敗:', error);
        return data;
    }
}

/**
 * ログレベル定義
 */
const LOG_LEVELS = {
    DEBUG: 0,
    INFO: 1,
    WARN: 2,
    ERROR: 3,
};

/**
 * 統一ロガークラス
 * プロジェクト全体で一貫したログ出力を提供
 */
class PTALogger {
    constructor(context = 'PTA') {
        this.context = context;
        this.currentLevel = this._getLogLevel();
        this.logs = [];
        this.maxLogs = 1000; // メモリ使用量制限
    }

    /**
     * 現在の環境に応じたログレベルを取得
     */
    _getLogLevel() {
        const isDevelopment = process?.env?.NODE_ENV === 'development' ||
            chrome?.runtime?.getManifest()?.version?.includes('dev');

        const envConfig = isDevelopment ? ENV_CONSTANTS.DEVELOPMENT : ENV_CONSTANTS.PRODUCTION;

        switch (envConfig.LOG_LEVEL) {
            case 'debug': return LOG_LEVELS.DEBUG;
            case 'info': return LOG_LEVELS.INFO;
            case 'warn': return LOG_LEVELS.WARN;
            case 'error': return LOG_LEVELS.ERROR;
            default: return LOG_LEVELS.INFO;
        }
    }

    /**
     * ログエントリを作成
     */
    _createLogEntry(level, message, data = null) {
        const timestamp = new Date().toISOString();
        
        // セキュリティ: ログデータのサニタイゼーション
        const sanitizedMessage = sanitizeLogData(message);
        const sanitizedData = sanitizeLogData(data);
        
        const entry = {
            timestamp,
            level,
            context: this.context,
            message: sanitizedMessage,
            data: sanitizedData,
            stack: level === 'ERROR' ? new Error().stack : null,
        };

        // メモリ管理：古いログを削除
        this.logs.push(entry);
        if (this.logs.length > this.maxLogs) {
            this.logs.shift();
        }

        return entry;
    }

    /**
     * ログを出力（レベルチェック付き）
     */
    _log(level, levelValue, message, data = null) {
        if (levelValue < this.currentLevel) return;

        const entry = this._createLogEntry(level, message, data);
        const prefix = `[${entry.timestamp}] [${this.context}] [${level}]`;

        // コンソール出力
        switch (level) {
            case 'DEBUG':
                console.debug(prefix, message, data || '');
                break;
            case 'INFO':
                console.info(prefix, message, data || '');
                break;
            case 'WARN':
                console.warn(prefix, message, data || '');
                break;
            case 'ERROR':
                console.error(prefix, message, data || '', entry.stack || '');
                break;
        }

        // 重要なエラーの場合は通知も送信
        if (level === 'ERROR' && this._shouldNotifyError(message)) {
            this._sendErrorNotification(entry);
        }
    }

    /**
     * エラー通知が必要かどうかを判定
     */
    _shouldNotifyError(message) {
        const criticalErrors = [
            'API接続エラー',
            'セキュリティ違反',
            '認証エラー',
            'データ破損',
        ];

        return criticalErrors.some(error => message.includes(error));
    }

    /**
     * エラー通知を送信
     */
    _sendErrorNotification(entry) {
        try {
            // バックグラウンドスクリプトに通知
            if (chrome?.runtime?.sendMessage) {
                chrome.runtime.sendMessage({
                    action: 'logCriticalError',
                    data: entry
                });
            }
        } catch (error) {
            console.error('エラー通知の送信に失敗:', error);
        }
    }

    /**
     * デバッグログ
     */
    debug(message, data = null) {
        this._log('DEBUG', LOG_LEVELS.DEBUG, message, data);
    }

    /**
     * 情報ログ
     */
    info(message, data = null) {
        this._log('INFO', LOG_LEVELS.INFO, message, data);
    }

    /**
     * 警告ログ
     */
    warn(message, data = null) {
        this._log('WARN', LOG_LEVELS.WARN, message, data);
    }

    /**
     * エラーログ
     */
    error(message, data = null) {
        this._log('ERROR', LOG_LEVELS.ERROR, message, data);
    }

    /**
     * API呼び出しログ（専用メソッド）
     */
    apiCall(method, url, duration, success, response = null) {
        const message = `API ${method} ${url} - ${duration}ms - ${success ? '成功' : '失敗'}`;
        if (success) {
            this.info(message, { method, url, duration, response });
        } else {
            this.error(message, { method, url, duration, error: response });
        }
    }

    /**
     * パフォーマンス測定開始
     */
    startPerformance(label) {
        const startTime = performance.now();
        return {
            end: () => {
                const duration = performance.now() - startTime;
                this.debug(`パフォーマンス測定 [${label}]: ${duration.toFixed(2)}ms`);
                return duration;
            }
        };
    }

    /**
     * ユーザーアクションログ
     */
    userAction(action, details = null) {
        this.info(`ユーザーアクション: ${action}`, details);
    }

    /**
     * セキュリティイベントログ
     */
    security(message, details = null) {
        this.warn(`セキュリティ: ${message}`, details);
    }

    /**
     * ログ履歴を取得
     */
    getHistory(level = null, limit = 100) {
        let filtered = this.logs;

        if (level) {
            filtered = this.logs.filter(log => log.level === level.toUpperCase());
        }

        return filtered.slice(-limit);
    }

    /**
     * ログをクリア
     */
    clear() {
        this.logs = [];
        this.info('ログ履歴をクリアしました');
    }

    /**
     * ログを文字列として出力
     */
    export(format = 'json') {
        switch (format) {
            case 'json':
                return JSON.stringify(this.logs, null, 2);
            case 'csv':
                const headers = 'timestamp,level,context,message,data\n';
                const rows = this.logs.map(log =>
                    `${log.timestamp},${log.level},${log.context},"${log.message}","${log.data || ''}"`
                ).join('\n');
                return headers + rows;
            case 'text':
                return this.logs.map(log =>
                    `[${log.timestamp}] [${log.context}] [${log.level}] ${log.message} ${log.data ? JSON.stringify(log.data) : ''}`
                ).join('\n');
            default:
                return JSON.stringify(this.logs);
        }
    }
}

/**
 * グローバルロガーインスタンス
 */
export const logger = new PTALogger('PTA-Extension');

/**
 * コンテキスト固有のロガーを作成
 */
export function createLogger(context) {
    return new PTALogger(context);
}

/**
 * 便利な関数群
 */
export const logInfo = (message, data) => logger.info(message, data);
export const logError = (message, data) => logger.error(message, data);
export const logWarn = (message, data) => logger.warn(message, data);
export const logDebug = (message, data) => logger.debug(message, data);

// 未処理のエラーをキャッチ
if (typeof window !== 'undefined') {
    window.addEventListener('error', (event) => {
        logger.error('未処理のエラー', {
            message: event.message,
            filename: event.filename,
            lineno: event.lineno,
            colno: event.colno,
            error: event.error?.stack,
        });
    });

    window.addEventListener('unhandledrejection', (event) => {
        logger.error('未処理のPromise拒否', {
            reason: event.reason,
            stack: event.reason?.stack,
        });
    });
}

/*
 * PTA AI業務支援ツール - 設定管理モジュール
 * Copyright (c) 2024 PTA Development Team
 */

import { APP_CONSTANTS } from './constants.js';
import { logger } from './logger.js';
import { handleError, PTAValidationError } from './error-handler.js';
import { eventManager, EVENTS } from './event-manager.js';

/**
 * 設定スキーマ定義
 */
const SETTINGS_SCHEMA = {
    apiSettings: {
        openaiApiKey: { type: 'string', required: true, sensitive: true },
        openaiEndpoint: { type: 'string', default: 'https://api.openai.com/v1/chat/completions' },
        openaiModel: { type: 'string', default: 'gpt-4' },
        maxTokens: { type: 'number', default: 2000, min: 100, max: 4000 },
        temperature: { type: 'number', default: 0.7, min: 0, max: 2 },
        timeout: { type: 'number', default: 30000, min: 5000, max: 120000 },
    },
    userPreferences: {
        language: { type: 'string', default: 'ja', enum: ['ja', 'en'] },
        theme: { type: 'string', default: 'auto', enum: ['light', 'dark', 'auto'] },
        autoAnalyze: { type: 'boolean', default: false },
        showNotifications: { type: 'boolean', default: true },
        debugMode: { type: 'boolean', default: false },
        defaultPromptType: { type: 'string', default: 'analysis', enum: ['analysis', 'summary', 'reply'] },
    },
    themeSettings: {
        primaryColor: { type: 'string', default: '#0078d4' },
        fontSize: { type: 'string', default: 'medium', enum: ['small', 'medium', 'large'] },
        buttonPosition: { type: 'string', default: 'bottom-right', enum: ['top-left', 'top-right', 'bottom-left', 'bottom-right'] },
        animationEnabled: { type: 'boolean', default: true },
    },
    analysisHistory: {
        maxEntries: { type: 'number', default: 100, min: 10, max: 1000 },
        autoSave: { type: 'boolean', default: true },
        retentionDays: { type: 'number', default: 30, min: 1, max: 365 },
    },
    cacheSettings: {
        enabled: { type: 'boolean', default: true },
        maxSize: { type: 'number', default: 10485760, min: 1048576, max: 104857600 }, // 10MB
        ttl: { type: 'number', default: 3600000, min: 300000, max: 86400000 }, // 1時間
    },
};

/**
 * 統一設定管理クラス
 */
export class PTASettingsManager {
    constructor() {
        this.settings = {};
        this.defaultSettings = this._generateDefaultSettings();
        this.isLoaded = false;
        this.watchers = new Map();
        this.encryptionKey = null;
    }

    /**
     * 設定を初期化
     */
    async initialize() {
        try {
            logger.info('設定管理システムを初期化しています...');

            // 暗号化キーの生成または取得
            await this._initializeEncryption();

            // 設定を読み込み
            await this.load();

            // マイグレーション実行
            await this._runMigrations();

            this.isLoaded = true;

            logger.info('設定管理システムの初期化が完了しました');
            eventManager.emit(EVENTS.SETTINGS_LOADED, this.settings);

        } catch (error) {
            handleError(error, '設定初期化');
            throw error;
        }
    }

    /**
     * デフォルト設定を生成
     */
    _generateDefaultSettings() {
        const defaults = {};

        Object.entries(SETTINGS_SCHEMA).forEach(([category, schema]) => {
            defaults[category] = {};
            Object.entries(schema).forEach(([key, config]) => {
                if (config.default !== undefined) {
                    defaults[category][key] = config.default;
                }
            });
        });

        return defaults;
    }

    /**
     * 設定を読み込み
     */
    async load() {
        try {
            const stored = await this._getStoredSettings();

            // デフォルト設定をベースにマージ
            this.settings = this._mergeSettings(this.defaultSettings, stored);

            // バリデーション実行
            this._validateSettings();

            logger.info('設定を読み込みました', { categories: Object.keys(this.settings) });

        } catch (error) {
            logger.warn('設定読み込みエラー、デフォルト設定を使用します', error);
            this.settings = { ...this.defaultSettings };
        }
    }

    /**
     * 設定を保存
     */
    async save() {
        try {
            // バリデーション実行
            this._validateSettings();

            // 機密情報の暗号化
            const settingsToStore = await this._encryptSensitiveData(this.settings);

            // ストレージに保存
            await this._storeSettings(settingsToStore);

            logger.info('設定を保存しました');
            eventManager.emit(EVENTS.SETTINGS_SAVED, this.settings);

        } catch (error) {
            handleError(error, '設定保存');
            throw error;
        }
    }

    /**
     * 設定値を取得
     */
    get(category, key = null, defaultValue = null) {
        if (!this.isLoaded) {
            logger.warn('設定がまだ読み込まれていません');
            return defaultValue;
        }

        if (key === null) {
            return this.settings[category] || defaultValue;
        }

        return this.settings[category]?.[key] ?? defaultValue;
    }

    /**
     * 設定値を設定
     */
    async set(category, key, value = null) {
        // 値が第3引数として渡された場合
        if (value === null && typeof key === 'object') {
            value = key;
            key = null;
        }

        try {
            if (key === null) {
                // カテゴリ全体を設定
                this.settings[category] = value;
            } else {
                // 個別の設定値を設定
                if (!this.settings[category]) {
                    this.settings[category] = {};
                }
                this.settings[category][key] = value;
            }

            // バリデーション実行
            this._validateSetting(category, key, value);

            // 変更を通知
            eventManager.emit(EVENTS.SETTINGS_CHANGED, {
                category,
                key,
                value,
                settings: this.settings,
            });

            // ウォッチャーに通知
            this._notifyWatchers(category, key, value);

            logger.debug('設定値を更新しました', { category, key });

        } catch (error) {
            handleError(error, '設定更新');
            throw error;
        }
    }

    /**
     * 設定の変更を監視
     */
    watch(category, key, callback) {
        const watchKey = key ? `${category}.${key}` : category;

        if (!this.watchers.has(watchKey)) {
            this.watchers.set(watchKey, []);
        }

        this.watchers.get(watchKey).push(callback);

        return () => {
            const callbacks = this.watchers.get(watchKey);
            if (callbacks) {
                const index = callbacks.indexOf(callback);
                if (index > -1) {
                    callbacks.splice(index, 1);
                }
            }
        };
    }

    /**
     * 設定をリセット
     */
    async reset(category = null) {
        try {
            if (category) {
                this.settings[category] = { ...this.defaultSettings[category] };
                logger.info(`設定カテゴリ「${category}」をリセットしました`);
            } else {
                this.settings = { ...this.defaultSettings };
                logger.info('すべての設定をリセットしました');
            }

            await this.save();

        } catch (error) {
            handleError(error, '設定リセット');
            throw error;
        }
    }

    /**
     * 設定をエクスポート
     */
    export(includeDefaults = false) {
        try {
            const exportData = {
                version: APP_CONSTANTS.VERSION,
                timestamp: new Date().toISOString(),
                settings: includeDefaults ? this.settings : this._getNonDefaultSettings(),
            };

            // 機密情報を除外
            const sanitized = this._sanitizeForExport(exportData);

            return JSON.stringify(sanitized, null, 2);

        } catch (error) {
            handleError(error, '設定エクスポート');
            throw error;
        }
    }

    /**
     * 設定をインポート
     */
    async import(importData) {
        try {
            const data = typeof importData === 'string' ? JSON.parse(importData) : importData;

            // バージョンチェック
            if (data.version && data.version !== APP_CONSTANTS.VERSION) {
                logger.warn('異なるバージョンの設定をインポートしています', {
                    current: APP_CONSTANTS.VERSION,
                    import: data.version,
                });
            }

            // 設定をマージ
            if (data.settings) {
                this.settings = this._mergeSettings(this.settings, data.settings);
                await this.save();
                logger.info('設定をインポートしました');
            }

        } catch (error) {
            handleError(error, '設定インポート');
            throw error;
        }
    }

    /**
     * ストレージから設定を取得
     */
    async _getStoredSettings() {
        if (typeof chrome !== 'undefined' && chrome.storage) {
            const result = await chrome.storage.local.get(Object.values(APP_CONSTANTS.STORAGE_KEYS));
            return result;
        } else {
            // フォールバック: localStorage
            const stored = {};
            Object.values(APP_CONSTANTS.STORAGE_KEYS).forEach(key => {
                const value = localStorage.getItem(key);
                if (value) {
                    try {
                        stored[key] = JSON.parse(value);
                    } catch (error) {
                        logger.warn(`ストレージキー ${key} の解析に失敗`, error);
                    }
                }
            });
            return stored;
        }
    }

    /**
     * 設定をストレージに保存
     */
    async _storeSettings(settings) {
        if (typeof chrome !== 'undefined' && chrome.storage) {
            await chrome.storage.local.set(settings);
        } else {
            // フォールバック: localStorage
            Object.entries(settings).forEach(([key, value]) => {
                localStorage.setItem(key, JSON.stringify(value));
            });
        }
    }

    /**
     * 設定をマージ
     */
    _mergeSettings(defaults, stored) {
        const merged = { ...defaults };

        Object.entries(stored).forEach(([key, value]) => {
            if (merged[key] && typeof merged[key] === 'object' && !Array.isArray(merged[key])) {
                merged[key] = { ...merged[key], ...value };
            } else {
                merged[key] = value;
            }
        });

        return merged;
    }

    /**
     * 設定をバリデーション
     */
    _validateSettings() {
        Object.entries(this.settings).forEach(([category, settings]) => {
            if (SETTINGS_SCHEMA[category]) {
                Object.entries(settings).forEach(([key, value]) => {
                    this._validateSetting(category, key, value);
                });
            }
        });
    }

    /**
     * 個別設定をバリデーション
     */
    _validateSetting(category, key, value) {
        const schema = SETTINGS_SCHEMA[category]?.[key];
        if (!schema) return;

        // 型チェック
        if (schema.type && typeof value !== schema.type) {
            throw new PTAValidationError(
                `設定「${category}.${key}」の型が正しくありません。期待: ${schema.type}, 実際: ${typeof value}`,
                `${category}.${key}`,
                value
            );
        }

        // 必須チェック
        if (schema.required && (value === null || value === undefined || value === '')) {
            throw new PTAValidationError(
                `設定「${category}.${key}」は必須です`,
                `${category}.${key}`,
                value
            );
        }

        // 列挙値チェック
        if (schema.enum && !schema.enum.includes(value)) {
            throw new PTAValidationError(
                `設定「${category}.${key}」の値が不正です。有効な値: ${schema.enum.join(', ')}`,
                `${category}.${key}`,
                value
            );
        }

        // 数値範囲チェック
        if (schema.type === 'number') {
            if (schema.min !== undefined && value < schema.min) {
                throw new PTAValidationError(
                    `設定「${category}.${key}」の値が小さすぎます。最小値: ${schema.min}`,
                    `${category}.${key}`,
                    value
                );
            }
            if (schema.max !== undefined && value > schema.max) {
                throw new PTAValidationError(
                    `設定「${category}.${key}」の値が大きすぎます。最大値: ${schema.max}`,
                    `${category}.${key}`,
                    value
                );
            }
        }
    }

    /**
     * ウォッチャーに通知
     */
    _notifyWatchers(category, key, value) {
        const watchKey = key ? `${category}.${key}` : category;
        const callbacks = this.watchers.get(watchKey) || [];

        callbacks.forEach(callback => {
            try {
                callback(value, { category, key });
            } catch (error) {
                logger.error('設定ウォッチャーでエラー', error);
            }
        });
    }

    /**
     * 暗号化の初期化
     */
    async _initializeEncryption() {
        // 簡単な暗号化キーの生成（実際のプロダクションではより強力な方法を使用）
        this.encryptionKey = 'pta-encryption-key';
    }

    /**
     * 機密データの暗号化
     */
    async _encryptSensitiveData(settings) {
        const encrypted = { ...settings };

        // APIキーなどの機密情報を暗号化（簡易実装）
        Object.entries(SETTINGS_SCHEMA).forEach(([category, schema]) => {
            Object.entries(schema).forEach(([key, config]) => {
                if (config.sensitive && encrypted[category]?.[key]) {
                    encrypted[category][key] = `encrypted:${btoa(encrypted[category][key])}`;
                }
            });
        });

        return encrypted;
    }

    /**
     * デフォルト以外の設定を取得
     */
    _getNonDefaultSettings() {
        const nonDefault = {};

        Object.entries(this.settings).forEach(([category, settings]) => {
            Object.entries(settings).forEach(([key, value]) => {
                if (this.defaultSettings[category]?.[key] !== value) {
                    if (!nonDefault[category]) {
                        nonDefault[category] = {};
                    }
                    nonDefault[category][key] = value;
                }
            });
        });

        return nonDefault;
    }

    /**
     * エクスポート用にサニタイズ
     */
    _sanitizeForExport(data) {
        const sanitized = JSON.parse(JSON.stringify(data));

        // 機密情報を除外
        Object.entries(SETTINGS_SCHEMA).forEach(([category, schema]) => {
            Object.entries(schema).forEach(([key, config]) => {
                if (config.sensitive && sanitized.settings[category]?.[key]) {
                    sanitized.settings[category][key] = '[REDACTED]';
                }
            });
        });

        return sanitized;
    }

    /**
     * マイグレーションを実行
     */
    async _runMigrations() {
        // 将来のバージョンアップ時のマイグレーション処理
        logger.debug('設定マイグレーション完了');
    }
}

/**
 * グローバル設定マネージャーインスタンス
 */
export const settingsManager = new PTASettingsManager();

/**
 * 便利な関数群
 */
export const getSetting = (category, key, defaultValue) =>
    settingsManager.get(category, key, defaultValue);

export const setSetting = (category, key, value) =>
    settingsManager.set(category, key, value);

export const watchSetting = (category, key, callback) =>
    settingsManager.watch(category, key, callback);

export const resetSettings = (category) =>
    settingsManager.reset(category);

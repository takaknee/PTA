/*
 * AI業務支援ツール Edge拡張機能 - データモデル定義
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * API設定モデル
 */
export class ApiSettings {
  constructor(data = {}) {
    this.provider = data.provider || 'azure';
    this.apiKey = data.apiKey || '';
    this.model = data.model || 'gpt-4o-mini';
    this.azureEndpoint = data.azureEndpoint || '';
    this.azureDeploymentName = data.azureDeploymentName || '';
    this.maxTokens = data.maxTokens || 2000;
    this.temperature = data.temperature || 0.7;
    this.requestTimeout = data.requestTimeout || 30000;
  }

  /**
   * 設定の妥当性を検証
   * @returns {ValidationResult} 検証結果
   */
  validate() {
    const errors = [];

    if (!this.apiKey) {
      errors.push('APIキーが設定されていません');
    }

    if (this.provider === 'azure') {
      if (!this.azureEndpoint) {
        errors.push('Azure エンドポイントが設定されていません');
      }
      if (!this.azureDeploymentName) {
        errors.push('Azure デプロイメント名が設定されていません');
      }
    }

    if (this.maxTokens < 1 || this.maxTokens > 4000) {
      errors.push('最大トークン数は1-4000の範囲で設定してください');
    }

    if (this.temperature < 0 || this.temperature > 2) {
      errors.push('Temperatureは0-2の範囲で設定してください');
    }

    return new ValidationResult(errors.length === 0, errors);
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      provider: this.provider,
      apiKey: this.apiKey,
      model: this.model,
      azureEndpoint: this.azureEndpoint,
      azureDeploymentName: this.azureDeploymentName,
      maxTokens: this.maxTokens,
      temperature: this.temperature,
      requestTimeout: this.requestTimeout
    };
  }
}

/**
 * 履歴項目モデル
 */
export class HistoryItem {
  constructor(data = {}) {
    this.id = data.id || this.generateId();
    this.type = data.type || '';
    this.timestamp = data.timestamp || Date.now();
    this.title = data.title || '';
    this.content = data.content || '';
    this.result = data.result || '';
    this.metadata = data.metadata || {};
    this.tags = data.tags || [];
  }

  /**
   * 一意のIDを生成
   * @returns {string} 生成されたID
   */
  generateId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * 履歴項目の検索可能テキストを取得
   * @returns {string} 検索用テキスト
   */
  getSearchableText() {
    return [this.title, this.content, this.result, ...this.tags].join(' ').toLowerCase();
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      id: this.id,
      type: this.type,
      timestamp: this.timestamp,
      title: this.title,
      content: this.content,
      result: this.result,
      metadata: this.metadata,
      tags: this.tags
    };
  }
}

/**
 * API リクエストモデル
 */
export class ApiRequest {
  constructor(data = {}) {
    this.endpoint = data.endpoint || '';
    this.headers = data.headers || {};
    this.body = data.body || '';
    this.method = data.method || 'POST';
    this.timeout = data.timeout || 30000;
    this.retryAttempts = data.retryAttempts || 3;
    this.retryDelay = data.retryDelay || 1000;
  }

  /**
   * リクエストの妥当性を検証
   * @returns {ValidationResult} 検証結果
   */
  validate() {
    const errors = [];

    if (!this.endpoint) {
      errors.push('エンドポイントが設定されていません');
    }

    if (!this.body && this.method === 'POST') {
      errors.push('POSTリクエストにはボディが必要です');
    }

    if (this.timeout < 1000 || this.timeout > 60000) {
      errors.push('タイムアウトは1000-60000msの範囲で設定してください');
    }

    return new ValidationResult(errors.length === 0, errors);
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      endpoint: this.endpoint,
      headers: this.headers,
      body: this.body,
      method: this.method,
      timeout: this.timeout,
      retryAttempts: this.retryAttempts,
      retryDelay: this.retryDelay
    };
  }
}

/**
 * API レスポンスモデル
 */
export class ApiResponse {
  constructor(data = {}) {
    this.success = data.success || false;
    this.data = data.data || null;
    this.error = data.error || null;
    this.statusCode = data.statusCode || 0;
    this.timestamp = data.timestamp || Date.now();
    this.requestId = data.requestId || null;
    this.duration = data.duration || 0;
  }

  /**
   * 成功レスポンスかどうかを判定
   * @returns {boolean} 成功の場合true
   */
  isSuccess() {
    return this.success && !this.error;
  }

  /**
   * エラーレスポンスかどうかを判定
   * @returns {boolean} エラーの場合true
   */
  isError() {
    return !this.success || this.error !== null;
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      success: this.success,
      data: this.data,
      error: this.error,
      statusCode: this.statusCode,
      timestamp: this.timestamp,
      requestId: this.requestId,
      duration: this.duration
    };
  }
}

/**
 * メッセージモデル
 */
export class Message {
  constructor(data = {}) {
    this.action = data.action || '';
    this.target = data.target || '';
    this.data = data.data || {};
    this.timestamp = data.timestamp || Date.now();
    this.requestId = data.requestId || this.generateId();
    this.sender = data.sender || '';
  }

  /**
   * 一意のIDを生成
   * @returns {string} 生成されたID
   */
  generateId() {
    return `msg-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      action: this.action,
      target: this.target,
      data: this.data,
      timestamp: this.timestamp,
      requestId: this.requestId,
      sender: this.sender
    };
  }
}

/**
 * 診断情報モデル
 */
export class DiagnosticsInfo {
  constructor(data = {}) {
    this.timestamp = data.timestamp || new Date().toISOString();
    this.chrome = data.chrome || {};
    this.permissions = data.permissions || {};
    this.offscreenDocument = data.offscreenDocument || {};
    this.network = data.network || {};
    this.apiStatus = data.apiStatus || {};
    this.settings = data.settings || {};
  }

  /**
   * 診断結果の概要を取得
   * @returns {string} 診断結果の概要
   */
  getSummary() {
    const issues = [];

    if (!this.network.basicConnectivity) {
      issues.push('ネットワーク接続に問題があります');
    }

    if (!this.offscreenDocument.canCreate) {
      issues.push('Offscreen Documentが利用できません');
    }

    if (!this.apiStatus.configured) {
      issues.push('API設定が不完全です');
    }

    return issues.length === 0 ? '正常' : issues.join(', ');
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      timestamp: this.timestamp,
      chrome: this.chrome,
      permissions: this.permissions,
      offscreenDocument: this.offscreenDocument,
      network: this.network,
      apiStatus: this.apiStatus,
      settings: this.settings
    };
  }
}

/**
 * 検証結果モデル
 */
export class ValidationResult {
  constructor(isValid = false, errors = []) {
    this.isValid = isValid;
    this.errors = errors;
  }

  /**
   * エラーメッセージを文字列として取得
   * @returns {string} エラーメッセージ
   */
  getErrorMessage() {
    return this.errors.join(', ');
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      isValid: this.isValid,
      errors: this.errors
    };
  }
}

/**
 * 統計情報モデル
 */
export class Statistics {
  constructor(data = {}) {
    this.totalRequests = data.totalRequests || 0;
    this.successfulRequests = data.successfulRequests || 0;
    this.failedRequests = data.failedRequests || 0;
    this.averageResponseTime = data.averageResponseTime || 0;
    this.lastRequestTime = data.lastRequestTime || null;
    this.requestsByType = data.requestsByType || {};
    this.requestsByDay = data.requestsByDay || {};
  }

  /**
   * 成功率を計算
   * @returns {number} 成功率（0-100）
   */
  getSuccessRate() {
    if (this.totalRequests === 0) return 0;
    return Math.round((this.successfulRequests / this.totalRequests) * 100);
  }

  /**
   * リクエストを記録
   * @param {string} type リクエストタイプ
   * @param {boolean} success 成功かどうか
   * @param {number} responseTime レスポンス時間（ms）
   */
  recordRequest(type, success, responseTime) {
    this.totalRequests++;

    if (success) {
      this.successfulRequests++;
    } else {
      this.failedRequests++;
    }

    // 平均レスポンス時間の更新
    this.averageResponseTime = (this.averageResponseTime * (this.totalRequests - 1) + responseTime) / this.totalRequests;

    // 最後のリクエスト時間を更新
    this.lastRequestTime = Date.now();

    // タイプ別統計の更新
    if (!this.requestsByType[type]) {
      this.requestsByType[type] = 0;
    }
    this.requestsByType[type]++;

    // 日別統計の更新
    const today = new Date().toISOString().split('T')[0];
    if (!this.requestsByDay[today]) {
      this.requestsByDay[today] = 0;
    }
    this.requestsByDay[today]++;
  }

  /**
   * JSONオブジェクトに変換
   * @returns {Object} JSON形式のデータ
   */
  toJSON() {
    return {
      totalRequests: this.totalRequests,
      successfulRequests: this.successfulRequests,
      failedRequests: this.failedRequests,
      averageResponseTime: this.averageResponseTime,
      lastRequestTime: this.lastRequestTime,
      requestsByType: this.requestsByType,
      requestsByDay: this.requestsByDay
    };
  }
}

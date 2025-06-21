/*
 * AI業務支援ツール Edge拡張機能 - 統計エンティティ
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * 統計エンティティ
 */
export class Statistics {
  constructor(data = {}) {
    this.totalRequests = data.totalRequests || 0;
    this.successfulRequests = data.successfulRequests || 0;
    this.failedRequests = data.failedRequests || 0;
    this.totalTokensUsed = data.totalTokensUsed || 0;
    this.totalDuration = data.totalDuration || 0;
    this.usageByModel = data.usageByModel || {};
    this.usageByType = data.usageByType || {};
    this.usageByDate = data.usageByDate || {};
    this.errorsByType = data.errorsByType || {};
    this.averageResponseTime = data.averageResponseTime || 0;
    this.firstUsed = data.firstUsed || null;
    this.lastUpdated = data.lastUpdated || new Date().toISOString();
  }

  /**
   * リクエストの記録
   * @param {Object} request - リクエスト情報
   */
  recordRequest(request) {
    this.totalRequests++;

    if (request.success) {
      this.successfulRequests++;

      // トークン使用量の記録
      if (request.tokenUsage) {
        this.totalTokensUsed += request.tokenUsage.total || 0;
      }

      // モデル別使用量の記録
      if (request.model) {
        if (!this.usageByModel[request.model]) {
          this.usageByModel[request.model] = {
            requests: 0,
            tokens: 0,
            duration: 0
          };
        }
        this.usageByModel[request.model].requests++;
        this.usageByModel[request.model].tokens += request.tokenUsage?.total || 0;
        this.usageByModel[request.model].duration += request.duration || 0;
      }
    } else {
      this.failedRequests++;

      // エラータイプ別の記録
      if (request.errorType) {
        this.errorsByType[request.errorType] = (this.errorsByType[request.errorType] || 0) + 1;
      }
    }

    // タイプ別使用量の記録
    if (request.type) {
      if (!this.usageByType[request.type]) {
        this.usageByType[request.type] = {
          requests: 0,
          successful: 0,
          failed: 0,
          tokens: 0
        };
      }
      this.usageByType[request.type].requests++;
      if (request.success) {
        this.usageByType[request.type].successful++;
        this.usageByType[request.type].tokens += request.tokenUsage?.total || 0;
      } else {
        this.usageByType[request.type].failed++;
      }
    }

    // 日付別使用量の記録
    const today = new Date().toISOString().split('T')[0];
    if (!this.usageByDate[today]) {
      this.usageByDate[today] = {
        requests: 0,
        tokens: 0,
        duration: 0
      };
    }
    this.usageByDate[today].requests++;
    this.usageByDate[today].tokens += request.tokenUsage?.total || 0;
    this.usageByDate[today].duration += request.duration || 0;

    // 実行時間の記録
    if (request.duration) {
      this.totalDuration += request.duration;
      this.averageResponseTime = this.totalDuration / this.totalRequests;
    }

    // 初回利用日の記録
    if (!this.firstUsed) {
      this.firstUsed = new Date().toISOString();
    }

    this.lastUpdated = new Date().toISOString();
  }

  /**
   * 成功率の計算
   * @returns {number} 成功率（0-100）
   */
  getSuccessRate() {
    if (this.totalRequests === 0) return 0;
    return (this.successfulRequests / this.totalRequests) * 100;
  }

  /**
   * 平均トークン使用量の計算
   * @returns {number} 平均トークン使用量
   */
  getAverageTokenUsage() {
    if (this.successfulRequests === 0) return 0;
    return this.totalTokensUsed / this.successfulRequests;
  }

  /**
   * 今日の使用統計の取得
   * @returns {Object} 今日の統計
   */
  getTodayUsage() {
    const today = new Date().toISOString().split('T')[0];
    return this.usageByDate[today] || {
      requests: 0,
      tokens: 0,
      duration: 0
    };
  }

  /**
   * 週間使用統計の取得
   * @returns {Object} 週間統計
   */
  getWeeklyUsage() {
    const weekAgo = new Date();
    weekAgo.setDate(weekAgo.getDate() - 7);

    const weekly = {
      requests: 0,
      tokens: 0,
      duration: 0,
      days: []
    };

    for (let i = 0; i < 7; i++) {
      const date = new Date();
      date.setDate(date.getDate() - i);
      const dateStr = date.toISOString().split('T')[0];

      const dayData = this.usageByDate[dateStr] || {
        requests: 0,
        tokens: 0,
        duration: 0
      };

      weekly.requests += dayData.requests;
      weekly.tokens += dayData.tokens;
      weekly.duration += dayData.duration;
      weekly.days.unshift({
        date: dateStr,
        ...dayData
      });
    }

    return weekly;
  }

  /**
   * 最も使用されているモデルの取得
   * @returns {Object} 最も使用されているモデル情報
   */
  getMostUsedModel() {
    let maxUsage = 0;
    let mostUsedModel = null;

    Object.entries(this.usageByModel).forEach(([model, usage]) => {
      if (usage.requests > maxUsage) {
        maxUsage = usage.requests;
        mostUsedModel = { model, ...usage };
      }
    });

    return mostUsedModel;
  }

  /**
   * 最も多いエラータイプの取得
   * @returns {Object} 最も多いエラー情報
   */
  getMostCommonError() {
    let maxCount = 0;
    let mostCommonError = null;

    Object.entries(this.errorsByType).forEach(([errorType, count]) => {
      if (count > maxCount) {
        maxCount = count;
        mostCommonError = { errorType, count };
      }
    });

    return mostCommonError;
  }

  /**
   * 古い日付データのクリーンアップ
   * @param {number} daysToKeep - 保持日数（デフォルト: 30日）
   */
  cleanupOldData(daysToKeep = 30) {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    const cutoffStr = cutoffDate.toISOString().split('T')[0];

    Object.keys(this.usageByDate).forEach(date => {
      if (date < cutoffStr) {
        delete this.usageByDate[date];
      }
    });

    this.lastUpdated = new Date().toISOString();
  }

  /**
   * 統計のリセット
   */
  reset() {
    this.totalRequests = 0;
    this.successfulRequests = 0;
    this.failedRequests = 0;
    this.totalTokensUsed = 0;
    this.totalDuration = 0;
    this.usageByModel = {};
    this.usageByType = {};
    this.usageByDate = {};
    this.errorsByType = {};
    this.averageResponseTime = 0;
    this.firstUsed = null;
    this.lastUpdated = new Date().toISOString();
  }

  /**
   * JSON表現への変換
   * @returns {Object} JSONオブジェクト
   */
  toJSON() {
    return {
      totalRequests: this.totalRequests,
      successfulRequests: this.successfulRequests,
      failedRequests: this.failedRequests,
      totalTokensUsed: this.totalTokensUsed,
      totalDuration: this.totalDuration,
      usageByModel: { ...this.usageByModel },
      usageByType: { ...this.usageByType },
      usageByDate: { ...this.usageByDate },
      errorsByType: { ...this.errorsByType },
      averageResponseTime: this.averageResponseTime,
      firstUsed: this.firstUsed,
      lastUpdated: this.lastUpdated
    };
  }

  /**
   * JSONからの復元
   * @param {Object} json - JSONオブジェクト
   * @returns {Statistics} 統計インスタンス
   */
  static fromJSON(json) {
    return new Statistics(json);
  }

  /**
   * 統計の概要取得
   * @returns {Object} 統計概要
   */
  getSummary() {
    return {
      totalRequests: this.totalRequests,
      successRate: this.getSuccessRate(),
      totalTokensUsed: this.totalTokensUsed,
      averageTokenUsage: this.getAverageTokenUsage(),
      averageResponseTime: this.averageResponseTime,
      todayUsage: this.getTodayUsage(),
      mostUsedModel: this.getMostUsedModel(),
      firstUsed: this.firstUsed,
      lastUpdated: this.lastUpdated
    };
  }
}

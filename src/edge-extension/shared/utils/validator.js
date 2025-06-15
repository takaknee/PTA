/*
 * AI業務支援ツール Edge拡張機能 - バリデーションユーティリティ
 * Copyright (c) 2024 AI Business Support Team
 */

import { API_VALIDATION, ERROR_CODES } from '../constants/index.js';

/**
 * バリデーションエラークラス
 */
export class ValidationError extends Error {
  constructor(code, message, field = null) {
    super(message);
    this.name = 'ValidationError';
    this.code = code;
    this.field = field;
  }
}

/**
 * バリデーションユーティリティクラス
 */
export class Validator {
  /**
   * APIキーのバリデーション
   * @param {string} apiKey - APIキー
   * @returns {boolean} バリデーション結果
   * @throws {ValidationError} バリデーションエラー
   */
  static validateApiKey(apiKey) {
    if (!apiKey) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_REQUIRED_FIELD,
        'APIキーは必須です',
        'apiKey'
      );
    }

    if (typeof apiKey !== 'string') {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_INVALID_FORMAT,
        'APIキーは文字列である必要があります',
        'apiKey'
      );
    }

    if (apiKey.length < API_VALIDATION.API_KEY_MIN_LENGTH) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_LENGTH_EXCEEDED,
        `APIキーは${API_VALIDATION.API_KEY_MIN_LENGTH}文字以上である必要があります`,
        'apiKey'
      );
    }

    if (!API_VALIDATION.API_KEY_PATTERN.test(apiKey)) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_INVALID_FORMAT,
        'APIキーの形式が正しくありません',
        'apiKey'
      );
    }

    return true;
  }

  /**
   * Azureエンドポイントのバリデーション
   * @param {string} endpoint - Azureエンドポイント
   * @returns {boolean} バリデーション結果
   * @throws {ValidationError} バリデーションエラー
   */
  static validateAzureEndpoint(endpoint) {
    if (!endpoint) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_REQUIRED_FIELD,
        'Azureエンドポイントは必須です',
        'azureEndpoint'
      );
    }

    try {
      const url = new URL(endpoint);

      if (url.protocol !== 'https:') {
        throw new ValidationError(
          ERROR_CODES.VALIDATION_INVALID_FORMAT,
          'HTTPSプロトコルが必要です',
          'azureEndpoint'
        );
      }

      if (!API_VALIDATION.AZURE_ENDPOINT_PATTERN.test(endpoint)) {
        throw new ValidationError(
          ERROR_CODES.VALIDATION_INVALID_FORMAT,
          'Azure OpenAIエンドポイントの形式が正しくありません',
          'azureEndpoint'
        );
      }
    } catch (error) {
      if (error instanceof ValidationError) throw error;

      throw new ValidationError(
        ERROR_CODES.VALIDATION_INVALID_FORMAT,
        'エンドポイントURLの形式が正しくありません',
        'azureEndpoint'
      );
    }

    return true;
  }

  /**
   * デプロイメント名のバリデーション
   * @param {string} deploymentName - デプロイメント名
   * @returns {boolean} バリデーション結果
   * @throws {ValidationError} バリデーションエラー
   */
  static validateDeploymentName(deploymentName) {
    if (!deploymentName) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_REQUIRED_FIELD,
        'デプロイメント名は必須です',
        'azureDeploymentName'
      );
    }

    if (!API_VALIDATION.DEPLOYMENT_NAME_PATTERN.test(deploymentName)) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_INVALID_FORMAT,
        'デプロイメント名の形式が正しくありません',
        'azureDeploymentName'
      );
    }

    return true;
  }

  /**
   * テキスト長のバリデーション
   * @param {string} text - テキスト
   * @param {number} maxLength - 最大長
   * @param {string} fieldName - フィールド名
   * @returns {boolean} バリデーション結果
   * @throws {ValidationError} バリデーションエラー
   */
  static validateTextLength(text, maxLength = API_VALIDATION.MAX_INPUT_LENGTH, fieldName = 'text') {
    if (!text) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_REQUIRED_FIELD,
        `${fieldName}は必須です`,
        fieldName
      );
    }

    if (text.length > maxLength) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_LENGTH_EXCEEDED,
        `${fieldName}は${maxLength}文字以下である必要があります`,
        fieldName
      );
    }

    return true;
  }

  /**
   * 数値範囲のバリデーション
   * @param {number} value - 値
   * @param {number} min - 最小値
   * @param {number} max - 最大値
   * @param {string} fieldName - フィールド名
   * @returns {boolean} バリデーション結果
   * @throws {ValidationError} バリデーションエラー
   */
  static validateNumberRange(value, min, max, fieldName = 'value') {
    if (typeof value !== 'number' || isNaN(value)) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_INVALID_FORMAT,
        `${fieldName}は数値である必要があります`,
        fieldName
      );
    }

    if (value < min || value > max) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_INVALID_VALUE,
        `${fieldName}は${min}から${max}の間である必要があります`,
        fieldName
      );
    }

    return true;
  }

  /**
   * 複合設定のバリデーション
   * @param {Object} settings - 設定オブジェクト
   * @returns {boolean} バリデーション結果
   * @throws {ValidationError} バリデーションエラー
   */
  static validateSettings(settings) {
    if (!settings || typeof settings !== 'object') {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_INVALID_FORMAT,
        '設定オブジェクトが無効です',
        'settings'
      );
    }

    const { api } = settings;
    if (!api) {
      throw new ValidationError(
        ERROR_CODES.VALIDATION_REQUIRED_FIELD,
        'API設定は必須です',
        'api'
      );
    }

    // APIキーのバリデーション
    this.validateApiKey(api.apiKey);

    // Azureプロバイダーの場合の追加チェック
    if (api.provider === 'azure') {
      this.validateAzureEndpoint(api.azureEndpoint);
      this.validateDeploymentName(api.azureDeploymentName);
    }

    // トークン数の範囲チェック
    if (api.maxTokens !== undefined) {
      this.validateNumberRange(api.maxTokens, 1, 4000, 'maxTokens');
    }

    // 温度パラメータの範囲チェック
    if (api.temperature !== undefined) {
      this.validateNumberRange(api.temperature, 0, 2, 'temperature');
    }

    return true;
  }

  /**
   * バリデーション結果の収集
   * @param {Function} validationFn - バリデーション関数
   * @returns {Object} バリデーション結果
   */
  static collectValidationResult(validationFn) {
    try {
      validationFn();
      return { valid: true, errors: [] };
    } catch (error) {
      if (error instanceof ValidationError) {
        return {
          valid: false,
          errors: [{
            code: error.code,
            message: error.message,
            field: error.field
          }]
        };
      }

      return {
        valid: false,
        errors: [{
          code: ERROR_CODES.SYSTEM_UNKNOWN_ERROR,
          message: error.message,
          field: null
        }]
      };
    }
  }

  /**
   * 複数のバリデーション結果の統合
   * @param {Function[]} validationFns - バリデーション関数配列
   * @returns {Object} 統合されたバリデーション結果
   */
  static collectMultipleValidationResults(validationFns) {
    const results = validationFns.map(fn => this.collectValidationResult(fn));
    const allErrors = results.flatMap(result => result.errors);

    return {
      valid: allErrors.length === 0,
      errors: allErrors
    };
  }
}

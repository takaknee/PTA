/*
 * AI業務支援ツール Edge拡張機能 - 共通型定義
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * API設定の型定義
 * @typedef {Object} ApiSettings
 * @property {string} provider - APIプロバイダー ('openai' | 'azure')
 * @property {string} model - AIモデル名
 * @property {string} apiKey - APIキー
 * @property {string} azureEndpoint - Azureエンドポイント（Azureの場合）
 * @property {string} azureDeploymentName - Azureデプロイメント名（Azureの場合）
 * @property {number} maxTokens - 最大トークン数
 * @property {number} temperature - 温度パラメータ
 * @property {number} requestTimeout - リクエストタイムアウト（ミリ秒）
 */

/**
 * 拡張機能設定の型定義
 * @typedef {Object} ExtensionSettings
 * @property {ApiSettings} api - API設定
 * @property {boolean} autoDetect - 自動検出有効
 * @property {boolean} showNotifications - 通知表示有効
 * @property {boolean} saveHistory - 履歴保存有効
 * @property {string} theme - テーマ ('light' | 'dark' | 'auto')
 * @property {Object} buttonPosition - ボタン位置設定
 */

/**
 * 履歴項目の型定義
 * @typedef {Object} HistoryItem
 * @property {string} id - 一意識別子
 * @property {string} type - 処理タイプ
 * @property {string} timestamp - タイムスタンプ
 * @property {string} title - タイトル
 * @property {string} content - 入力コンテンツ
 * @property {string} result - 処理結果
 * @property {string} service - 実行されたサービス
 * @property {Object} metadata - メタデータ
 */

/**
 * 統計情報の型定義
 * @typedef {Object} Statistics
 * @property {number} totalRequests - 総リクエスト数
 * @property {number} successfulRequests - 成功リクエスト数
 * @property {number} failedRequests - 失敗リクエスト数
 * @property {number} totalTokensUsed - 総使用トークン数
 * @property {Object} usageByModel - モデル別使用統計
 * @property {string} lastUpdated - 最終更新日時
 */

/**
 * APIレスポンスの型定義
 * @typedef {Object} ApiResponse
 * @property {boolean} success - 成功フラグ
 * @property {string|Object} data - レスポンスデータ
 * @property {string} [error] - エラーメッセージ（失敗時）
 * @property {Object} [metadata] - メタデータ
 */

/**
 * メッセージの型定義
 * @typedef {Object} ExtensionMessage
 * @property {string} action - アクション名
 * @property {Object} [data] - データ
 * @property {string} [target] - ターゲット
 * @property {string} [source] - 送信元
 * @property {string} [id] - メッセージID
 */

/**
 * 通知の型定義
 * @typedef {Object} Notification
 * @property {string} type - 通知タイプ ('success' | 'error' | 'warning' | 'info')
 * @property {string} message - メッセージ
 * @property {string} [title] - タイトル
 * @property {number} [duration] - 表示時間（ミリ秒）
 * @property {boolean} [persistent] - 永続表示フラグ
 */

// 型定義をエクスポート（JSDocとして使用）
export { };

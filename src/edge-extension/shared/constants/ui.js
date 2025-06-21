/*
 * AI業務支援ツール Edge拡張機能 - UI定数
 * Copyright (c) 2024 AI Business Support Team
 */

/**
 * HTML要素ID定数
 */
export const ELEMENT_IDS = Object.freeze({
  AI_BUTTON: 'ai-support-button',
  AI_DIALOG: 'ai-dialog',
  AI_OVERLAY: 'ai-overlay',
  NOTIFICATION_CONTAINER: 'ai-notification-container',
  LOADING_INDICATOR: 'ai-loading-indicator',
  RESULTS_CONTAINER: 'ai-results-container',
  EXPAND_MODAL: 'ai-expand-modal'
});

/**
 * CSS クラス名定数
 */
export const CSS_CLASSES = Object.freeze({
  // 基本コンポーネント
  AI_BUTTON: 'ai-support-button',
  AI_DIALOG: 'ai-dialog',
  AI_OVERLAY: 'ai-overlay',
  AI_MODAL: 'ai-modal',

  // 状態クラス
  LOADING: 'ai-loading',
  SUCCESS: 'ai-success',
  ERROR: 'ai-error',
  WARNING: 'ai-warning',
  INFO: 'ai-info',
  ACTIVE: 'ai-active',
  DISABLED: 'ai-disabled',
  HIDDEN: 'ai-hidden',

  // テーマクラス
  DARK_THEME: 'ai-dark-theme',
  LIGHT_THEME: 'ai-light-theme',

  // レイアウトクラス
  FLOATING_BUTTON: 'ai-floating-button',
  INLINE_BUTTON: 'ai-inline-button',
  DIALOG_OVERLAY: 'ai-dialog-overlay',
  DIALOG_CONTENT: 'ai-dialog-content',
  EXPAND_VIEW: 'ai-expand-view',

  // アニメーションクラス
  FADE_IN: 'ai-fade-in',
  FADE_OUT: 'ai-fade-out',
  SLIDE_IN: 'ai-slide-in',
  SLIDE_OUT: 'ai-slide-out',
  SPIN: 'ai-spin'
});

/**
 * 色定数（ダークテーマ対応）
 */
export const COLORS = Object.freeze({
  LIGHT: {
    PRIMARY: '#2196F3',
    SUCCESS: '#4CAF50',
    WARNING: '#FF9800',
    ERROR: '#f44336',
    INFO: '#2196F3',

    BACKGROUND: '#ffffff',
    SURFACE: '#f5f5f5',
    BORDER: '#e0e0e0',
    TEXT: '#333333',
    TEXT_SECONDARY: '#666666',

    SHADOW: 'rgba(0, 0, 0, 0.1)',
    OVERLAY: 'rgba(0, 0, 0, 0.5)'
  },

  DARK: {
    PRIMARY: '#1976D2',
    SUCCESS: '#388E3C',
    WARNING: '#F57C00',
    ERROR: '#D32F2F',
    INFO: '#1976D2',

    BACKGROUND: '#1e1e1e',
    SURFACE: '#2d2d2d',
    BORDER: '#404040',
    TEXT: '#ffffff',
    TEXT_SECONDARY: '#b0b0b0',

    SHADOW: 'rgba(0, 0, 0, 0.3)',
    OVERLAY: 'rgba(0, 0, 0, 0.7)'
  }
});

/**
 * アニメーション設定
 */
export const ANIMATIONS = Object.freeze({
  DURATION: {
    FAST: 150,
    NORMAL: 300,
    SLOW: 500
  },

  EASING: {
    EASE_IN: 'ease-in',
    EASE_OUT: 'ease-out',
    EASE_IN_OUT: 'ease-in-out'
  }
});

/**
 * Z-Indexレベル
 */
export const Z_INDEX = Object.freeze({
  FLOATING_BUTTON: 10000,
  DIALOG_OVERLAY: 10001,
  DIALOG_CONTENT: 10002,
  NOTIFICATION: 10003,
  EXPAND_MODAL: 10004,
  LOADING_OVERLAY: 10005
});

/**
 * ブレークポイント
 */
export const BREAKPOINTS = Object.freeze({
  MOBILE: 768,
  TABLET: 1024,
  DESKTOP: 1440
});

/**
 * UI設定
 */
export const UI_CONFIG = Object.freeze({
  NOTIFICATION_DURATION: 3000,
  TOOLTIP_DELAY: 500,
  DEBOUNCE_DELAY: 300,
  TYPING_INDICATOR_DELAY: 1000,

  MAX_DIALOG_WIDTH: 800,
  MAX_DIALOG_HEIGHT: 600,

  BUTTON_SIZE: {
    SMALL: 32,
    MEDIUM: 40,
    LARGE: 48
  }
});

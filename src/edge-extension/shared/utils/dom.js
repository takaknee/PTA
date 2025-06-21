/*
 * AI業務支援ツール Edge拡張機能 - DOMユーティリティ
 * Copyright (c) 2024 AI Business Support Team
 */

import { ELEMENT_IDS, CSS_CLASSES, Z_INDEX, ANIMATIONS } from '../constants/index.js';
import { createLogger } from './logger.js';

const logger = createLogger('DOMUtils');

/**
 * DOMユーティリティクラス
 */
export class DOMUtils {
  /**
   * 安全な要素作成
   * @param {string} tagName - タグ名
   * @param {Object} options - オプション
   * @returns {HTMLElement} 作成された要素
   */
  static createElement(tagName, options = {}) {
    try {
      const element = document.createElement(tagName);

      // 属性の設定
      if (options.id) element.id = options.id;
      if (options.className) element.className = options.className;
      if (options.textContent) element.textContent = options.textContent;
      if (options.innerHTML) element.innerHTML = options.innerHTML;

      // スタイルの設定
      if (options.style) {
        Object.assign(element.style, options.style);
      }

      // その他の属性
      if (options.attributes) {
        Object.entries(options.attributes).forEach(([key, value]) => {
          element.setAttribute(key, value);
        });
      }

      return element;
    } catch (error) {
      logger.error('要素作成エラー', { tagName, options, error: error.message });
      throw error;
    }
  }

  /**
   * 安全な要素取得
   * @param {string} selector - セレクタ
   * @param {HTMLElement} parent - 親要素（省略時はdocument）
   * @returns {HTMLElement|null} 見つかった要素
   */
  static querySelector(selector, parent = document) {
    try {
      return parent.querySelector(selector);
    } catch (error) {
      logger.warn('要素取得エラー', { selector, error: error.message });
      return null;
    }
  }

  /**
   * 安全な複数要素取得
   * @param {string} selector - セレクタ
   * @param {HTMLElement} parent - 親要素（省略時はdocument）
   * @returns {NodeList} 見つかった要素のリスト
   */
  static querySelectorAll(selector, parent = document) {
    try {
      return parent.querySelectorAll(selector);
    } catch (error) {
      logger.warn('複数要素取得エラー', { selector, error: error.message });
      return [];
    }
  }

  /**
   * 要素の安全な削除
   * @param {string|HTMLElement} target - 削除対象（セレクタまたは要素）
   * @returns {boolean} 削除成功フラグ
   */
  static removeElement(target) {
    try {
      const element = typeof target === 'string'
        ? this.querySelector(target)
        : target;

      if (element && element.parentNode) {
        element.parentNode.removeChild(element);
        return true;
      }
      return false;
    } catch (error) {
      logger.warn('要素削除エラー', { target, error: error.message });
      return false;
    }
  }

  /**
   * クラスの安全な追加
   * @param {HTMLElement} element - 対象要素
   * @param {...string} classNames - クラス名
   */
  static addClass(element, ...classNames) {
    if (element && element.classList) {
      classNames.forEach(className => {
        if (className) element.classList.add(className);
      });
    }
  }

  /**
   * クラスの安全な削除
   * @param {HTMLElement} element - 対象要素
   * @param {...string} classNames - クラス名
   */
  static removeClass(element, ...classNames) {
    if (element && element.classList) {
      classNames.forEach(className => {
        if (className) element.classList.remove(className);
      });
    }
  }

  /**
   * クラスの切り替え
   * @param {HTMLElement} element - 対象要素
   * @param {string} className - クラス名
   * @param {boolean} force - 強制フラグ
   * @returns {boolean} 切り替え後の状態
   */
  static toggleClass(element, className, force = undefined) {
    if (element && element.classList && className) {
      return element.classList.toggle(className, force);
    }
    return false;
  }

  /**
   * スタイルの安全な設定
   * @param {HTMLElement} element - 対象要素
   * @param {Object} styles - スタイルオブジェクト
   */
  static setStyles(element, styles) {
    if (element && element.style && styles) {
      Object.entries(styles).forEach(([property, value]) => {
        try {
          element.style[property] = value;
        } catch (error) {
          logger.warn('スタイル設定エラー', { property, value, error: error.message });
        }
      });
    }
  }

  /**
   * イベントリスナーの安全な追加
   * @param {HTMLElement} element - 対象要素
   * @param {string} eventType - イベントタイプ
   * @param {Function} handler - ハンドラー関数
   * @param {Object} options - オプション
   */
  static addEventListener(element, eventType, handler, options = {}) {
    if (element && typeof handler === 'function') {
      try {
        element.addEventListener(eventType, handler, options);
      } catch (error) {
        logger.error('イベントリスナー追加エラー', { eventType, error: error.message });
      }
    }
  }

  /**
   * イベントリスナーの安全な削除
   * @param {HTMLElement} element - 対象要素
   * @param {string} eventType - イベントタイプ
   * @param {Function} handler - ハンドラー関数
   * @param {Object} options - オプション
   */
  static removeEventListener(element, eventType, handler, options = {}) {
    if (element && typeof handler === 'function') {
      try {
        element.removeEventListener(eventType, handler, options);
      } catch (error) {
        logger.error('イベントリスナー削除エラー', { eventType, error: error.message });
      }
    }
  }

  /**
   * 要素の表示/非表示切り替え
   * @param {HTMLElement} element - 対象要素
   * @param {boolean} visible - 表示フラグ
   * @param {string} displayType - 表示タイプ
   */
  static setVisible(element, visible, displayType = 'block') {
    if (element) {
      element.style.display = visible ? displayType : 'none';
    }
  }

  /**
   * アニメーション付き表示
   * @param {HTMLElement} element - 対象要素
   * @param {string} animationClass - アニメーションクラス
   * @param {number} duration - 継続時間
   * @returns {Promise} アニメーション完了Promise
   */
  static animateShow(element, animationClass = CSS_CLASSES.FADE_IN, duration = ANIMATIONS.DURATION.NORMAL) {
    return new Promise(resolve => {
      if (!element) {
        resolve();
        return;
      }

      this.addClass(element, animationClass);
      this.setVisible(element, true);

      setTimeout(() => {
        this.removeClass(element, animationClass);
        resolve();
      }, duration);
    });
  }

  /**
   * アニメーション付き非表示
   * @param {HTMLElement} element - 対象要素
   * @param {string} animationClass - アニメーションクラス
   * @param {number} duration - 継続時間
   * @returns {Promise} アニメーション完了Promise
   */
  static animateHide(element, animationClass = CSS_CLASSES.FADE_OUT, duration = ANIMATIONS.DURATION.NORMAL) {
    return new Promise(resolve => {
      if (!element) {
        resolve();
        return;
      }

      this.addClass(element, animationClass);

      setTimeout(() => {
        this.setVisible(element, false);
        this.removeClass(element, animationClass);
        resolve();
      }, duration);
    });
  }

  /**
   * 要素の位置取得
   * @param {HTMLElement} element - 対象要素
   * @returns {Object} 位置情報
   */
  static getElementPosition(element) {
    if (!element) return null;

    const rect = element.getBoundingClientRect();
    return {
      top: rect.top + window.pageYOffset,
      left: rect.left + window.pageXOffset,
      width: rect.width,
      height: rect.height,
      bottom: rect.bottom + window.pageYOffset,
      right: rect.right + window.pageXOffset
    };
  }

  /**
   * ビューポート内判定
   * @param {HTMLElement} element - 対象要素
   * @returns {boolean} ビューポート内フラグ
   */
  static isInViewport(element) {
    if (!element) return false;

    const rect = element.getBoundingClientRect();
    return (
      rect.top >= 0 &&
      rect.left >= 0 &&
      rect.bottom <= (window.innerHeight || document.documentElement.clientHeight) &&
      rect.right <= (window.innerWidth || document.documentElement.clientWidth)
    );
  }

  /**
   * 要素のスムーズスクロール
   * @param {HTMLElement} element - 対象要素
   * @param {Object} options - スクロールオプション
   */
  static scrollToElement(element, options = {}) {
    if (!element) return;

    const defaultOptions = {
      behavior: 'smooth',
      block: 'center',
      inline: 'nearest'
    };

    try {
      element.scrollIntoView({ ...defaultOptions, ...options });
    } catch (error) {
      // フォールバック: 標準的なスクロール
      const position = this.getElementPosition(element);
      if (position) {
        window.scrollTo({
          top: position.top - 100,
          behavior: 'smooth'
        });
      }
    }
  }

  /**
   * 子要素の全削除
   * @param {HTMLElement} element - 対象要素
   */
  static clearChildren(element) {
    if (element) {
      while (element.firstChild) {
        element.removeChild(element.firstChild);
      }
    }
  }

  /**
   * 安全なinnerHTML設定
   * @param {HTMLElement} element - 対象要素
   * @param {string} html - HTML文字列
   */
  static safeSetInnerHTML(element, html) {
    if (!element) return;

    try {
      // 統一セキュリティサニタイザーを使用
      if (typeof globalThis !== 'undefined' && globalThis.unifiedSanitizer) {
        const sanitized = globalThis.unifiedSanitizer.sanitizeHTML(html);
        element.innerHTML = sanitized;
      } else {
        // セキュリティサニタイザーが利用できない場合はプレーンテキストとして設定
        console.warn('統一セキュリティサニタイザーが利用できません。プレーンテキストとして設定します。');
        element.textContent = html;
      }
    } catch (error) {
      logger.error('innerHTML設定エラー', { error: error.message });
      element.textContent = html; // フォールバック
    }
  }
}

/**
 * テーマ検出ユーティリティ
 */
export class ThemeUtils {
  /**
   * ダークテーマ判定
   * @returns {boolean} ダークテーマフラグ
   */
  static isDarkTheme() {
    // システム設定の確認
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
      return true;
    }

    // ページのテーマ検出
    const html = document.documentElement;
    const body = document.body;

    return html.classList.contains('dark') ||
      body.classList.contains('dark') ||
      html.getAttribute('data-theme') === 'dark' ||
      body.getAttribute('data-theme') === 'dark';
  }

  /**
   * テーマ変更の監視
   * @param {Function} callback - コールバック関数
   * @returns {Function} 監視解除関数
   */
  static watchThemeChange(callback) {
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');

    const handler = (e) => callback(e.matches);
    mediaQuery.addListener(handler);

    // 初回実行
    callback(mediaQuery.matches);

    // 監視解除関数を返す
    return () => mediaQuery.removeListener(handler);
  }
}

/**
 * URL検証ユーティリティ - セキュアなURL検証機能を提供
 * 
 * セキュリティ原則:
 * - ホスト名の完全一致チェック
 * - パスプレフィックスの厳密な検証
 * - URL偽装攻撃の防止
 * - 無効なURL形式の適切な処理
 */

/**
 * VSCodeドキュメントページの安全な検証
 * @param {string} url - 検証対象のURL
 * @returns {boolean} VSCodeドキュメントページかどうか
 */
function isVSCodeDocumentPage(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }

  try {
    const urlObj = new URL(url);

    // 許可されたホスト名（完全一致）
    const allowedHosts = [
      'code.visualstudio.com',
      'marketplace.visualstudio.com'
    ];

    // ホストとパスの組み合わせで許可するパターン
    const allowedHostsWithPath = [
      {
        host: 'docs.microsoft.com',
        pathPrefix: '/ja-jp/azure/developer/javascript/'
      },
      {
        host: 'docs.microsoft.com',
        pathPrefix: '/en-us/azure/developer/javascript/'
      }
    ];

    // 完全一致するホストをチェック
    if (allowedHosts.includes(urlObj.hostname)) {
      return true;
    }

    // ホストとパスの組み合わせをチェック
    for (const allowed of allowedHostsWithPath) {
      if (urlObj.hostname === allowed.host &&
        urlObj.pathname.startsWith(allowed.pathPrefix)) {
        return true;
      }
    }

    return false;

  } catch (error) {
    // 無効なURLの場合はfalseを返す
    console.warn('URL検証エラー - 無効なURL形式:', url, error.message);
    return false;
  }
}

/**
 * 一般的な開発者ドキュメントページの安全な検証
 * @param {string} url - 検証対象のURL
 * @returns {boolean} 開発者ドキュメントページかどうか
 */
function isDeveloperDocumentPage(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }

  try {
    const urlObj = new URL(url);

    // 許可されたホスト名（完全一致）
    const allowedHosts = [
      'developer.mozilla.org',
      'nodejs.org',
      'reactjs.org',
      'angular.io',
      'vuejs.org'
    ];

    // ホストとパスの組み合わせで許可するパターン
    const allowedHostsWithPath = [
      {
        host: 'docs.microsoft.com',
        pathPrefix: '/ja-jp/'
      },
      {
        host: 'docs.microsoft.com',
        pathPrefix: '/en-us/'
      },
      {
        host: 'github.com',
        pathPrefix: '/'
      }
    ];

    // 完全一致するホストをチェック
    if (allowedHosts.includes(urlObj.hostname)) {
      return true;
    }

    // ホストとパスの組み合わせをチェック
    for (const allowed of allowedHostsWithPath) {
      if (urlObj.hostname === allowed.host &&
        urlObj.pathname.startsWith(allowed.pathPrefix)) {
        return true;
      }
    }

    return false;

  } catch (error) {
    // 無効なURLの場合はfalseを返す
    console.warn('URL検証エラー - 無効なURL形式:', url, error.message);
    return false;
  }
}

/**
 * セキュアなURL検証設定の取得
 * @returns {Object} URL検証設定オブジェクト
 */
function getSecureUrlValidationConfig() {
  return {
    vscode: {
      allowedHosts: [
        'code.visualstudio.com',
        'marketplace.visualstudio.com'
      ],
      allowedHostsWithPath: [
        {
          host: 'docs.microsoft.com',
          pathPrefix: '/ja-jp/azure/developer/javascript/'
        },
        {
          host: 'docs.microsoft.com',
          pathPrefix: '/en-us/azure/developer/javascript/'
        }
      ]
    },
    developer: {
      allowedHosts: [
        'developer.mozilla.org',
        'nodejs.org',
        'reactjs.org',
        'angular.io',
        'vuejs.org'
      ],
      allowedHostsWithPath: [
        {
          host: 'docs.microsoft.com',
          pathPrefix: '/ja-jp/'
        },
        {
          host: 'docs.microsoft.com',
          pathPrefix: '/en-us/'
        },
        {
          host: 'github.com',
          pathPrefix: '/'
        }
      ]
    }
  };
}

/**
 * URL検証ログの記録
 * @param {string} url - 検証したURL 
 * @param {boolean} result - 検証結果
 * @param {string} type - 検証タイプ
 */
function logUrlValidation(url, result, type = 'unknown') {
  const logData = {
    timestamp: new Date().toISOString(),
    url: url,
    result: result,
    type: type,
    hostname: result ? new URL(url).hostname : null
  };

  console.log('URL検証ログ:', logData);
}

// Chrome Extension環境での使用を考慮したエクスポート
if (typeof chrome !== 'undefined' && chrome.runtime) {
  // Chrome Extension環境
  window.UrlValidator = {
    isVSCodeDocumentPage,
    isDeveloperDocumentPage,
    getSecureUrlValidationConfig,
    logUrlValidation
  };
} else if (typeof globalThis !== 'undefined') {
  // その他のグローバル環境
  globalThis.UrlValidator = {
    isVSCodeDocumentPage,
    isDeveloperDocumentPage,
    getSecureUrlValidationConfig,
    logUrlValidation
  };
}

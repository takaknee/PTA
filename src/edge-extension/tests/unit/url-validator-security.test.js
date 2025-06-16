/**
 * URL検証ユーティリティのセキュリティテスト
 * 
 * このテストファイルは、URL検証機能のセキュリティ脆弱性を確認するためのものです。
 * 悪意のあるURL偽装攻撃を模擬して、検証ロジックの安全性を確認します。
 */

// テスト用のURL検証ユーティリティを読み込み
// 実際のテスト実行時はurl-validator.jsを読み込む必要があります

/**
 * セキュリティテストケース
 */
const SECURITY_TEST_CASES = [
  // 正常なURL
  {
    url: 'https://code.visualstudio.com/docs',
    expected: true,
    description: '正常なVSCodeドキュメントURL'
  },
  {
    url: 'https://marketplace.visualstudio.com/items?itemName=ms-vscode.vscode-typescript-next',
    expected: true,
    description: '正常なVSCode Marketplace URL'
  },
  {
    url: 'https://docs.microsoft.com/ja-jp/azure/developer/javascript/tutorial-vscode-azure-app-service-node-01',
    expected: true,
    description: '正常なMicrosoft Azure JavaScript ドキュメントURL'
  },

  // 悪意のあるURL偽装攻撃
  {
    url: 'https://malicious.com/code.visualstudio.com/fake',
    expected: false,
    description: '悪意のあるURL偽装攻撃 - パスに正規ドメインを含む'
  },
  {
    url: 'https://code.visualstudio.com.malicious.com/',
    expected: false,
    description: '悪意のあるURL偽装攻撃 - サブドメイン偽装'
  },
  {
    url: 'https://fake-marketplace.visualstudio.com/',
    expected: false,
    description: '悪意のあるURL偽装攻撃 - 類似ドメイン'
  },
  {
    url: 'https://code-visualstudio.com/',
    expected: false,
    description: '悪意のあるURL偽装攻撃 - ハイフン置換'
  },
  {
    url: 'https://evil.com?redirect=code.visualstudio.com',
    expected: false,
    description: '悪意のあるURL偽装攻撃 - クエリパラメータに正規ドメイン'
  },
  {
    url: 'https://malicious.com#code.visualstudio.com',
    expected: false,
    description: '悪意のあるURL偽装攻撃 - フラグメントに正規ドメイン'
  },
  {
    url: 'https://docs.microsoft.com/fake-path/code.visualstudio.com',
    expected: false,
    description: '悪意のあるURL偽装攻撃 - 許可されたドメインの不正パス'
  },

  // 無効なURL形式
  {
    url: 'invalid-url',
    expected: false,
    description: '無効なURL形式'
  },
  {
    url: '',
    expected: false,
    description: '空文字'
  },
  {
    url: null,
    expected: false,
    description: 'null値'
  },
  {
    url: undefined,
    expected: false,
    description: 'undefined値'
  },

  // エッジケース
  {
    url: 'http://code.visualstudio.com/',
    expected: true,
    description: 'HTTPプロトコル（正常）'
  },
  {
    url: 'ftp://code.visualstudio.com/',
    expected: false,
    description: 'FTPプロトコル（無効）'
  },
  {
    url: 'https://CODE.VISUALSTUDIO.COM/',
    expected: false,
    description: '大文字ドメイン（注意：JavaScriptのURL.hostnameは小文字に変換される）'
  }
];

/**
 * セキュリティテストの実行
 * @param {Function} validationFunction - テスト対象のURL検証関数
 * @returns {Object} テスト結果
 */
function runSecurityTests(validationFunction) {
  const results = {
    passed: 0,
    failed: 0,
    total: SECURITY_TEST_CASES.length,
    failures: []
  };

  console.log('🔒 URL検証セキュリティテストを開始します...\n');

  SECURITY_TEST_CASES.forEach((testCase, index) => {
    try {
      const result = validationFunction(testCase.url);
      const passed = result === testCase.expected;

      if (passed) {
        results.passed++;
        console.log(`✅ Test ${index + 1}: ${testCase.description}`);
      } else {
        results.failed++;
        const failure = {
          testCase: testCase,
          actualResult: result,
          index: index + 1
        };
        results.failures.push(failure);
        console.error(`❌ Test ${index + 1}: ${testCase.description}`);
        console.error(`   Expected: ${testCase.expected}, Got: ${result}`);
        console.error(`   URL: ${testCase.url}`);
      }
    } catch (error) {
      results.failed++;
      const failure = {
        testCase: testCase,
        error: error.message,
        index: index + 1
      };
      results.failures.push(failure);
      console.error(`💥 Test ${index + 1}: ${testCase.description} - エラー: ${error.message}`);
    }
  });

  return results;
}

/**
 * テスト結果のサマリー表示
 * @param {Object} results - テスト結果
 */
function displayTestSummary(results) {
  console.log('\n📊 セキュリティテスト結果サマリー');
  console.log('='.repeat(50));
  console.log(`総テスト数: ${results.total}`);
  console.log(`成功: ${results.passed}`);
  console.log(`失敗: ${results.failed}`);
  console.log(`成功率: ${((results.passed / results.total) * 100).toFixed(1)}%`);

  if (results.failures.length > 0) {
    console.log('\n❌ 失敗したテスト:');
    results.failures.forEach(failure => {
      console.log(`  ${failure.index}. ${failure.testCase.description}`);
      if (failure.error) {
        console.log(`     エラー: ${failure.error}`);
      } else {
        console.log(`     期待値: ${failure.testCase.expected}, 実際: ${failure.actualResult}`);
      }
    });
  }

  if (results.failed === 0) {
    console.log('\n🎉 すべてのセキュリティテストが成功しました！');
  } else {
    console.log('\n⚠️  失敗したテストがあります。セキュリティ脆弱性を確認してください。');
  }
}

/**
 * セキュリティテストの実行（メイン関数）
 * 使用例：
 * runFullSecurityTest(window.UrlValidator.isVSCodeDocumentPage);
 */
function runFullSecurityTest(validationFunction) {
  if (typeof validationFunction !== 'function') {
    console.error('❌ 検証関数が提供されていません');
    return;
  }

  const results = runSecurityTests(validationFunction);
  displayTestSummary(results);

  return results;
}

// Chrome Extension環境での使用を考慮したエクスポート
if (typeof chrome !== 'undefined' && chrome.runtime) {
  // Chrome Extension環境
  window.SecurityTester = {
    runSecurityTests,
    displayTestSummary,
    runFullSecurityTest,
    SECURITY_TEST_CASES
  };
} else if (typeof globalThis !== 'undefined') {
  // その他のグローバル環境
  globalThis.SecurityTester = {
    runSecurityTests,
    displayTestSummary,
    runFullSecurityTest,
    SECURITY_TEST_CASES
  };
}

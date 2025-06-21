/*
 * コピーボタン機能テスト用スクリプト
 * content.jsから抜き出したコピー関数群
 */

/**
 * フォールバック用クリップボードコピー機能
 * @param {string} text - コピーするテキスト
 */
function fallbackCopyToClipboard(text) {
  try {
    // テキストエリアを作成してコピー
    const textArea = document.createElement('textarea');
    textArea.value = text;
    textArea.style.position = 'fixed';
    textArea.style.top = '0';
    textArea.style.left = '0';
    textArea.style.width = '2em';
    textArea.style.height = '2em';
    textArea.style.padding = '0';
    textArea.style.border = 'none';
    textArea.style.outline = 'none';
    textArea.style.boxShadow = 'none';
    textArea.style.background = 'transparent';

    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();

    document.execCommand('copy');
    document.body.removeChild(textArea);

    showNotification('クリップボードにコピーしました', 'success');
    console.log('フォールバック方式でコピーが完了しました');
  } catch (err) {
    console.error('フォールバックコピーも失敗しました:', err);
    alert('クリップボードへのコピーに失敗しました。手動でコピーしてください。');
  }
}

/**
 * AI結果の汎用コピー機能
 * @param {string} elementId - コピー対象要素のID（オプション）
 * @param {HTMLElement} buttonElement - クリックされたボタン要素（オプション）
 */
function copyAIResult(elementId, buttonElement) {
  try {
    let contentToCopy = '';

    if (elementId) {
      // 特定の要素からコンテンツを取得
      const element = document.getElementById(elementId);
      if (element) {
        contentToCopy = element.innerText || element.textContent;
      }
    } else if (buttonElement) {
      // ボタンの親要素から結果コンテンツを探す
      const resultContainer = buttonElement.closest('.ai-result-container');
      if (resultContainer) {
        // ヘッダーとページ情報を除いたメインコンテンツを取得
        const mainContent = resultContainer.querySelector('.ai-result-content, #vscode-analysis-content, .analysis-content, .vscode-analysis-content');
        if (mainContent) {
          contentToCopy = mainContent.innerText || mainContent.textContent;
        } else {
          // フォールバック: 結果コンテナ全体のテキスト
          contentToCopy = resultContainer.innerText || resultContainer.textContent;
        }
      }
    }

    if (!contentToCopy) {
      alert('コピーするコンテンツが見つかりませんでした。');
      return;
    }

    // クリップボードにコピー
    navigator.clipboard.writeText(contentToCopy).then(() => {
      console.log('AI結果をクリップボードにコピーしました');
      showNotification('AI解析結果をコピーしました', 'success');

      // ボタンのフィードバック
      if (buttonElement) {
        const originalText = buttonElement.textContent;
        buttonElement.textContent = '✅ コピー完了!';
        buttonElement.classList.add('copied');

        setTimeout(() => {
          buttonElement.textContent = originalText;
          buttonElement.classList.remove('copied');
        }, 2000);
      }
    }).catch(err => {
      console.error('AI結果のコピーに失敗しました:', err);
      fallbackCopyToClipboard(contentToCopy);
    });
  } catch (error) {
    console.error('AI結果コピー処理中にエラー:', error);
    alert('コピー処理中にエラーが発生しました。');
  }
}

/**
 * VSCode解析結果のコピー機能
 */
function copyVSCodeAnalysis() {
  try {
    const content = document.getElementById('vscode-analysis-content');
    if (!content) {
      alert('コピーするコンテンツが見つかりませんでした。');
      return;
    }

    const textContent = content.innerText || content.textContent;

    // クリップボードにコピー
    navigator.clipboard.writeText(textContent).then(() => {
      console.log('VSCode解析結果をクリップボードにコピーしました');
      showNotification('VSCode解析結果をコピーしました', 'success');

      // ボタンのフィードバック
      const button = document.querySelector('.ai-result-copy-btn');
      if (button) {
        const originalText = button.textContent;
        button.textContent = '✅ コピー完了!';
        button.classList.add('copied');

        setTimeout(() => {
          button.textContent = originalText;
          button.classList.remove('copied');
        }, 2000);
      }
    }).catch(err => {
      console.error('VSCode解析結果のコピーに失敗しました:', err);
      fallbackCopyToClipboard(textContent);
    });
  } catch (error) {
    console.error('VSCode解析結果コピー処理中にエラー:', error);
    alert('コピー処理中にエラーが発生しました。');
  }
}

/**
 * 一般的なコピーボタンイベントハンドラ
 * @param {Event} event - クリックイベント
 */
function handleCopyButtonClick(event) {
  event.preventDefault();
  event.stopPropagation();

  const button = event.target;
  const targetId = button.getAttribute('data-copy-target');

  if (targetId) {
    // 特定のIDからコピー
    copyAIResult(targetId, button);
  } else {
    // ボタンの親要素から自動検出
    copyAIResult(null, button);
  }
}

/**
 * VSCode設定ファイル（settings.json）のコピー機能
 * @param {HTMLElement} buttonElement - クリックされたコピーボタン
 */
function copySettingsJSON(buttonElement) {
  try {
    // 設定JSONの親コンテナを取得
    const container = buttonElement.closest('.settings-json-container');
    if (!container) {
      console.error('設定JSONコンテナが見つかりません');
      alert('コピーする設定ファイルが見つかりませんでした。');
      return;
    }

    // JSONコードブロックを取得
    const codeElement = container.querySelector('.settings-json code');
    if (!codeElement) {
      console.error('設定JSONコードが見つかりません');
      alert('コピーする設定ファイルが見つかりませんでした。');
      return;
    }

    // JSONコンテンツを取得（コメントを含む）
    const jsonContent = codeElement.textContent || codeElement.innerText;

    // クリップボードにコピー
    navigator.clipboard.writeText(jsonContent).then(() => {
      console.log('VSCode設定ファイルをクリップボードにコピーしました');
      showNotification('設定ファイル（settings.json）をコピーしました', 'success');

      // ボタンのフィードバック
      const originalText = buttonElement.textContent;
      buttonElement.textContent = '✅ コピー完了!';
      buttonElement.classList.add('copied');

      setTimeout(() => {
        buttonElement.textContent = originalText;
        buttonElement.classList.remove('copied');
      }, 2000);
    }).catch(err => {
      console.error('VSCode設定ファイルのコピーに失敗しました:', err);

      // フォールバック: 従来のコピー方法
      fallbackCopyToClipboard(jsonContent);

      // ボタンのフィードバック（フォールバック成功時）
      const originalText = buttonElement.textContent;
      buttonElement.textContent = '✅ コピー完了!';
      buttonElement.classList.add('copied');

      setTimeout(() => {
        buttonElement.textContent = originalText;
        buttonElement.classList.remove('copied');
      }, 2000);
    });
  } catch (error) {
    console.error('VSCode設定ファイルコピー処理中にエラー:', error);
    alert('設定ファイルのコピー処理中にエラーが発生しました。');
  }
}

// グローバル関数として公開
window.copyAIResult = copyAIResult;
window.copyVSCodeAnalysis = copyVSCodeAnalysis;
window.copySettingsJSON = copySettingsJSON;
window.handleCopyButtonClick = handleCopyButtonClick;

console.log('コピーボタン機能が読み込まれました');

/*
 * Shima Edge拡張機能 - 設定スクリプト
 * Copyright (c) 2024 Shima Development Team
 */

// DOM読み込み完了時の初期化
document.addEventListener('DOMContentLoaded', function () {
    initializeOptions();
});

/**
 * 設定画面の初期化
 */
function initializeOptions() {
    // イベントリスナーの設定
    document.getElementById('provider').addEventListener('change', handleProviderChange);
    document.getElementById('model-select').addEventListener('change', handleModelChange);
    document.getElementById('test-connection').addEventListener('click', testConnection);
    document.getElementById('save-settings').addEventListener('click', saveSettings);
    document.getElementById('reset-settings').addEventListener('click', resetSettings);
    document.getElementById('export-settings').addEventListener('click', exportSettings);
    document.getElementById('import-settings').addEventListener('click', importSettings);
    document.getElementById('export-history').addEventListener('click', exportHistory);
    document.getElementById('clear-all-data').addEventListener('click', clearAllData);
    document.getElementById('import-file').addEventListener('change', handleImportFile);

    // Azure エンドポイントのリアルタイムフィードバック
    const azureEndpointInput = document.getElementById('azure-endpoint');
    if (azureEndpointInput) {
        azureEndpointInput.addEventListener('input', function () {
            handleEndpointInput(this);
        });
        azureEndpointInput.addEventListener('blur', function () {
            validateAzureEndpoint(this.value);
        });
        azureEndpointInput.addEventListener('focus', function () {
            this.classList.add('typing');
        });
    }

    // APIキーのリアルタイムフィードバック
    const apiKeyInput = document.getElementById('api-key');
    if (apiKeyInput) {
        apiKeyInput.addEventListener('input', function () {
            handleApiKeyInput(this);
        });
        apiKeyInput.addEventListener('focus', function () {
            this.classList.add('typing');
        });
        apiKeyInput.addEventListener('blur', function () {
            validateApiKey(this.value);
        });
    }

    // カスタムモデル入力のフィードバック
    const customModelInput = document.getElementById('custom-model');
    if (customModelInput) {
        customModelInput.addEventListener('input', function () {
            handleCustomModelInput(this);
        });
    }

    // 現在の設定を読み込み
    loadSettings();
    loadStatistics();
}

/**
 * プロバイダー変更時の処理
 */
function handleProviderChange() {
    const provider = document.getElementById('provider').value;
    const azureSettings = document.getElementById('azure-settings');

    if (provider === 'azure') {
        azureSettings.style.display = 'block';
    } else {
        azureSettings.style.display = 'none';
    }
}

/**
 * モデル選択変更時の処理
 */
function handleModelChange() {
    const modelSelect = document.getElementById('model-select');
    const customModelInput = document.getElementById('custom-model-input');

    if (modelSelect.value === 'custom') {
        customModelInput.style.display = 'block';
        document.getElementById('custom-model').focus();
    } else {
        customModelInput.style.display = 'none';
    }
}

/**
 * 設定を読み込み
 */
function loadSettings() {
    chrome.storage.local.get(['ai_settings'], (result) => {
        const settings = result.ai_settings || {
            provider: 'azure',
            model: 'gpt-4o-mini',
            customModel: '',
            apiKey: '',
            azureEndpoint: '',
            autoDetect: true,
            showNotifications: true,
            saveHistory: true
        };

        // UI要素に設定値を反映
        document.getElementById('provider').value = settings.provider || 'azure';
        document.getElementById('api-key').value = settings.apiKey || '';
        document.getElementById('azure-endpoint').value = settings.azureEndpoint || '';
        document.getElementById('auto-detect').checked = settings.autoDetect !== false;
        document.getElementById('show-notifications').checked = settings.showNotifications !== false;
        document.getElementById('save-history').checked = settings.saveHistory !== false;

        // モデル設定の復元
        const modelSelect = document.getElementById('model-select');
        const customModelInput = document.getElementById('custom-model');

        if (settings.customModel && settings.customModel.trim() !== '') {
            // カスタムモデルが設定されている場合
            modelSelect.value = 'custom';
            customModelInput.value = settings.customModel;
            document.getElementById('custom-model-input').style.display = 'block';
        } else {
            // 標準モデルの場合
            const modelValue = settings.model || 'gpt-4o-mini';
            // 選択肢に存在するかチェック
            const optionExists = Array.from(modelSelect.options).some(option => option.value === modelValue);

            if (optionExists) {
                modelSelect.value = modelValue;
            } else {
                // 存在しない場合はカスタムモデルとして設定
                modelSelect.value = 'custom';
                customModelInput.value = modelValue;
                document.getElementById('custom-model-input').style.display = 'block';
            }
        }

        // プロバイダー設定の表示制御
        handleProviderChange();

        // 入力値の初期バリデーション
        if (settings.azureEndpoint) {
            validateAzureEndpoint(settings.azureEndpoint);
        }
        if (settings.apiKey) {
            validateApiKey(settings.apiKey);
        }
    });
}

/**
 * 統計情報を読み込み
 */
function loadStatistics() {
    chrome.storage.local.get(['ai_history', 'ai_statistics'], (result) => {
        const history = result.ai_history || [];

        // 解析・作成回数の集計
        const analyses = history.filter(item => item.type === 'analysis').length;
        const compositions = history.filter(item => item.type === 'composition').length;

        document.getElementById('total-analyses').textContent = analyses;
        document.getElementById('total-compositions').textContent = compositions;
        document.getElementById('history-count').textContent = history.length;
    });
}

/**
 * 接続テスト
 */
async function testConnection() {
    const settings = getCurrentSettings();
    const testResult = document.getElementById('test-result');
    const testButton = document.getElementById('test-connection');

    if (!settings.apiKey) {
        showTestResult('APIキーを入力してください', 'error');
        return;
    } if (settings.provider === 'azure' && !settings.azureEndpoint) {
        showTestResult('Azureエンドポイントを入力してください', 'error');
        return;
    }

    // Azure エンドポイントの詳細バリデーション
    if (settings.provider === 'azure') {
        const isValidEndpoint = validateAzureEndpoint(settings.azureEndpoint);
        if (!isValidEndpoint) {
            showTestResult('Azure エンドポイントの形式が正しくありません。正しい形式: https://your-resource-name.openai.azure.com', 'error');
            return;
        }
    }

    testButton.disabled = true;
    testButton.textContent = '🔄 接続テスト中...';
    testResult.style.display = 'none';

    // 詳細なテストログ用の要素を作成
    const detailsElement = document.getElementById('test-details') || createTestDetailsElement();
    detailsElement.innerHTML = '';
    detailsElement.style.display = 'block';

    try {
        // テスト開始のログ
        appendTestLog(detailsElement, '🔄 接続テストを開始します...', 'info');
        appendTestLog(detailsElement, `プロバイダー: ${settings.provider}`, 'info');
        appendTestLog(detailsElement, `モデル: ${settings.model}`, 'info');

        if (settings.provider === 'azure') {
            const endpointUrl = new URL(settings.azureEndpoint);
            appendTestLog(detailsElement, `エンドポイント: ${endpointUrl.hostname}`, 'info');
        }

        // バックグラウンドスクリプトに接続テストを依頼
        const response = await new Promise((resolve, reject) => {
            const timeout = setTimeout(() => {
                reject(new Error('接続テストがタイムアウトしました（30秒）'));
            }, 30000);

            chrome.runtime.sendMessage({
                action: 'testApiConnection',
                data: settings
            }, (response) => {
                clearTimeout(timeout);

                if (chrome.runtime.lastError) {
                    reject(new Error('拡張機能の通信エラー: ' + chrome.runtime.lastError.message));
                    return;
                }

                resolve(response);
            });
        });

        testButton.disabled = false;
        testButton.textContent = '🔧 接続テスト'; if (response.success) {
            appendTestLog(detailsElement, '✅ 接続テストが成功しました', 'success');
            appendTestLog(detailsElement, `応答: ${response.result || '正常に応答を受信'}`, 'success');

            // 診断情報を表示
            if (response.diagnostics) {
                displayDiagnostics(response.diagnostics, detailsElement);
            }

            showTestResult('✅ 接続テストが成功しました', 'success');
        } else {
            appendTestLog(detailsElement, '❌ 接続テストに失敗しました', 'error');
            appendTestLog(detailsElement, `エラー: ${response.error}`, 'error');

            // 診断情報を表示
            if (response.diagnostics) {
                displayDiagnostics(response.diagnostics, detailsElement);
            }

            // エラーの種類に応じた対策情報を追加
            addTroubleshootingInfo(detailsElement, response.error);

            showTestResult(`❌ 接続テストに失敗しました: ${response.error}`, 'error');
        }

    } catch (error) {
        testButton.disabled = false;
        testButton.textContent = '🔧 接続テスト';

        appendTestLog(detailsElement, '❌ テスト実行中にエラーが発生しました', 'error');
        appendTestLog(detailsElement, `エラー: ${error.message}`, 'error');

        showTestResult(`❌ テスト実行エラー: ${error.message}`, 'error');
    }
}

/**
 * テスト詳細要素を作成
 */
function createTestDetailsElement() {
    const existingDetails = document.getElementById('test-details');
    if (existingDetails) {
        return existingDetails;
    }

    const detailsElement = document.createElement('div');
    detailsElement.id = 'test-details';
    detailsElement.className = 'test-details';
    detailsElement.style.cssText = `
        margin-top: 15px;
        padding: 15px;
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        font-family: 'Consolas', 'Monaco', monospace;
        font-size: 12px;
        line-height: 1.4;
        max-height: 300px;
        overflow-y: auto;
        display: none;
    `;

    const testButton = document.getElementById('test-connection');
    testButton.parentNode.insertBefore(detailsElement, testButton.nextSibling);

    return detailsElement;
}

/**
 * テストログを追加
 */
function appendTestLog(container, message, type) {
    const logElement = document.createElement('div');
    logElement.style.cssText = `
        margin: 3px 0;
        padding: 2px 0;
        color: ${getLogColor(type)};
    `;
    logElement.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
    container.appendChild(logElement);
    container.scrollTop = container.scrollHeight;
}

/**
 * ログタイプに応じた色を取得
 */
function getLogColor(type) {
    switch (type) {
        case 'success': return '#28a745';
        case 'error': return '#dc3545';
        case 'warning': return '#ffc107';
        case 'info':
        default: return '#6c757d';
    }
}

/**
 * トラブルシューティング情報を追加
 */
function addTroubleshootingInfo(container, error) {
    const troubleshootingDiv = document.createElement('div');
    troubleshootingDiv.style.cssText = `
        margin-top: 10px;
        padding: 10px;
        background: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 6px;
        color: #856404;
    `;

    let troubleshootingInfo = '💡 トラブルシューティング:\n';

    if (error.includes('Failed to fetch') || error.includes('ネットワーク')) {
        troubleshootingInfo += `
• インターネット接続を確認してください
• プロキシ設定を確認してください
• ファイアウォールやセキュリティソフトを確認してください
• VPN接続を確認してください
• 拡張機能を無効化→有効化してみてください
• ブラウザを再起動してみてください`;
    } else if (error.includes('401') || error.includes('APIキー')) {
        troubleshootingInfo += `
• APIキーが正しく入力されているか確認してください
• APIキーが有効期限内か確認してください
• APIキーの権限設定を確認してください`;
    } else if (error.includes('404') || error.includes('エンドポイント')) {
        troubleshootingInfo += `
• Azure OpenAIの場合、エンドポイントURLを確認してください
• デプロイメント名が正しいか確認してください
• モデル名が正しいか確認してください`;
    } else if (error.includes('CORS')) {
        troubleshootingInfo += `
• CORS制限の問題が発生しています
• Offscreen documentが正しく動作していない可能性があります
• 拡張機能を再インストールしてみてください`;
    } else {
        troubleshootingInfo += `
• 設定値を再確認してください
• しばらく時間をおいて再試行してください
• 開発者ツールのコンソールを確認してください`;
    }

    troubleshootingDiv.textContent = troubleshootingInfo;
    container.appendChild(troubleshootingDiv);
}

/**
 * 診断情報表示
 */
function displayDiagnostics(diagnostics, container) {
    if (!diagnostics) return;

    const diagnosticsDiv = document.createElement('div');
    diagnosticsDiv.style.cssText = `
        margin-top: 15px;
        padding: 15px;
        background: #e3f2fd;
        border: 1px solid #2196f3;
        border-radius: 6px;
        font-size: 12px;
        color: #1565c0;
    `;

    let diagnosticsInfo = '🔍 システム診断結果:\n\n';

    // Offscreen Document状態
    diagnosticsInfo += `📄 Offscreen Document:\n`;
    diagnosticsInfo += `• 対応: ${diagnostics.offscreenDocument.canCreate ? '✅' : '❌'}\n`;
    diagnosticsInfo += `• 存在: ${diagnostics.offscreenDocument.exists ? '✅' : '❌'}\n`;
    if (diagnostics.offscreenDocument.error) {
        diagnosticsInfo += `• エラー: ${diagnostics.offscreenDocument.error}\n`;
    }

    // ネットワーク状態
    diagnosticsInfo += `\n🌐 ネットワーク接続:\n`;
    diagnosticsInfo += `• 基本接続: ${diagnostics.network.basicConnectivity ? '✅' : '❌'}\n`;
    diagnosticsInfo += `• OpenAI到達: ${diagnostics.network.openaiReachable ? '✅' : '❌'}\n`;

    // Chrome機能
    diagnosticsInfo += `\n🔧 Chrome機能:\n`;
    diagnosticsInfo += `• Offscreen対応: ${diagnostics.chrome.offscreenSupport ? '✅' : '❌'}\n`;
    diagnosticsInfo += `• Runtime対応: ${diagnostics.chrome.runtimeSupport ? '✅' : '❌'}\n`;

    // 権限
    if (diagnostics.permissions && diagnostics.permissions.permissions) {
        diagnosticsInfo += `\n🔐 権限: ${diagnostics.permissions.permissions.join(', ')}\n`;
    }

    diagnosticsInfo += `\n⏰ 診断時刻: ${new Date(diagnostics.timestamp).toLocaleString()}`;

    diagnosticsDiv.textContent = diagnosticsInfo;
    container.appendChild(diagnosticsDiv);
}

/**
 * テスト結果表示
 */
function showTestResult(message, type) {
    const testResult = document.getElementById('test-result');
    testResult.textContent = message;
    testResult.className = `test-result ${type}`;
    testResult.style.display = 'block';

    // 5秒後に自動で非表示
    setTimeout(() => {
        testResult.style.display = 'none';
    }, 5000);
}

/**
 * 現在の設定を取得（改善版）
 */
function getCurrentSettings() {
    const modelSelect = document.getElementById('model-select');
    const customModel = document.getElementById('custom-model');

    let finalModel;
    let customModelValue = '';

    if (modelSelect.value === 'custom') {
        finalModel = customModel.value.trim();
        customModelValue = finalModel;
    } else {
        finalModel = modelSelect.value;
        customModelValue = '';
    }

    return {
        provider: document.getElementById('provider').value,
        model: finalModel,
        customModel: customModelValue,
        apiKey: document.getElementById('api-key').value.trim(),
        azureEndpoint: document.getElementById('azure-endpoint').value.trim(),
        autoDetect: document.getElementById('auto-detect').checked,
        showNotifications: document.getElementById('show-notifications').checked,
        saveHistory: document.getElementById('save-history').checked
    };
}

/**
 * 設定を保存
 */
/**
 * 設定保存
 */
function saveSettings() {
    const settings = getCurrentSettings();

    // 必須項目の検証
    if (!settings.apiKey) {
        showNotification('APIキーを入力してください', 'error');
        document.getElementById('api-key').focus();
        return;
    }

    if (settings.provider === 'azure' && !settings.azureEndpoint) {
        showNotification('Azureエンドポイントを入力してください', 'error');
        document.getElementById('azure-endpoint').focus();
        return;
    }

    // Azure エンドポイントの詳細バリデーション
    if (settings.provider === 'azure') {
        const isValidEndpoint = validateAzureEndpoint(settings.azureEndpoint);
        if (!isValidEndpoint) {
            showNotification('Azure エンドポイントの形式が正しくありません。正しい形式で入力してください。', 'error');
            document.getElementById('azure-endpoint').focus();
            return;
        }
    }

    // APIキーの詳細バリデーション
    const isValidApiKey = validateApiKey(settings.apiKey);
    if (!isValidApiKey) {
        showNotification('APIキーの形式を確認してください。', 'error');
        document.getElementById('api-key').focus();
        return;
    }

    // モデル名の検証
    if (!settings.model || settings.model.trim() === '') {
        showNotification('モデル名を指定してください。', 'error');
        return;
    }

    // 保存実行
    chrome.storage.local.set({ 'ai_settings': settings }, () => {
        showNotification('✅ 設定を保存しました', 'success');
    });
}

/**
 * 設定をリセット
 */
function resetSettings() {
    showConfirmDialog('設定をリセットしますか？', () => {
        const defaultSettings = {
            provider: 'azure',
            model: 'gpt-4o-mini',
            apiKey: '',
            azureEndpoint: '',
            autoDetect: true,
            showNotifications: true,
            saveHistory: true
        }; chrome.storage.local.set({ 'ai_settings': defaultSettings }, () => {
            loadSettings();
            showNotification('設定をリセットしました', 'info');
        });
    });
}

/**
 * 設定をエクスポート
 */
function exportSettings() {
    chrome.storage.local.get(['ai_settings'], (result) => {
        const settings = result.ai_settings || {};

        // APIキーを除外
        const exportData = { ...settings };
        delete exportData.apiKey;

        const blob = new Blob([JSON.stringify(exportData, null, 2)], {
            type: 'application/json'
        });

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'ai-settings.json';
        a.click();

        URL.revokeObjectURL(url);
        showNotification('設定をエクスポートしました', 'success');
    });
}

/**
 * 設定をインポート
 */
function importSettings() {
    document.getElementById('import-file').click();
}

/**
 * インポートファイル処理
 */
function handleImportFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const importedSettings = JSON.parse(e.target.result);            // 現在の設定とマージ（APIキーは保持）
            chrome.storage.local.get(['ai_settings'], (result) => {
                const currentSettings = result.ai_settings || {};
                const mergedSettings = {
                    ...importedSettings,
                    apiKey: currentSettings.apiKey // APIキーは保持
                };

                chrome.storage.local.set({ 'ai_settings': mergedSettings }, () => {
                    loadSettings();
                    showNotification('設定をインポートしました', 'success');
                });
            });

        } catch (error) {
            showNotification('設定ファイルの形式が正しくありません', 'error');
        }
    };

    reader.readAsText(file);

    // ファイル選択をリセット
    event.target.value = '';
}

/**
 * 履歴をエクスポート
 */
function exportHistory() {
    chrome.storage.local.get(['ai_history'], (result) => {
        const history = result.ai_history || [];

        if (history.length === 0) {
            showNotification('エクスポートする履歴がありません', 'info');
            return;
        }

        const blob = new Blob([JSON.stringify(history, null, 2)], {
            type: 'application/json'
        });

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'ai-history.json';
        a.click();

        URL.revokeObjectURL(url);
        showNotification('履歴をエクスポートしました', 'success');
    });
}

/**
 * 全データを削除
 */
function clearAllData() {
    showConfirmDialog('すべてのデータ（設定・履歴）を削除しますか？\nこの操作は元に戻せません。', () => {
        chrome.storage.local.clear(() => {
            loadSettings();
            loadStatistics();
            showNotification('すべてのデータを削除しました', 'info');
        });
    });
}

/**
 * 確認ダイアログ表示
 */
function showConfirmDialog(message, callback) {
    const modal = document.getElementById('confirm-modal');
    const messageElement = document.getElementById('confirm-message');
    const okButton = document.getElementById('confirm-ok');

    messageElement.textContent = message;
    modal.style.display = 'flex';

    // 既存のイベントリスナーを削除
    okButton.replaceWith(okButton.cloneNode(true));
    const newOkButton = document.getElementById('confirm-ok');

    newOkButton.addEventListener('click', () => {
        closeConfirmModal();
        callback();
    });
}

/**
 * パスワード表示切り替え
 */
function togglePassword() {
    const passwordField = document.getElementById('api-key');
    const toggleButton = document.querySelector('.toggle-password');

    if (passwordField.type === 'password') {
        passwordField.type = 'text';
        toggleButton.textContent = '🙈';
    } else {
        passwordField.type = 'password';
        toggleButton.textContent = '👁️';
    }
}

/**
 * 確認ダイアログを閉じる
 */
function closeConfirmModal() {
    document.getElementById('confirm-modal').style.display = 'none';
}

// ...existing code...

/**
 * Azure エンドポイント入力時の処理
 */
function handleEndpointInput(input) {
    const statusElement = document.getElementById('endpoint-status');
    const value = input.value.trim();

    // 入力状態のスタイル更新
    input.classList.remove('valid', 'invalid');
    input.classList.add('typing');

    if (value === '') {
        updateInputStatus(statusElement, 'empty', '⚪', '未入力');
        input.classList.remove('typing');
        return;
    }

    // リアルタイムでの簡易バリデーション
    if (value.startsWith('https://') && value.includes('.openai.azure.com')) {
        updateInputStatus(statusElement, 'typing', '🔄', '入力中...');
    } else {
        updateInputStatus(statusElement, 'typing', '⚠️', '形式を確認中...');
    }

    // デバウンス処理で本格的なバリデーションを実行
    clearTimeout(input.validationTimeout);
    input.validationTimeout = setTimeout(() => {
        validateAzureEndpoint(value, statusElement, input);
    }, 1000);
}

/**
 * APIキー入力時の処理
 */
function handleApiKeyInput(input) {
    const statusElement = document.getElementById('apikey-status');
    const value = input.value.trim();

    input.classList.remove('valid', 'invalid');
    input.classList.add('typing');

    if (value === '') {
        updateInputStatus(statusElement, 'empty', '⚪', '未入力');
        input.classList.remove('typing');
        return;
    }

    // APIキーの基本的な形式チェック
    if (value.length < 10) {
        updateInputStatus(statusElement, 'typing', '⚠️', '短すぎる可能性があります...');
    } else if (value.length > 100) {
        updateInputStatus(statusElement, 'invalid', '❌', 'APIキーが長すぎます');
        input.classList.remove('typing');
        input.classList.add('invalid');
    } else {
        updateInputStatus(statusElement, 'typing', '🔄', 'APIキー確認中...');

        // デバウンス処理
        clearTimeout(input.validationTimeout);
        input.validationTimeout = setTimeout(() => {
            validateApiKey(value, statusElement, input);
        }, 1000);
    }
}

/**
 * カスタムモデル入力時の処理
 */
function handleCustomModelInput(input) {
    const value = input.value.trim();

    input.classList.remove('valid', 'invalid');

    if (value === '') {
        input.classList.remove('valid', 'invalid');
        return;
    }

    // モデル名の基本的な形式チェック
    const modelNamePattern = /^[a-zA-Z0-9\-_.]+$/;
    if (modelNamePattern.test(value)) {
        input.classList.add('valid');
    } else {
        input.classList.add('invalid');
    }
}

/**
 * 入力ステータスの更新
 */
function updateInputStatus(statusElement, statusClass, icon, text) {
    statusElement.className = `input-status ${statusClass}`;
    statusElement.querySelector('.status-icon').textContent = icon;
    statusElement.querySelector('.status-text').textContent = text;
}

/**
 * Azure エンドポイントのバリデーション（改善版）
 */
function validateAzureEndpoint(url, statusElement = null, inputElement = null) {
    if (!statusElement) {
        statusElement = document.getElementById('endpoint-status');
    }
    if (!inputElement) {
        inputElement = document.getElementById('azure-endpoint');
    }

    inputElement.classList.remove('typing');

    if (!url || url.trim() === '') {
        updateInputStatus(statusElement, 'empty', '⚪', '未入力');
        inputElement.classList.remove('valid', 'invalid');
        return false;
    }

    try {
        const urlObj = new URL(url);

        // Azure OpenAI の基本的な URL 形式チェック
        if (urlObj.protocol !== 'https:') {
            updateInputStatus(statusElement, 'invalid', '❌', 'HTTPSが必要です');
            inputElement.classList.add('invalid');
            return false;
        }

        if (!urlObj.hostname.includes('.openai.azure.com')) {
            updateInputStatus(statusElement, 'invalid', '❌', 'Azure OpenAI のエンドポイントではありません');
            inputElement.classList.add('invalid');
            return false;
        }

        // 有効な形式
        updateInputStatus(statusElement, 'valid', '✅', '有効なエンドポイントです');
        inputElement.classList.add('valid');
        return true;

    } catch (error) {
        updateInputStatus(statusElement, 'invalid', '❌', '無効なURL形式です');
        inputElement.classList.add('invalid');
        return false;
    }
}

/**
 * APIキーのバリデーション
 */
function validateApiKey(apiKey, statusElement = null, inputElement = null) {
    if (!statusElement) {
        statusElement = document.getElementById('apikey-status');
    }
    if (!inputElement) {
        inputElement = document.getElementById('api-key');
    }

    inputElement.classList.remove('typing');

    if (!apiKey || apiKey.trim() === '') {
        updateInputStatus(statusElement, 'empty', '⚪', '未入力');
        inputElement.classList.remove('valid', 'invalid');
        return false;
    }

    // APIキーの基本的な検証
    const trimmedKey = apiKey.trim();

    if (trimmedKey.length < 20) {
        updateInputStatus(statusElement, 'invalid', '❌', 'APIキーが短すぎます');
        inputElement.classList.add('invalid');
        return false;
    }

    if (trimmedKey.length > 200) {
        updateInputStatus(statusElement, 'invalid', '❌', 'APIキーが長すぎます');
        inputElement.classList.add('invalid');
        return false;
    }

    // 有効そうなAPIキー
    updateInputStatus(statusElement, 'valid', '✅', 'APIキーが設定されました');
    inputElement.classList.add('valid');
    return true;
}
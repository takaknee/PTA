/*
 * PTA Edge拡張機能 - 設定スクリプト
 * Copyright (c) 2024 PTA Development Team
 */

// DOM読み込み完了時の初期化
document.addEventListener('DOMContentLoaded', function() {
    initializeOptions();
});

/**
 * 設定画面の初期化
 */
function initializeOptions() {
    // イベントリスナーの設定
    document.getElementById('provider').addEventListener('change', handleProviderChange);
    document.getElementById('test-connection').addEventListener('click', testConnection);
    document.getElementById('save-settings').addEventListener('click', saveSettings);
    document.getElementById('reset-settings').addEventListener('click', resetSettings);
    document.getElementById('export-settings').addEventListener('click', exportSettings);
    document.getElementById('import-settings').addEventListener('click', importSettings);
    document.getElementById('export-history').addEventListener('click', exportHistory);
    document.getElementById('clear-all-data').addEventListener('click', clearAllData);
    document.getElementById('import-file').addEventListener('change', handleImportFile);
    
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
 * 設定を読み込み
 */
function loadSettings() {
    chrome.storage.local.get(['pta_settings'], (result) => {
        const settings = result.pta_settings || {
            provider: 'azure',
            model: 'gpt-4',
            apiKey: '',
            azureEndpoint: '',
            autoDetect: true,
            showNotifications: true,
            saveHistory: true
        };
        
        // UI要素に設定値を反映
        document.getElementById('provider').value = settings.provider || 'azure';
        document.getElementById('model').value = settings.model || 'gpt-4';
        document.getElementById('api-key').value = settings.apiKey || '';
        document.getElementById('azure-endpoint').value = settings.azureEndpoint || '';
        document.getElementById('auto-detect').checked = settings.autoDetect !== false;
        document.getElementById('show-notifications').checked = settings.showNotifications !== false;
        document.getElementById('save-history').checked = settings.saveHistory !== false;
        
        // プロバイダー設定の表示制御
        handleProviderChange();
    });
}

/**
 * 統計情報を読み込み
 */
function loadStatistics() {
    chrome.storage.local.get(['pta_history', 'pta_statistics'], (result) => {
        const history = result.pta_history || [];
        
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
function testConnection() {
    const settings = getCurrentSettings();
    const testResult = document.getElementById('test-result');
    const testButton = document.getElementById('test-connection');
    
    if (!settings.apiKey) {
        showTestResult('APIキーを入力してください', 'error');
        return;
    }
    
    if (settings.provider === 'azure' && !settings.azureEndpoint) {
        showTestResult('Azureエンドポイントを入力してください', 'error');
        return;
    }
    
    testButton.disabled = true;
    testButton.textContent = '🔄 テスト中...';
    testResult.style.display = 'none';
    
    // バックグラウンドスクリプトに接続テストを依頼
    chrome.runtime.sendMessage({
        action: 'testApiConnection',
        data: settings
    }, (response) => {
        testButton.disabled = false;
        testButton.textContent = '🔧 接続テスト';
        
        if (response.success) {
            showTestResult('✅ 接続テストが成功しました', 'success');
        } else {
            showTestResult(`❌ 接続テストに失敗しました: ${response.error}`, 'error');
        }
    });
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
 * 現在の設定を取得
 */
function getCurrentSettings() {
    return {
        provider: document.getElementById('provider').value,
        model: document.getElementById('model').value,
        apiKey: document.getElementById('api-key').value,
        azureEndpoint: document.getElementById('azure-endpoint').value,
        autoDetect: document.getElementById('auto-detect').checked,
        showNotifications: document.getElementById('show-notifications').checked,
        saveHistory: document.getElementById('save-history').checked
    };
}

/**
 * 設定を保存
 */
function saveSettings() {
    const settings = getCurrentSettings();
    
    // 必須項目の検証
    if (!settings.apiKey) {
        showNotification('APIキーを入力してください', 'error');
        return;
    }
    
    if (settings.provider === 'azure' && !settings.azureEndpoint) {
        showNotification('Azureエンドポイントを入力してください', 'error');
        return;
    }
    
    // 保存実行
    chrome.storage.local.set({ 'pta_settings': settings }, () => {
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
            model: 'gpt-4',
            apiKey: '',
            azureEndpoint: '',
            autoDetect: true,
            showNotifications: true,
            saveHistory: true
        };
        
        chrome.storage.local.set({ 'pta_settings': defaultSettings }, () => {
            loadSettings();
            showNotification('設定をリセットしました', 'info');
        });
    });
}

/**
 * 設定をエクスポート
 */
function exportSettings() {
    chrome.storage.local.get(['pta_settings'], (result) => {
        const settings = result.pta_settings || {};
        
        // APIキーを除外
        const exportData = { ...settings };
        delete exportData.apiKey;
        
        const blob = new Blob([JSON.stringify(exportData, null, 2)], {
            type: 'application/json'
        });
        
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'pta-settings.json';
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
    reader.onload = function(e) {
        try {
            const importedSettings = JSON.parse(e.target.result);
            
            // 現在の設定とマージ（APIキーは保持）
            chrome.storage.local.get(['pta_settings'], (result) => {
                const currentSettings = result.pta_settings || {};
                const mergedSettings = {
                    ...importedSettings,
                    apiKey: currentSettings.apiKey // APIキーは保持
                };
                
                chrome.storage.local.set({ 'pta_settings': mergedSettings }, () => {
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
    chrome.storage.local.get(['pta_history'], (result) => {
        const history = result.pta_history || [];
        
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
        a.download = 'pta-history.json';
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
 * 確認ダイアログを閉じる
 */
function closeConfirmModal() {
    document.getElementById('confirm-modal').style.display = 'none';
}

/**
 * 通知表示
 */
function showNotification(message, type = 'info') {
    const notification = document.getElementById('notification');
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
    
    // 3秒後に自動で非表示
    setTimeout(() => {
        notification.style.display = 'none';
    }, 3000);
}

/**
 * パスワード表示切り替え
 */
function togglePassword() {
    const apiKeyInput = document.getElementById('api-key');
    const toggleButton = document.querySelector('.toggle-password');
    
    if (apiKeyInput.type === 'password') {
        apiKeyInput.type = 'text';
        toggleButton.textContent = '🙈';
    } else {
        apiKeyInput.type = 'password';
        toggleButton.textContent = '👁️';
    }
}

// グローバル関数として公開
window.closeConfirmModal = closeConfirmModal;
window.togglePassword = togglePassword;
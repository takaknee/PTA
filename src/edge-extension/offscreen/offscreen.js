/*
 * Shima Edge拡張機能 - Offscreen API処理
 * Copyright (c) 2024 Shima Development Team
 * 
 * Service WorkerからのCORS制限を回避するためのOffscreen Document
 */

// メッセージリスナー
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log('Offscreen: メッセージ受信:', message);
  console.log('Offscreen: Sender info:', sender);

  // 対象がoffscreenでない場合は無視
  if (message.target && message.target !== 'offscreen') {
    console.log('Offscreen: 対象外のメッセージ:', message.target);
    return false;
  }

  if (message.action === 'fetchAPI') {
    console.log('Offscreen: API呼び出し処理開始');

    handleAPIFetch(message.data)
      .then(result => {
        console.log('Offscreen: 処理成功:', result);
        sendResponse({ success: true, data: result });
      })
      .catch(error => {
        console.error('Offscreen: 処理エラー:', error);
        console.error('Offscreen: エラー詳細:', error.stack);
        sendResponse({ success: false, error: error.message });
      });
    return true; // 非同期レスポンス
  }

  // その他のメッセージは無視
  console.log('Offscreen: 不明なメッセージ:', message.action);
  return false;
});

/**
 * API呼び出し処理（CORS制限回避版）
 */
async function handleAPIFetch(requestData) {
  const { endpoint, headers, body, provider } = requestData;

  console.log('Offscreen API呼び出し開始:', { endpoint, provider, headersCount: Object.keys(headers).length });
  console.log('Offscreen API Headers:', headers);
  console.log('Offscreen API Body preview:', body.substring(0, 200) + '...');

  try {
    console.log('Offscreen: fetch実行中...');

    const response = await fetch(endpoint, {
      method: 'POST',
      headers: headers,
      body: body,
      // Offscreen documentではCORS制限が緩和される
      mode: 'cors',
      credentials: 'omit',
      // 追加のオプション
      cache: 'no-cache',
      redirect: 'follow'
    });

    console.log('Offscreen API応答受信:', {
      status: response.status,
      ok: response.ok,
      statusText: response.statusText,
      headers: Object.fromEntries(response.headers.entries())
    });

    if (!response.ok) {
      let errorText;
      try {
        errorText = await response.text();
      } catch (textError) {
        errorText = `応答テキスト取得エラー: ${textError.message}`;
      }

      console.error('Offscreen APIエラー:', {
        status: response.status,
        statusText: response.statusText,
        error: errorText
      });

      // 詳細なエラーメッセージ
      switch (response.status) {
        case 401:
          throw new Error('APIキーが無効です。設定を確認してください。');
        case 403:
          throw new Error('APIアクセスが拒否されました。権限またはデプロイメント名を確認してください。');
        case 404:
          throw new Error('APIエンドポイントが見つかりません。URLまたはデプロイメント名を確認してください。');
        case 429:
          throw new Error('API利用制限に達しました。しばらく待ってから再試行してください。');
        case 500:
        case 502:
        case 503:
          throw new Error('APIサーバーエラーが発生しました。しばらく待ってから再試行してください。');
        default:
          throw new Error(`APIエラー (${response.status}): ${errorText || 'Unknown error'}`);
      }
    }

    console.log('Offscreen: JSON解析中...');
    const data = await response.json();
    console.log('Offscreen API成功:', data);

    // OpenAI/Azure OpenAI の応答解析
    if (provider === 'openai' || provider === 'azure') {
      if (data.choices && data.choices.length > 0) {
        const content = data.choices[0].message.content.trim();
        console.log('Offscreen: AI応答コンテンツ取得成功:', content.substring(0, 100) + '...');
        return content;
      } else {
        console.error('Offscreen: 無効なAI応答形式:', data);
        throw new Error('AIからの有効な応答が得られませんでした');
      }
    }

    console.error('Offscreen: 予期しないAPI応答形式:', data);
    throw new Error('予期しないAPI応答形式です');

  } catch (error) {
    console.error('Offscreen fetch エラー:', error);
    console.error('Offscreen fetch エラー詳細:', error.stack);

    // TypeError: Failed to fetch の詳細化
    if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
      throw new Error(
        'ネットワーク接続エラー（Offscreen document内）:\n' +
        '• インターネット接続を確認してください\n' +
        '• APIエンドポイントURLが正しいか確認してください\n' +
        '• プロキシまたはファイアウォール設定を確認してください\n' +
        '• VPNやセキュリティソフトの影響も考慮してください\n' +
        `詳細: ${error.message}`
      );
    }

    throw error;
  }
}

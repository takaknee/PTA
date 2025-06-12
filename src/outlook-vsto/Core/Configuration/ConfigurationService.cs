using System;
using System.Configuration;
using Microsoft.Extensions.Logging;
using OutlookPTAAddin.Core.Models;

namespace OutlookPTAAddin.Core.Configuration
{
    /// &lt;summary&gt;
    /// 設定管理サービス
    /// VBAの設定管理機能をVSTOで再実装
    /// &lt;/summary&gt;
    public class ConfigurationService
    {
        #region フィールド

        private readonly ILogger&lt;ConfigurationService&gt; _logger;

        // デフォルト設定値（VBA版と同様）
        private const string DEFAULT_ENDPOINT = "https://your-azure-openai-endpoint.openai.azure.com/openai/deployments/gpt-4/chat/completions?api-version=2024-02-15-preview";
        private const string DEFAULT_API_KEY = "YOUR_API_KEY_HERE";
        private const string DEFAULT_MODEL = "gpt-4";
        private const string DEFAULT_APP_NAME = "Outlook AI Helper";
        private const string DEFAULT_APP_VERSION = "1.0.0 VSTO";

        #endregion

        #region コンストラクター

        /// &lt;summary&gt;
        /// コンストラクター
        /// &lt;/summary&gt;
        /// &lt;param name="logger"&gt;ログサービス&lt;/param&gt;
        public ConfigurationService(ILogger&lt;ConfigurationService&gt; logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        #endregion

        #region パブリックメソッド

        /// &lt;summary&gt;
        /// 設定の初期化
        /// &lt;/summary&gt;
        public void Initialize()
        {
            try
            {
                _logger.LogInformation("設定サービスを初期化しています");

                // 必要に応じて初期設定を作成
                EnsureDefaultSettings();

                _logger.LogInformation("設定サービスの初期化完了");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "設定サービスの初期化中にエラーが発生しました");
                throw new ConfigurationException("設定サービスの初期化に失敗しました", ex);
            }
        }

        /// &lt;summary&gt;
        /// OpenAI API キーを取得する
        /// &lt;/summary&gt;
        /// &lt;returns&gt;APIキー&lt;/returns&gt;
        public string GetOpenAIApiKey()
        {
            try
            {
                // 環境変数から取得を試行
                var envKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
                if (!string.IsNullOrWhiteSpace(envKey))
                {
                    return envKey;
                }

                // アプリケーション設定から取得
                var configKey = ConfigurationManager.AppSettings["OpenAIApiKey"];
                if (!string.IsNullOrWhiteSpace(configKey))
                {
                    return configKey;
                }

                // Windows資格情報マネージャーから取得（将来実装）
                // var credentialKey = GetFromCredentialManager("PTA_OpenAI_ApiKey");
                // if (!string.IsNullOrWhiteSpace(credentialKey))
                // {
                //     return credentialKey;
                // }

                _logger.LogWarning("OpenAI APIキーが設定されていません。デフォルト値を使用します。");
                return DEFAULT_API_KEY;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI APIキー取得中にエラーが発生しました");
                return DEFAULT_API_KEY;
            }
        }

        /// &lt;summary&gt;
        /// OpenAI エンドポイントを取得する
        /// &lt;/summary&gt;
        /// &lt;returns&gt;エンドポイントURL&lt;/returns&gt;
        public string GetOpenAIEndpoint()
        {
            try
            {
                var endpoint = ConfigurationManager.AppSettings["OpenAIEndpoint"];
                return string.IsNullOrWhiteSpace(endpoint) ? DEFAULT_ENDPOINT : endpoint;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI エンドポイント取得中にエラーが発生しました");
                return DEFAULT_ENDPOINT;
            }
        }

        /// &lt;summary&gt;
        /// OpenAI モデル名を取得する
        /// &lt;/summary&gt;
        /// &lt;returns&gt;モデル名&lt;/returns&gt;
        public string GetOpenAIModel()
        {
            try
            {
                var model = ConfigurationManager.AppSettings["OpenAIModel"];
                return string.IsNullOrWhiteSpace(model) ? DEFAULT_MODEL : model;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI モデル名取得中にエラーが発生しました");
                return DEFAULT_MODEL;
            }
        }

        /// &lt;summary&gt;
        /// アプリケーション名を取得する
        /// &lt;/summary&gt;
        /// &lt;returns&gt;アプリケーション名&lt;/returns&gt;
        public string GetAppName()
        {
            try
            {
                var appName = ConfigurationManager.AppSettings["AppName"];
                return string.IsNullOrWhiteSpace(appName) ? DEFAULT_APP_NAME : appName;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "アプリケーション名取得中にエラーが発生しました");
                return DEFAULT_APP_NAME;
            }
        }

        /// &lt;summary&gt;
        /// アプリケーションバージョンを取得する
        /// &lt;/summary&gt;
        /// &lt;returns&gt;バージョン&lt;/returns&gt;
        public string GetAppVersion()
        {
            try
            {
                var version = ConfigurationManager.AppSettings["AppVersion"];
                return string.IsNullOrWhiteSpace(version) ? DEFAULT_APP_VERSION : version;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "アプリケーションバージョン取得中にエラーが発生しました");
                return DEFAULT_APP_VERSION;
            }
        }

        /// &lt;summary&gt;
        /// 設定情報を文字列で取得する（VBA版のShowConfigurationInfoと同等）
        /// &lt;/summary&gt;
        /// &lt;returns&gt;設定情報文字列&lt;/returns&gt;
        public string GetConfigurationInfo()
        {
            try
            {
                var info = new System.Text.StringBuilder();
                
                info.AppendLine("現在の設定:");
                info.AppendLine();
                info.AppendLine($"アプリケーション名: {GetAppName()}");
                info.AppendLine($"バージョン: {GetAppVersion()}");
                info.AppendLine($"API エンドポイント: {MaskSensitiveUrl(GetOpenAIEndpoint())}");
                info.AppendLine($"API キー: {MaskSensitiveData(GetOpenAIApiKey())}");
                info.AppendLine($"モデル: {GetOpenAIModel()}");
                info.AppendLine();
                info.AppendLine("システム情報:");
                info.AppendLine($"OS: {Environment.OSVersion}");
                info.AppendLine($".NET Framework: {Environment.Version}");
                info.AppendLine($"実行時間: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");

                return info.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "設定情報取得中にエラーが発生しました");
                return $"設定情報の取得に失敗しました: {ex.Message}";
            }
        }

        /// &lt;summary&gt;
        /// OpenAI APIキーを設定する
        /// &lt;/summary&gt;
        /// &lt;param name="apiKey"&gt;APIキー&lt;/param&gt;
        public void SetOpenAIApiKey(string apiKey)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    throw new ArgumentException("APIキーが指定されていません", nameof(apiKey));
                }

                // 将来的にWindows資格情報マネージャーに保存
                // SaveToCredentialManager("PTA_OpenAI_ApiKey", apiKey);
                
                _logger.LogInformation("OpenAI APIキーが更新されました");
                
                // 注意: 実際の実装では、app.configファイルの更新や、
                // Windows資格情報マネージャーへの保存を行う
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI APIキー設定中にエラーが発生しました");
                throw new ConfigurationException("APIキーの設定に失敗しました", ex);
            }
        }

        /// &lt;summary&gt;
        /// 設定のバリデーションを実行する
        /// &lt;/summary&gt;
        /// &lt;returns&gt;バリデーション結果&lt;/returns&gt;
        public bool ValidateConfiguration()
        {
            try
            {
                _logger.LogInformation("設定のバリデーションを実行しています");

                var apiKey = GetOpenAIApiKey();
                var endpoint = GetOpenAIEndpoint();

                var isValid = true;
                var issues = new System.Collections.Generic.List&lt;string&gt;();

                // APIキーの検証
                if (string.IsNullOrWhiteSpace(apiKey) || apiKey == DEFAULT_API_KEY)
                {
                    issues.Add("APIキーが設定されていません");
                    isValid = false;
                }

                // エンドポイントの検証
                if (string.IsNullOrWhiteSpace(endpoint) || endpoint == DEFAULT_ENDPOINT)
                {
                    issues.Add("エンドポイントが設定されていません");
                    isValid = false;
                }

                if (!isValid)
                {
                    _logger.LogWarning($"設定に問題があります: {string.Join(", ", issues)}");
                }
                else
                {
                    _logger.LogInformation("設定のバリデーション完了: 問題なし");
                }

                return isValid;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "設定バリデーション中にエラーが発生しました");
                return false;
            }
        }

        #endregion

        #region プライベートメソッド

        /// &lt;summary&gt;
        /// デフォルト設定を確保する
        /// &lt;/summary&gt;
        private void EnsureDefaultSettings()
        {
            // 必要に応じてデフォルト設定ファイルを作成
            // 現在は何もしない（将来の拡張用）
        }

        /// &lt;summary&gt;
        /// 機密データをマスクする
        /// &lt;/summary&gt;
        /// &lt;param name="data"&gt;機密データ&lt;/param&gt;
        /// &lt;returns&gt;マスクされたデータ&lt;/returns&gt;
        private string MaskSensitiveData(string data)
        {
            if (string.IsNullOrWhiteSpace(data))
            {
                return "（未設定）";
            }

            if (data.Length &lt;= 10)
            {
                return new string('*', data.Length);
            }

            return data.Substring(0, 10) + "...";
        }

        /// &lt;summary&gt;
        /// 機密URLをマスクする
        /// &lt;/summary&gt;
        /// &lt;param name="url"&gt;URL&lt;/param&gt;
        /// &lt;returns&gt;マスクされたURL&lt;/returns&gt;
        private string MaskSensitiveUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return "（未設定）";
            }

            if (url.Length &lt;= 50)
            {
                return url;
            }

            return url.Substring(0, 50) + "...";
        }

        #endregion
    }
}
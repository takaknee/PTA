using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using OutlookPTAAddin.Core.Models;
using OutlookPTAAddin.Core.Configuration;
using OutlookPTAAddin.Core.Services;

namespace OutlookPTAAddin.Infrastructure.Claude
{
    /// <summary>
    /// Claude AI API サービス
    /// Anthropic Claude APIクライアント実装
    /// </summary>
    public class ClaudeService : IAIService
    {
        #region フィールド

        private readonly HttpClient _httpClient;
        private readonly ILogger<ClaudeService> _logger;
        private readonly ConfigurationService _configService;

        // Claude API設定
        private const int MAX_TOKENS = 2000;
        private const int TIMEOUT_SECONDS = 30;
        private const string CLAUDE_API_URL = "https://api.anthropic.com/v1/messages";

        #endregion

        #region IAIService プロパティ

        /// <summary>
        /// AIプロバイダー名
        /// </summary>
        public string ProviderName => "Claude (Anthropic)";

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="httpClient">HTTPクライアント</param>
        /// <param name="logger">ログサービス</param>
        /// <param name="configService">設定サービス</param>
        public ClaudeService(HttpClient httpClient, ILogger<ClaudeService> logger, ConfigurationService configService)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));

            // HTTPクライアントの設定
            _httpClient.Timeout = TimeSpan.FromSeconds(TIMEOUT_SECONDS);
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// テキスト解析を実行する
        /// </summary>
        /// <param name="systemPrompt">システムプロンプト</param>
        /// <param name="userContent">ユーザーコンテンツ</param>
        /// <returns>解析結果</returns>
        public async Task<string> AnalyzeTextAsync(string systemPrompt, string userContent)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(systemPrompt))
                {
                    throw new ArgumentException("システムプロンプトが指定されていません", nameof(systemPrompt));
                }

                if (string.IsNullOrWhiteSpace(userContent))
                {
                    throw new ArgumentException("解析対象のコンテンツが指定されていません", nameof(userContent));
                }

                _logger.LogInformation("Claude テキスト解析開始");

                var response = await CallClaudeAPIAsync(systemPrompt, userContent);

                _logger.LogInformation("Claude テキスト解析完了");
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Claude テキスト解析中にエラーが発生しました");
                throw new ClaudeException("テキスト解析に失敗しました", ex);
            }
        }

        /// <summary>
        /// テキスト生成を実行する
        /// </summary>
        /// <param name="systemPrompt">システムプロンプト</param>
        /// <param name="userPrompt">ユーザープロンプト</param>
        /// <returns>生成されたテキスト</returns>
        public async Task<string> GenerateTextAsync(string systemPrompt, string userPrompt)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(systemPrompt))
                {
                    throw new ArgumentException("システムプロンプトが指定されていません", nameof(systemPrompt));
                }

                if (string.IsNullOrWhiteSpace(userPrompt))
                {
                    throw new ArgumentException("ユーザープロンプトが指定されていません", nameof(userPrompt));
                }

                _logger.LogInformation("Claude テキスト生成開始");

                var response = await CallClaudeAPIAsync(systemPrompt, userPrompt);

                _logger.LogInformation("Claude テキスト生成完了");
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Claude テキスト生成中にエラーが発生しました");
                throw new ClaudeException("テキスト生成に失敗しました", ex);
            }
        }

        /// <summary>
        /// API接続テストを実行する
        /// </summary>
        /// <returns>テスト結果メッセージ</returns>
        public async Task<string> TestConnectionAsync()
        {
            try
            {
                _logger.LogInformation("Claude API接続テスト開始");

                var testPrompt = "これはAPI接続テストです。「テスト成功」と日本語で返答してください。";
                var response = await CallClaudeAPIAsync("テスト用プロンプト", testPrompt);

                var result = $"✅ Claude API接続テスト成功\n\nレスポンス: {response}\n\n設定情報:\n" +
                           $"- プロバイダー: {ProviderName}\n" +
                           $"- モデル: {GetClaudeModel()}\n" +
                           $"- 最大トークン数: {MAX_TOKENS}\n" +
                           $"- タイムアウト: {TIMEOUT_SECONDS}秒";

                _logger.LogInformation("Claude API接続テスト成功");
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Claude API接続テスト失敗");
                
                var errorResult = $"❌ Claude API接続テスト失敗\n\nエラー詳細: {ex.Message}\n\n" +
                                $"確認事項:\n" +
                                $"1. Claude APIキーが正しく設定されているか\n" +
                                $"2. インターネット接続が有効か\n" +
                                $"3. ファイアウォール設定に問題はないか\n" +
                                $"4. Claude APIのクォータ制限に達していないか";

                return errorResult;
            }
        }

        /// <summary>
        /// サービスが利用可能かどうかを確認する
        /// </summary>
        /// <returns>利用可能な場合はtrue</returns>
        public bool IsAvailable()
        {
            try
            {
                var apiKey = GetClaudeApiKey();

                return !string.IsNullOrWhiteSpace(apiKey) && 
                       apiKey != "YOUR_CLAUDE_API_KEY_HERE";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Claude サービス可用性チェック中にエラーが発生しました");
                return false;
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// Claude APIを呼び出す
        /// </summary>
        /// <param name="systemPrompt">システムプロンプト</param>
        /// <param name="userPrompt">ユーザープロンプト</param>
        /// <returns>APIレスポンス</returns>
        private async Task<string> CallClaudeAPIAsync(string systemPrompt, string userPrompt)
        {
            try
            {
                // 設定値の取得
                var apiKey = GetClaudeApiKey();
                var model = GetClaudeModel();

                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    throw new ClaudeException("Claude APIキーが設定されていません");
                }

                // リクエストボディの作成（Claude APIフォーマット）
                var requestBody = new
                {
                    model = model,
                    max_tokens = MAX_TOKENS,
                    system = systemPrompt,
                    messages = new[]
                    {
                        new { role = "user", content = userPrompt }
                    }
                };

                var json = JsonConvert.SerializeObject(requestBody, Formatting.None);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // HTTPヘッダーの設定
                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
                _httpClient.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01");
                _httpClient.DefaultRequestHeaders.Add("User-Agent", "PTA-Outlook-Addin/1.0");

                _logger.LogDebug($"Claude API呼び出し: {CLAUDE_API_URL}");

                // API呼び出し
                var response = await _httpClient.PostAsync(CLAUDE_API_URL, content);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogError($"Claude API エラー: {response.StatusCode} - {responseContent}");
                    throw new ClaudeException($"API呼び出しエラー: {response.StatusCode} - {GetErrorMessage(responseContent)}");
                }

                // レスポンスの解析
                var responseData = JsonConvert.DeserializeObject<dynamic>(responseContent);
                
                if (responseData?.content != null && responseData.content.Count > 0)
                {
                    var messageContent = responseData.content[0]?.text?.ToString();
                    if (!string.IsNullOrEmpty(messageContent))
                    {
                        return messageContent.Trim();
                    }
                }

                throw new ClaudeException("APIレスポンスの解析に失敗しました");
            }
            catch (HttpRequestException ex)
            {
                throw new ClaudeException("ネットワークエラーが発生しました", ex);
            }
            catch (TaskCanceledException ex)
            {
                throw new ClaudeException("API呼び出しがタイムアウトしました", ex);
            }
            catch (JsonException ex)
            {
                throw new ClaudeException("APIレスポンスの解析に失敗しました", ex);
            }
        }

        /// <summary>
        /// Claude APIキーを取得する
        /// </summary>
        /// <returns>APIキー</returns>
        private string GetClaudeApiKey()
        {
            try
            {
                // 環境変数から取得を試行
                var envKey = Environment.GetEnvironmentVariable("CLAUDE_API_KEY");
                if (!string.IsNullOrWhiteSpace(envKey))
                {
                    return envKey;
                }

                // アプリケーション設定から取得
                var configKey = System.Configuration.ConfigurationManager.AppSettings["ClaudeApiKey"];
                if (!string.IsNullOrWhiteSpace(configKey))
                {
                    return configKey;
                }

                _logger.LogWarning("Claude APIキーが設定されていません");
                return "YOUR_CLAUDE_API_KEY_HERE";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Claude APIキー取得中にエラーが発生しました");
                return "YOUR_CLAUDE_API_KEY_HERE";
            }
        }

        /// <summary>
        /// Claude モデル名を取得する
        /// </summary>
        /// <returns>モデル名</returns>
        private string GetClaudeModel()
        {
            try
            {
                var model = System.Configuration.ConfigurationManager.AppSettings["ClaudeModel"];
                return string.IsNullOrWhiteSpace(model) ? "claude-3-sonnet-20240229" : model;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Claude モデル名取得中にエラーが発生しました");
                return "claude-3-sonnet-20240229";
            }
        }

        /// <summary>
        /// エラーレスポンスからメッセージを抽出する
        /// </summary>
        /// <param name="responseContent">レスポンス内容</param>
        /// <returns>エラーメッセージ</returns>
        private string GetErrorMessage(string responseContent)
        {
            try
            {
                var errorData = JsonConvert.DeserializeObject<dynamic>(responseContent);
                return errorData?.error?.message?.ToString() ?? "詳細不明";
            }
            catch
            {
                return responseContent;
            }
        }

        #endregion
    }

    /// <summary>
    /// Claude API 例外クラス
    /// </summary>
    public class ClaudeException : Exception
    {
        public ClaudeException(string message) : base(message) { }
        public ClaudeException(string message, Exception innerException) : base(message, innerException) { }
    }
}
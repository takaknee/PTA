using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using OutlookPTAAddin.Core.Models;
using OutlookPTAAddin.Core.Services;

namespace OutlookPTAAddin.Infrastructure.OpenAI
{
    /// <summary>
    /// OpenAI API サービス
    /// VBAのOpenAI API呼び出し機能をVSTOで再実装
    /// </summary>
    public class OpenAIService : IAIService
    {
        #region フィールド

        private readonly HttpClient _httpClient;
        private readonly ILogger<OpenAIService> _logger;
        private readonly ConfigurationService _configService;

        // VBA版と同様の設定値
        private const int MAX_TOKENS = 2000;
        private const int TIMEOUT_SECONDS = 30;

        #endregion

        #region IAIService プロパティ

        /// <summary>
        /// AIプロバイダー名
        /// </summary>
        public string ProviderName => "OpenAI";

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="httpClient">HTTPクライアント</param>
        /// <param name="logger">ログサービス</param>
        /// <param name="configService">設定サービス</param>
        public OpenAIService(HttpClient httpClient, ILogger<OpenAIService> logger, ConfigurationService configService)
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

                _logger.LogInformation("OpenAI テキスト解析開始");

                var response = await CallOpenAIAPIAsync(systemPrompt, userContent);

                _logger.LogInformation("OpenAI テキスト解析完了");
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI テキスト解析中にエラーが発生しました");
                throw new OpenAIException("テキスト解析に失敗しました", ex);
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

                _logger.LogInformation("OpenAI テキスト生成開始");

                var response = await CallOpenAIAPIAsync(systemPrompt, userPrompt);

                _logger.LogInformation("OpenAI テキスト生成完了");
                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI テキスト生成中にエラーが発生しました");
                throw new OpenAIException("テキスト生成に失敗しました", ex);
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
                _logger.LogInformation("OpenAI API接続テスト開始");

                var testPrompt = "これはAPI接続テストです。「テスト成功」と日本語で返答してください。";
                var response = await CallOpenAIAPIAsync("テスト用プロンプト", testPrompt);

                var result = $"✅ API接続テスト成功\\n\\nレスポンス: {response}\\n\\n設定情報:\\n" +
                           $"- エンドポイント: {_configService.GetOpenAIEndpoint()}\\n" +
                           $"- モデル: {_configService.GetOpenAIModel()}\\n" +
                           $"- 最大トークン数: {MAX_TOKENS}\\n" +
                           $"- タイムアウト: {TIMEOUT_SECONDS}秒";

                _logger.LogInformation("OpenAI API接続テスト成功");
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI API接続テスト失敗");
                
                var errorResult = $"❌ API接続テスト失敗\\n\\nエラー詳細: {ex.Message}\\n\\n" +
                                $"確認事項:\\n" +
                                $"1. APIキーが正しく設定されているか\\n" +
                                $"2. エンドポイントURLが正しいか\\n" +
                                $"3. インターネット接続が有効か\\n" +
                                $"4. ファイアウォール設定に問題はないか";

                return errorResult;
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// OpenAI APIを呼び出す
        /// </summary>
        /// <param name="systemPrompt">システムプロンプト</param>
        /// <param name="userPrompt">ユーザープロンプト</param>
        /// <returns>APIレスポンス</returns>
        private async Task<string> CallOpenAIAPIAsync(string systemPrompt, string userPrompt)
        {
            try
            {
                // 設定値の取得
                var apiKey = _configService.GetOpenAIApiKey();
                var endpoint = _configService.GetOpenAIEndpoint();
                var model = _configService.GetOpenAIModel();

                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    throw new OpenAIException("OpenAI APIキーが設定されていません");
                }

                if (string.IsNullOrWhiteSpace(endpoint))
                {
                    throw new OpenAIException("OpenAI エンドポイントが設定されていません");
                }

                // リクエストボディの作成
                var requestBody = new
                {
                    model = model,
                    messages = new[]
                    {
                        new { role = "system", content = systemPrompt },
                        new { role = "user", content = userPrompt }
                    },
                    max_tokens = MAX_TOKENS,
                    temperature = 0.7,
                    top_p = 1.0,
                    frequency_penalty = 0.0,
                    presence_penalty = 0.0
                };

                var json = JsonConvert.SerializeObject(requestBody, Formatting.None);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // HTTPヘッダーの設定
                _httpClient.DefaultRequestHeaders.Clear();
                
                // Azure OpenAI の場合は api-key、OpenAI の場合は Authorization
                if (endpoint.Contains("openai.azure.com"))
                {
                    _httpClient.DefaultRequestHeaders.Add("api-key", apiKey);
                }
                else
                {
                    _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                }

                _httpClient.DefaultRequestHeaders.Add("User-Agent", "PTA-Outlook-Addin/1.0");

                _logger.LogDebug($"OpenAI API呼び出し: {endpoint}");

                // API呼び出し
                var response = await _httpClient.PostAsync(endpoint, content);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    _logger.LogError($"OpenAI API エラー: {response.StatusCode} - {responseContent}");
                    throw new OpenAIException($"API呼び出しエラー: {response.StatusCode} - {GetErrorMessage(responseContent)}");
                }

                // レスポンスの解析
                var responseData = JsonConvert.DeserializeObject<dynamic>(responseContent);
                
                if (responseData?.choices != null && responseData.choices.Count > 0)
                {
                    var messageContent = responseData.choices[0].message?.content?.ToString();
                    if (!string.IsNullOrEmpty(messageContent))
                    {
                        return messageContent.Trim();
                    }
                }

                throw new OpenAIException("APIレスポンスの解析に失敗しました");
            }
            catch (HttpRequestException ex)
            {
                throw new OpenAIException("ネットワークエラーが発生しました", ex);
            }
            catch (TaskCanceledException ex)
            {
                throw new OpenAIException("API呼び出しがタイムアウトしました", ex);
            }
            catch (JsonException ex)
            {
                throw new OpenAIException("APIレスポンスの解析に失敗しました", ex);
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

        /// <summary>
        /// サービスが利用可能かどうかを確認する
        /// </summary>
        /// <returns>利用可能な場合はtrue</returns>
        public bool IsAvailable()
        {
            try
            {
                var apiKey = _configService.GetOpenAIApiKey();
                var endpoint = _configService.GetOpenAIEndpoint();

                return !string.IsNullOrWhiteSpace(apiKey) && 
                       !string.IsNullOrWhiteSpace(endpoint) &&
                       apiKey != "YOUR_API_KEY_HERE";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OpenAI サービス可用性チェック中にエラーが発生しました");
                return false;
            }
        }

        #endregion
    }
}
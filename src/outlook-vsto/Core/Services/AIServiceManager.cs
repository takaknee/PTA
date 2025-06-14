using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using OutlookPTAAddin.Core.Models;

namespace OutlookPTAAddin.Core.Services
{
    /// <summary>
    /// AIサービス管理クラス
    /// 複数のAIプロバイダーを統合管理し、可用性に応じて自動切り替えを行う
    /// </summary>
    public class AIServiceManager
    {
        #region フィールド

        private readonly ILogger<AIServiceManager> _logger;
        private readonly List<IAIService> _aiServices;
        private IAIService _primaryService;

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="logger">ログサービス</param>
        /// <param name="aiServices">利用可能なAIサービス一覧</param>
        public AIServiceManager(ILogger<AIServiceManager> logger, IEnumerable<IAIService> aiServices)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _aiServices = aiServices?.ToList() ?? throw new ArgumentNullException(nameof(aiServices));

            if (!_aiServices.Any())
            {
                throw new ArgumentException("AIサービスが登録されていません", nameof(aiServices));
            }

            // プライマリサービスの設定（最初に利用可能なサービス）
            RefreshPrimaryService();
        }

        #endregion

        #region パブリックプロパティ

        /// <summary>
        /// 現在のプライマリAIサービス
        /// </summary>
        public IAIService PrimaryService => _primaryService;

        /// <summary>
        /// 利用可能なAIサービス一覧
        /// </summary>
        public IReadOnlyList<IAIService> AvailableServices => _aiServices.Where(s => s.IsAvailable()).ToList();

        /// <summary>
        /// 登録されているAIサービス一覧
        /// </summary>
        public IReadOnlyList<IAIService> RegisteredServices => _aiServices.AsReadOnly();

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// テキスト解析を実行する（フォールバック付き）
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

                _logger.LogInformation("AIサービスマネージャー: テキスト解析開始");

                // プライマリサービスを試行
                if (_primaryService?.IsAvailable() == true)
                {
                    try
                    {
                        _logger.LogDebug($"プライマリサービス使用: {_primaryService.ProviderName}");
                        var result = await _primaryService.AnalyzeTextAsync(systemPrompt, userContent);
                        _logger.LogInformation($"テキスト解析成功 - プロバイダー: {_primaryService.ProviderName}");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"プライマリサービス（{_primaryService.ProviderName}）で解析に失敗しました。フォールバックを試行します。");
                    }
                }

                // フォールバックサービスを試行
                var availableServices = AvailableServices.Where(s => s != _primaryService).ToList();
                
                foreach (var service in availableServices)
                {
                    try
                    {
                        _logger.LogDebug($"フォールバックサービス使用: {service.ProviderName}");
                        var result = await service.AnalyzeTextAsync(systemPrompt, userContent);
                        _logger.LogInformation($"テキスト解析成功（フォールバック） - プロバイダー: {service.ProviderName}");
                        
                        // 成功したサービスをプライマリに設定
                        _primaryService = service;
                        
                        return result;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"フォールバックサービス（{service.ProviderName}）でも解析に失敗しました。");
                    }
                }

                throw new AIServiceException("すべてのAIサービスで解析に失敗しました");
            }
            catch (Exception ex) when (!(ex is AIServiceException))
            {
                _logger.LogError(ex, "AIサービスマネージャー: テキスト解析中にエラーが発生しました");
                throw new AIServiceException("テキスト解析に失敗しました", ex);
            }
        }

        /// <summary>
        /// テキスト生成を実行する（フォールバック付き）
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

                _logger.LogInformation("AIサービスマネージャー: テキスト生成開始");

                // プライマリサービスを試行
                if (_primaryService?.IsAvailable() == true)
                {
                    try
                    {
                        _logger.LogDebug($"プライマリサービス使用: {_primaryService.ProviderName}");
                        var result = await _primaryService.GenerateTextAsync(systemPrompt, userPrompt);
                        _logger.LogInformation($"テキスト生成成功 - プロバイダー: {_primaryService.ProviderName}");
                        return result;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"プライマリサービス（{_primaryService.ProviderName}）で生成に失敗しました。フォールバックを試行します。");
                    }
                }

                // フォールバックサービスを試行
                var availableServices = AvailableServices.Where(s => s != _primaryService).ToList();
                
                foreach (var service in availableServices)
                {
                    try
                    {
                        _logger.LogDebug($"フォールバックサービス使用: {service.ProviderName}");
                        var result = await service.GenerateTextAsync(systemPrompt, userPrompt);
                        _logger.LogInformation($"テキスト生成成功（フォールバック） - プロバイダー: {service.ProviderName}");
                        
                        // 成功したサービスをプライマリに設定
                        _primaryService = service;
                        
                        return result;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"フォールバックサービス（{service.ProviderName}）でも生成に失敗しました。");
                    }
                }

                throw new AIServiceException("すべてのAIサービスで生成に失敗しました");
            }
            catch (Exception ex) when (!(ex is AIServiceException))
            {
                _logger.LogError(ex, "AIサービスマネージャー: テキスト生成中にエラーが発生しました");
                throw new AIServiceException("テキスト生成に失敗しました", ex);
            }
        }

        /// <summary>
        /// 指定されたプロバイダーのAPI接続テストを実行する
        /// </summary>
        /// <param name="providerName">プロバイダー名（nullの場合は全て）</param>
        /// <returns>テスト結果メッセージ</returns>
        public async Task<string> TestConnectionAsync(string providerName = null)
        {
            try
            {
                _logger.LogInformation($"AIサービスマネージャー: 接続テスト開始 - プロバイダー: {providerName ?? "全て"}");

                var results = new List<string>();

                // 指定されたプロバイダーまたは全てのプロバイダーをテスト
                var servicesToTest = string.IsNullOrWhiteSpace(providerName) 
                    ? _aiServices 
                    : _aiServices.Where(s => s.ProviderName.Equals(providerName, StringComparison.OrdinalIgnoreCase)).ToList();

                if (!servicesToTest.Any())
                {
                    return $"❌ 指定されたプロバイダー '{providerName}' が見つかりません";
                }

                foreach (var service in servicesToTest)
                {
                    try
                    {
                        _logger.LogDebug($"接続テスト実行: {service.ProviderName}");
                        var testResult = await service.TestConnectionAsync();
                        results.Add($"📡 {service.ProviderName}:\n{testResult}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"接続テスト失敗: {service.ProviderName}");
                        results.Add($"📡 {service.ProviderName}:\n❌ 接続テスト失敗\nエラー: {ex.Message}");
                    }
                }

                var finalResult = string.Join("\n\n" + new string('=', 50) + "\n\n", results);
                
                // プライマリサービスの更新
                RefreshPrimaryService();
                
                var summary = $"=== AIサービス接続テスト結果 ===\n\n{finalResult}\n\n" +
                             $"現在のプライマリサービス: {(_primaryService?.ProviderName ?? "なし")}\n" +
                             $"利用可能なサービス数: {AvailableServices.Count}/{_aiServices.Count}";

                _logger.LogInformation("AIサービス接続テスト完了");
                return summary;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "AIサービス接続テスト中にエラーが発生しました");
                return $"❌ 接続テスト中にエラーが発生しました: {ex.Message}";
            }
        }

        /// <summary>
        /// プライマリサービスを手動で設定する
        /// </summary>
        /// <param name="providerName">プロバイダー名</param>
        /// <returns>設定に成功した場合はtrue</returns>
        public bool SetPrimaryService(string providerName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(providerName))
                {
                    throw new ArgumentException("プロバイダー名が指定されていません", nameof(providerName));
                }

                var service = _aiServices.FirstOrDefault(s => 
                    s.ProviderName.Equals(providerName, StringComparison.OrdinalIgnoreCase));

                if (service == null)
                {
                    _logger.LogWarning($"指定されたプロバイダーが見つかりません: {providerName}");
                    return false;
                }

                if (!service.IsAvailable())
                {
                    _logger.LogWarning($"指定されたプロバイダーは利用できません: {providerName}");
                    return false;
                }

                _primaryService = service;
                _logger.LogInformation($"プライマリサービスを設定しました: {providerName}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "プライマリサービス設定中にエラーが発生しました");
                return false;
            }
        }

        /// <summary>
        /// AIサービスの状態情報を取得する
        /// </summary>
        /// <returns>状態情報文字列</returns>
        public string GetServiceStatus()
        {
            try
            {
                var status = new List<string>
                {
                    "=== AIサービス状態 ===",
                    ""
                };

                foreach (var service in _aiServices)
                {
                    var isAvailable = service.IsAvailable();
                    var isPrimary = service == _primaryService;
                    
                    var statusIcon = isAvailable ? "✅" : "❌";
                    var primaryIcon = isPrimary ? "⭐" : "  ";
                    
                    status.Add($"{primaryIcon} {statusIcon} {service.ProviderName} - {(isAvailable ? "利用可能" : "利用不可")}");
                }

                status.Add("");
                status.Add($"プライマリサービス: {(_primaryService?.ProviderName ?? "なし")}");
                status.Add($"利用可能サービス: {AvailableServices.Count}/{_aiServices.Count}");

                return string.Join("\n", status);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "サービス状態取得中にエラーが発生しました");
                return $"状態取得エラー: {ex.Message}";
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// プライマリサービスを更新する
        /// </summary>
        private void RefreshPrimaryService()
        {
            try
            {
                // 現在のプライマリサービスが利用可能な場合はそのまま
                if (_primaryService?.IsAvailable() == true)
                {
                    return;
                }

                // 利用可能な最初のサービスを選択
                _primaryService = _aiServices.FirstOrDefault(s => s.IsAvailable());

                if (_primaryService != null)
                {
                    _logger.LogInformation($"プライマリサービスを更新しました: {_primaryService.ProviderName}");
                }
                else
                {
                    _logger.LogWarning("利用可能なAIサービスがありません");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "プライマリサービス更新中にエラーが発生しました");
            }
        }

        #endregion
    }

    /// <summary>
    /// AIサービス関連の例外
    /// </summary>
    public class AIServiceException : Exception
    {
        public AIServiceException(string message) : base(message) { }
        public AIServiceException(string message, Exception innerException) : base(message, innerException) { }
    }
}
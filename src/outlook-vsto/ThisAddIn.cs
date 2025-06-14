using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Outlook;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OutlookPTAAddin.Core.Services;
using OutlookPTAAddin.Core.Configuration;
using OutlookPTAAddin.Infrastructure.OpenAI;
using OutlookPTAAddin.Infrastructure.Logging;
using OutlookPTAAddin.Infrastructure.Claude;

namespace OutlookPTAAddin
{
    /// <summary>
    /// PTA情報配信システム - Outlook VSTOアドイン
    /// </summary>
    [ComVisible(true)]
    public partial class ThisAddIn
    {
        #region フィールド

        private ServiceProvider _serviceProvider;
        private ILogger<ThisAddIn> _logger;
        private EmailAnalysisService _emailAnalysisService;
        private EmailComposerService _emailComposerService;
        private ConfigurationService _configurationService;
        private AIServiceManager _aiServiceManager;

        #endregion

        #region イベントハンドラー

        /// <summary>
        /// アドイン起動時の初期化処理
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // 依存関係注入コンテナの設定
                ConfigureServices();

                // サービスの初期化
                InitializeServices();

                // UI要素の初期化
                InitializeUI();

                // イベントハンドラーの登録
                RegisterEventHandlers();

                _logger?.LogInformation("PTA Outlook アドインが正常に起動しました");
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "アドイン起動時にエラーが発生しました");
                
                // ユーザーにエラー通知
                System.Windows.Forms.MessageBox.Show(
                    $"PTA Outlookアドインの起動に失敗しました。\\n\\n詳細: {ex.Message}",
                    "アドイン起動エラー",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// アドイン終了時の cleanup 処理
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                _logger?.LogInformation("PTA Outlook アドインをシャットダウンしています");

                // イベントハンドラーの登録解除
                UnregisterEventHandlers();

                // リソースのクリーンアップ
                CleanupResources();

                _logger?.LogInformation("PTA Outlook アドインが正常にシャットダウンしました");
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "アドイン終了時にエラーが発生しました");
            }
        }

        #endregion

        #region 初期化メソッド

        /// <summary>
        /// 依存関係注入の設定
        /// </summary>
        private void ConfigureServices()
        {
            var services = new ServiceCollection();

            // ログサービスの設定
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.AddProvider(new FileLoggerProvider("PTA_Outlook_Addin.log"));
            });

            // 設定サービス
            services.AddSingleton<ConfigurationService>();

            // AI サービス  
            services.AddHttpClient<OpenAIService>();
            services.AddHttpClient<ClaudeService>();
            
            // AI サービスの登録
            services.AddSingleton<OpenAIService>();
            services.AddSingleton<ClaudeService>();
            
            // IAIService の実装を登録
            services.AddSingleton<IAIService>(provider => provider.GetRequiredService<OpenAIService>());
            services.AddSingleton<IAIService>(provider => provider.GetRequiredService<ClaudeService>());
            
            // AI サービス管理者
            services.AddSingleton<AIServiceManager>(provider =>
            {
                var logger = provider.GetRequiredService<ILogger<AIServiceManager>>();
                var openAIService = provider.GetRequiredService<OpenAIService>();
                var claudeService = provider.GetRequiredService<ClaudeService>();
                return new AIServiceManager(logger, new IAIService[] { openAIService, claudeService });
            });

            // ビジネスロジックサービス
            services.AddSingleton<EmailAnalysisService>();
            services.AddSingleton<EmailComposerService>();

            // Graph API サービス
            services.AddSingleton<Infrastructure.Graph.GraphService>();

            // 統合サービス
            services.AddSingleton<IntegrationService>();

            _serviceProvider = services.BuildServiceProvider();
        }

        /// <summary>
        /// サービスの初期化
        /// </summary>
        private void InitializeServices()
        {
            _logger = _serviceProvider.GetRequiredService<ILogger<ThisAddIn>>();
            _configurationService = _serviceProvider.GetRequiredService<ConfigurationService>();
            _emailAnalysisService = _serviceProvider.GetRequiredService<EmailAnalysisService>();
            _emailComposerService = _serviceProvider.GetRequiredService<EmailComposerService>();
            _aiServiceManager = _serviceProvider.GetRequiredService<AIServiceManager>();

            // 設定の初期化
            _configurationService.Initialize();
        }

        /// <summary>
        /// UI要素の初期化
        /// </summary>
        private void InitializeUI()
        {
            // リボンUIの初期化はRibbonクラスで実行される
            _logger?.LogDebug("UI要素の初期化完了");
        }

        /// <summary>
        /// イベントハンドラーの登録
        /// </summary>
        private void RegisterEventHandlers()
        {
            // Outlookアプリケーションレベルのイベント
            this.Application.ItemSend += Application_ItemSend;
            
            _logger?.LogDebug("イベントハンドラーの登録完了");
        }

        /// <summary>
        /// イベントハンドラーの登録解除
        /// </summary>
        private void UnregisterEventHandlers()
        {
            try
            {
                this.Application.ItemSend -= Application_ItemSend;
                _logger?.LogDebug("イベントハンドラーの登録解除完了");
            }
            catch (System.Exception ex)
            {
                _logger?.LogWarning(ex, "イベントハンドラーの登録解除中にエラーが発生しました");
            }
        }

        /// <summary>
        /// リソースのクリーンアップ
        /// </summary>
        private void CleanupResources()
        {
            try
            {
                _serviceProvider?.Dispose();
            }
            catch (System.Exception ex)
            {
                _logger?.LogWarning(ex, "リソースクリーンアップ中にエラーが発生しました");
            }
        }

        #endregion

        #region Outlookイベントハンドラー

        /// <summary>
        /// メール送信前のイベントハンドラー
        /// </summary>
        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            try
            {
                if (Item is MailItem mailItem)
                {
                    _logger?.LogDebug($"メール送信イベント: {mailItem.Subject}");
                    
                    // 必要に応じて送信前のチェック処理を追加
                    // Cancel = ShouldCancelSend(mailItem);
                }
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "メール送信イベント処理中にエラーが発生しました");
            }
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// メール解析の実行
        /// </summary>
        /// <param name="mailItem">解析対象のメールアイテム</param>
        /// <returns>解析結果</returns>
        public async System.Threading.Tasks.Task<string> AnalyzeEmailAsync(MailItem mailItem)
        {
            try
            {
                if (mailItem == null)
                {
                    throw new ArgumentNullException(nameof(mailItem));
                }

                _logger?.LogInformation($"メール解析開始: {mailItem.Subject}");
                
                var result = await _emailAnalysisService.AnalyzeEmailAsync(mailItem);
                
                _logger?.LogInformation("メール解析完了");
                return result;
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "メール解析中にエラーが発生しました");
                throw;
            }
        }

        /// <summary>
        /// 営業断りメールの作成
        /// </summary>
        /// <param name="originalMail">元のメール</param>
        /// <returns>作成されたメール</returns>
        public async System.Threading.Tasks.Task<MailItem> CreateRejectionEmailAsync(MailItem originalMail)
        {
            try
            {
                _logger?.LogInformation("営業断りメール作成開始");
                
                var newMail = await _emailComposerService.CreateRejectionEmailAsync(originalMail);
                
                _logger?.LogInformation("営業断りメール作成完了");
                return newMail;
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "営業断りメール作成中にエラーが発生しました");
                throw;
            }
        }

        /// <summary>
        /// 承諾メールの作成
        /// </summary>
        /// <param name="originalMail">元のメール</param>
        /// <returns>作成されたメール</returns>
        public async System.Threading.Tasks.Task<MailItem> CreateAcceptanceEmailAsync(MailItem originalMail)
        {
            try
            {
                _logger?.LogInformation("承諾メール作成開始");
                
                var newMail = await _emailComposerService.CreateAcceptanceEmailAsync(originalMail);
                
                _logger?.LogInformation("承諾メール作成完了");
                return newMail;
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "承諾メール作成中にエラーが発生しました");
                throw;
            }
        }

        #endregion

        #region プロパティ

        /// <summary>
        /// サービスプロバイダーへのアクセス
        /// </summary>
        public ServiceProvider ServiceProvider => _serviceProvider;

        #endregion

        #region VSTOが生成するコード

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
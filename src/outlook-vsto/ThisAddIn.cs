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

namespace OutlookPTAAddin
{
    /// &lt;summary&gt;
    /// PTA情報配信システム - Outlook VSTOアドイン
    /// &lt;/summary&gt;
    [ComVisible(true)]
    public partial class ThisAddIn
    {
        #region フィールド

        private ServiceProvider _serviceProvider;
        private ILogger&lt;ThisAddIn&gt; _logger;
        private EmailAnalysisService _emailAnalysisService;
        private EmailComposerService _emailComposerService;
        private ConfigurationService _configurationService;

        #endregion

        #region イベントハンドラー

        /// &lt;summary&gt;
        /// アドイン起動時の初期化処理
        /// &lt;/summary&gt;
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

        /// &lt;summary&gt;
        /// アドイン終了時の cleanup 処理
        /// &lt;/summary&gt;
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

        /// &lt;summary&gt;
        /// 依存関係注入の設定
        /// &lt;/summary&gt;
        private void ConfigureServices()
        {
            var services = new ServiceCollection();

            // ログサービスの設定
            services.AddLogging(builder =&gt;
            {
                builder.AddConsole();
                builder.AddProvider(new FileLoggerProvider("PTA_Outlook_Addin.log"));
            });

            // 設定サービス
            services.AddSingleton&lt;ConfigurationService&gt;();

            // OpenAI サービス  
            services.AddHttpClient&lt;OpenAIService&gt;();
            services.AddSingleton&lt;OpenAIService&gt;();

            // ビジネスロジックサービス
            services.AddSingleton&lt;EmailAnalysisService&gt;();
            services.AddSingleton&lt;EmailComposerService&gt;();

            _serviceProvider = services.BuildServiceProvider();
        }

        /// &lt;summary&gt;
        /// サービスの初期化
        /// &lt;/summary&gt;
        private void InitializeServices()
        {
            _logger = _serviceProvider.GetRequiredService&lt;ILogger&lt;ThisAddIn&gt;&gt;();
            _configurationService = _serviceProvider.GetRequiredService&lt;ConfigurationService&gt;();
            _emailAnalysisService = _serviceProvider.GetRequiredService&lt;EmailAnalysisService&gt;();
            _emailComposerService = _serviceProvider.GetRequiredService&lt;EmailComposerService&gt;();

            // 設定の初期化
            _configurationService.Initialize();
        }

        /// &lt;summary&gt;
        /// UI要素の初期化
        /// &lt;/summary&gt;
        private void InitializeUI()
        {
            // リボンUIの初期化はRibbonクラスで実行される
            _logger?.LogDebug("UI要素の初期化完了");
        }

        /// &lt;summary&gt;
        /// イベントハンドラーの登録
        /// &lt;/summary&gt;
        private void RegisterEventHandlers()
        {
            // Outlookアプリケーションレベルのイベント
            this.Application.ItemSend += Application_ItemSend;
            
            _logger?.LogDebug("イベントハンドラーの登録完了");
        }

        /// &lt;summary&gt;
        /// イベントハンドラーの登録解除
        /// &lt;/summary&gt;
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

        /// &lt;summary&gt;
        /// リソースのクリーンアップ
        /// &lt;/summary&gt;
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

        /// &lt;summary&gt;
        /// メール送信前のイベントハンドラー
        /// &lt;/summary&gt;
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

        /// &lt;summary&gt;
        /// メール解析の実行
        /// &lt;/summary&gt;
        /// &lt;param name="mailItem"&gt;解析対象のメールアイテム&lt;/param&gt;
        /// &lt;returns&gt;解析結果&lt;/returns&gt;
        public async System.Threading.Tasks.Task&lt;string&gt; AnalyzeEmailAsync(MailItem mailItem)
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

        /// &lt;summary&gt;
        /// 営業断りメールの作成
        /// &lt;/summary&gt;
        /// &lt;param name="originalMail"&gt;元のメール&lt;/param&gt;
        /// &lt;returns&gt;作成されたメール&lt;/returns&gt;
        public async System.Threading.Tasks.Task&lt;MailItem&gt; CreateRejectionEmailAsync(MailItem originalMail)
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

        /// &lt;summary&gt;
        /// 承諾メールの作成
        /// &lt;/summary&gt;
        /// &lt;param name="originalMail"&gt;元のメール&lt;/param&gt;
        /// &lt;returns&gt;作成されたメール&lt;/returns&gt;
        public async System.Threading.Tasks.Task&lt;MailItem&gt; CreateAcceptanceEmailAsync(MailItem originalMail)
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

        /// &lt;summary&gt;
        /// サービスプロバイダーへのアクセス
        /// &lt;/summary&gt;
        public ServiceProvider ServiceProvider =&gt; _serviceProvider;

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
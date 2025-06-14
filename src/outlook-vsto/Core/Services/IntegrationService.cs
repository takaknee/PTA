using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using OutlookPTAAddin.Core.Services;
using OutlookPTAAddin.Infrastructure.Graph;

namespace OutlookPTAAddin.Core.Services
{
    /// <summary>
    /// 統合サービス管理クラス
    /// AI、Graph、その他の外部サービスを統合管理
    /// </summary>
    public class IntegrationService
    {
        #region フィールド

        private readonly ILogger<IntegrationService> _logger;
        private readonly AIServiceManager _aiServiceManager;
        private readonly GraphService _graphService;

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="logger">ログサービス</param>
        /// <param name="aiServiceManager">AIサービス管理者</param>
        /// <param name="graphService">Graph サービス</param>
        public IntegrationService(
            ILogger<IntegrationService> logger,
            AIServiceManager aiServiceManager,
            GraphService graphService)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _aiServiceManager = aiServiceManager ?? throw new ArgumentNullException(nameof(aiServiceManager));
            _graphService = graphService ?? throw new ArgumentNullException(nameof(graphService));
        }

        #endregion

        #region AI + Teams 統合機能

        /// <summary>
        /// メール解析結果をTeamsに送信する
        /// </summary>
        /// <param name="emailContent">メール内容</param>
        /// <param name="teamId">チームID</param>
        /// <param name="channelId">チャンネルID</param>
        /// <returns>送信結果</returns>
        public async Task<bool> AnalyzeEmailAndShareToTeamsAsync(string emailContent, string teamId, string channelId)
        {
            try
            {
                _logger.LogInformation("メール解析 → Teams 共有統合処理開始");

                // Step 1: AI でメール解析
                var systemPrompt = @"以下のメールを解析し、Teams共有用の簡潔なサマリーを作成してください。
重要なポイント、アクション項目、期限などを含めてください。
Markdown形式で見やすく整理してください。";

                var analysisResult = await _aiServiceManager.AnalyzeTextAsync(systemPrompt, emailContent);

                // Step 2: Teams メッセージフォーマット作成
                var teamsMessage = FormatForTeams(analysisResult);

                // Step 3: Teams に送信
                var success = await _graphService.SendTeamsMessageAsync(teamId, channelId, teamsMessage);

                _logger.LogInformation($"メール解析 → Teams 共有統合処理完了: {success}");
                return success;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "メール解析 → Teams 共有統合処理中にエラーが発生しました");
                throw new IntegrationException("メール解析・Teams共有に失敗しました", ex);
            }
        }

        /// <summary>
        /// AI要約結果をSharePointリストに保存する
        /// </summary>
        /// <param name="emailContent">メール内容</param>
        /// <param name="siteId">サイトID</param>
        /// <param name="listId">リストID</param>
        /// <returns>保存結果</returns>
        public async Task<string> AnalyzeEmailAndSaveToSharePointAsync(string emailContent, string siteId, string listId)
        {
            try
            {
                _logger.LogInformation("メール解析 → SharePoint 保存統合処理開始");

                // Step 1: AI でメール解析・構造化
                var systemPrompt = @"以下のメールを解析し、構造化されたデータとして整理してください。
以下のJSONフォーマットで出力してください：
{
  ""title"": ""メールの件名または要約"",
  ""sender"": ""送信者"",
  ""category"": ""カテゴリ（営業/問い合わせ/重要など）"",
  ""priority"": ""優先度（高/中/低）"",
  ""summary"": ""要約"",
  ""actionItems"": [""アクション項目のリスト""],
  ""deadline"": ""期限（あれば）"",
  ""tags"": [""タグのリスト""]
}";

                var analysisResult = await _aiServiceManager.AnalyzeTextAsync(systemPrompt, emailContent);

                // Step 2: SharePoint リストアイテムとして保存
                // 将来実装: Graph API を使用してSharePointに保存
                // var itemId = await SaveToSharePointList(siteId, listId, analysisResult);

                _logger.LogInformation("メール解析 → SharePoint 保存統合処理完了（準備中）");
                return "sample-item-id";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "メール解析 → SharePoint 保存統合処理中にエラーが発生しました");
                throw new IntegrationException("メール解析・SharePoint保存に失敗しました", ex);
            }
        }

        #endregion

        #region AI + Calendar 統合機能

        /// <summary>
        /// メールからミーティング情報を抽出してカレンダーイベントを作成する
        /// </summary>
        /// <param name="emailContent">メール内容</param>
        /// <returns>作成されたイベントID</returns>
        public async Task<string> ExtractMeetingAndCreateCalendarEventAsync(string emailContent)
        {
            try
            {
                _logger.LogInformation("ミーティング抽出 → カレンダー作成統合処理開始");

                // Step 1: AI でミーティング情報抽出
                var systemPrompt = @"以下のメールからミーティング・会議の情報を抽出してください。
以下のJSONフォーマットで出力してください：
{
  ""isMeetingRequest"": true/false,
  ""subject"": ""会議の件名"",
  ""startDateTime"": ""開始日時（ISO8601形式）"",
  ""endDateTime"": ""終了日時（ISO8601形式）"",
  ""location"": ""場所"",
  ""attendees"": [""参加者のメールアドレスリスト""],
  ""agenda"": ""議題・アジェンダ"",
  ""description"": ""詳細説明""
}

ミーティングの情報が含まれていない場合は isMeetingRequest を false にしてください。";

                var extractionResult = await _aiServiceManager.AnalyzeTextAsync(systemPrompt, emailContent);

                // Step 2: 抽出結果の解析とカレンダーイベント作成
                // 将来実装: JSON解析とGraph API呼び出し
                // var meetingInfo = JsonConvert.DeserializeObject<MeetingInfo>(extractionResult);
                // if (meetingInfo.IsMeetingRequest)
                // {
                //     var eventId = await _graphService.CreateCalendarEventAsync(
                //         meetingInfo.Subject,
                //         meetingInfo.StartDateTime,
                //         meetingInfo.EndDateTime,
                //         meetingInfo.Attendees);
                //     return eventId;
                // }

                _logger.LogInformation("ミーティング抽出 → カレンダー作成統合処理完了（準備中）");
                return "sample-event-id";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "ミーティング抽出 → カレンダー作成統合処理中にエラーが発生しました");
                throw new IntegrationException("ミーティング抽出・カレンダー作成に失敗しました", ex);
            }
        }

        #endregion

        #region 統合ダッシュボード機能

        /// <summary>
        /// 統合サービスの状態情報を取得する
        /// </summary>
        /// <returns>状態情報</returns>
        public async Task<IntegrationStatusInfo> GetIntegrationStatusAsync()
        {
            try
            {
                _logger.LogDebug("統合サービス状態情報取得開始");

                var status = new IntegrationStatusInfo
                {
                    Timestamp = DateTime.Now,
                    AIServiceStatus = _aiServiceManager.GetServiceStatus(),
                    AIServiceCount = _aiServiceManager.AvailableServices.Count,
                    PrimaryAIProvider = _aiServiceManager.PrimaryService?.ProviderName ?? "なし",
                    GraphServiceAvailable = _graphService.IsAvailable(),
                    ServicesOperational = _aiServiceManager.AvailableServices.Count > 0
                };

                // 各サービスの詳細テスト
                status.LastAITest = await TestAIServicesAsync();
                status.LastGraphTest = await _graphService.TestConnectionAsync();

                _logger.LogDebug("統合サービス状態情報取得完了");
                return status;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "統合サービス状態情報取得中にエラーが発生しました");
                throw new IntegrationException("統合サービス状態取得に失敗しました", ex);
            }
        }

        /// <summary>
        /// 統合機能のデモンストレーションを実行する
        /// </summary>
        /// <returns>デモ実行結果</returns>
        public async Task<string> RunIntegrationDemoAsync()
        {
            try
            {
                _logger.LogInformation("統合機能デモンストレーション開始");

                var results = new List<string>
                {
                    "=== PTA Outlook アドイン 統合機能デモ ===",
                    ""
                };

                // AI サービステスト
                results.Add("📚 AI サービステスト:");
                try
                {
                    var aiTestResult = await _aiServiceManager.TestConnectionAsync();
                    results.Add($"✅ AI サービス: 利用可能");
                    results.Add($"   プライマリ: {_aiServiceManager.PrimaryService?.ProviderName ?? "なし"}");
                    results.Add($"   利用可能数: {_aiServiceManager.AvailableServices.Count}");
                }
                catch (Exception ex)
                {
                    results.Add($"❌ AI サービス: エラー - {ex.Message}");
                }

                results.Add("");

                // Graph サービステスト
                results.Add("📊 Microsoft Graph テスト:");
                try
                {
                    var graphTestResult = await _graphService.TestConnectionAsync();
                    results.Add("🚧 Graph サービス: 準備中");
                    results.Add("   実装予定: Teams、SharePoint、Calendar 連携");
                }
                catch (Exception ex)
                {
                    results.Add($"❌ Graph サービス: エラー - {ex.Message}");
                }

                results.Add("");

                // 統合機能の説明
                results.Add("🔗 統合機能一覧:");
                results.Add("• メール解析 → Teams 共有");
                results.Add("• メール解析 → SharePoint 保存");
                results.Add("• ミーティング抽出 → カレンダー作成");
                results.Add("• AI要約 → 複数プラットフォーム配信");

                results.Add("");
                results.Add("📈 今後の拡張予定:");
                results.Add("• Power BI レポート連携");
                results.Add("• OneDrive ファイル管理");
                results.Add("• Planner タスク作成");
                results.Add("• Yammer 投稿連携");

                var finalResult = string.Join("\n", results);
                _logger.LogInformation("統合機能デモンストレーション完了");

                return finalResult;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "統合機能デモンストレーション中にエラーが発生しました");
                return $"❌ 統合機能デモ実行中にエラーが発生しました: {ex.Message}";
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// Teams メッセージ用にフォーマットする
        /// </summary>
        /// <param name="analysisResult">解析結果</param>
        /// <returns>フォーマットされたメッセージ</returns>
        private string FormatForTeams(string analysisResult)
        {
            var message = $"🤖 **AI メール解析結果**\n\n{analysisResult}\n\n" +
                         $"---\n*{DateTime.Now:yyyy/MM/dd HH:mm} PTA Outlook アドインより*";
            return message;
        }

        /// <summary>
        /// AI サービスのテストを実行する
        /// </summary>
        /// <returns>テスト結果</returns>
        private async Task<string> TestAIServicesAsync()
        {
            try
            {
                return await _aiServiceManager.TestConnectionAsync();
            }
            catch (Exception ex)
            {
                return $"AI テストエラー: {ex.Message}";
            }
        }

        #endregion
    }

    #region データモデル

    /// <summary>
    /// 統合サービス状態情報
    /// </summary>
    public class IntegrationStatusInfo
    {
        public DateTime Timestamp { get; set; }
        public string AIServiceStatus { get; set; }
        public int AIServiceCount { get; set; }
        public string PrimaryAIProvider { get; set; }
        public bool GraphServiceAvailable { get; set; }
        public bool ServicesOperational { get; set; }
        public string LastAITest { get; set; }
        public string LastGraphTest { get; set; }
    }

    /// <summary>
    /// ミーティング情報
    /// </summary>
    public class MeetingInfo
    {
        public bool IsMeetingRequest { get; set; }
        public string Subject { get; set; }
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public string Location { get; set; }
        public List<string> Attendees { get; set; } = new List<string>();
        public string Agenda { get; set; }
        public string Description { get; set; }
    }

    /// <summary>
    /// 統合サービス例外
    /// </summary>
    public class IntegrationException : Exception
    {
        public IntegrationException(string message) : base(message) { }
        public IntegrationException(string message, Exception innerException) : base(message, innerException) { }
    }

    #endregion
}
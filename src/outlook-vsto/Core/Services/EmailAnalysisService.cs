using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Outlook;
using OutlookPTAAddin.Infrastructure.OpenAI;
using OutlookPTAAddin.Core.Models;

namespace OutlookPTAAddin.Core.Services
{
    /// <summary>
    /// メール解析サービス
    /// VBAの AnalyzeSelectedEmail 機能をVSTOで再実装
    /// </summary>
    public class EmailAnalysisService
    {
        #region フィールド

        private readonly ILogger<EmailAnalysisService> _logger;
        private readonly AIServiceManager _aiServiceManager;

        // VBA版と同様のプロンプトテンプレート
        private const string SYSTEM_PROMPT = @"あなたは日本語のビジネスメール分析の専門家です。
受信したメールの内容を分析し、以下の観点から重要な情報を抽出してください：

1. 送信者の情報（会社名、部署、役職等）
2. メールの目的・要件
3. 重要なポイント・キーワード
4. 期限や日程に関する情報
5. 必要なアクション
6. 緊急度・重要度の評価

分析結果は分かりやすく整理して、日本語で回答してください。";

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="logger">ログサービス</param>
        /// <param name="aiServiceManager">AIサービス管理者</param>
        public EmailAnalysisService(ILogger<EmailAnalysisService> logger, AIServiceManager aiServiceManager)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _aiServiceManager = aiServiceManager ?? throw new ArgumentNullException(nameof(aiServiceManager));
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// メール内容を解析する
        /// </summary>
        /// <param name="mailItem">解析対象のメールアイテム</param>
        /// <returns>解析結果</returns>
        public async Task<string> AnalyzeEmailAsync(MailItem mailItem)
        {
            try
            {
                if (mailItem == null)
                {
                    throw new ArgumentNullException(nameof(mailItem), "解析対象のメールが指定されていません");
                }

                _logger.LogInformation($"メール解析開始: 件名={mailItem.Subject}");

                // メール内容の抽出
                var emailContent = ExtractEmailContent(mailItem);
                
                if (string.IsNullOrWhiteSpace(emailContent))
                {
                    throw new InvalidOperationException("メール内容が空です");
                }

                // 内容の長さチェック（VBA版と同様の制限）
                const int maxContentLength = 50000;
                if (emailContent.Length > maxContentLength)
                {
                    _logger.LogWarning($"メール内容が長すぎます（{emailContent.Length}文字）。切り詰めます。");
                    emailContent = emailContent.Substring(0, maxContentLength) + "\\n\\n[内容が切り詰められました]";
                }

                // AI APIで解析実行
                var analysisResult = await _aiServiceManager.AnalyzeTextAsync(SYSTEM_PROMPT, emailContent);

                _logger.LogInformation("メール解析完了");
                return analysisResult;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "メール解析中にエラーが発生しました");
                throw new EmailAnalysisException("メール解析に失敗しました", ex);
            }
        }

        /// <summary>
        /// 選択されたメールを解析する（VBA版のAnalyzeSelectedEmailと同等）
        /// </summary>
        /// <returns>解析結果</returns>
        public async Task<string> AnalyzeSelectedEmailAsync()
        {
            try
            {
                // アクティブなOutlookアプリケーションを取得
                var outlookApp = Globals.ThisAddIn.Application;
                
                // 選択されたアイテムを取得
                var selection = outlookApp.ActiveExplorer()?.Selection;
                
                if (selection == null || selection.Count == 0)
                {
                    throw new InvalidOperationException("メールが選択されていません。解析したいメールを選択してください。");
                }

                var selectedItem = selection[1];
                
                if (!(selectedItem is MailItem mailItem))
                {
                    throw new InvalidOperationException("選択されたアイテムはメールではありません。");
                }

                return await AnalyzeEmailAsync(mailItem);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "選択メール解析中にエラーが発生しました");
                throw;
            }
        }

        /// <summary>
        /// 検索フォルダの内容を分析する（VBA版のAnalyzeSearchFoldersと同等）
        /// </summary>
        /// <param name="folderName">検索フォルダ名</param>
        /// <returns>分析結果</returns>
        public async Task<string> AnalyzeSearchFolderAsync(string folderName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(folderName))
                {
                    throw new ArgumentException("検索フォルダ名が指定されていません", nameof(folderName));
                }

                _logger.LogInformation($"検索フォルダ分析開始: {folderName}");

                var outlookApp = Globals.ThisAddIn.Application;
                var nameSpace = outlookApp.GetNamespace("MAPI");
                
                // 検索フォルダを取得
                var searchFolder = FindSearchFolder(nameSpace, folderName);
                if (searchFolder == null)
                {
                    throw new InvalidOperationException($"検索フォルダ '{folderName}' が見つかりません");
                }

                // フォルダ内のメールを収集・分析
                var items = searchFolder.Items;
                var analysisResults = new System.Text.StringBuilder();
                
                analysisResults.AppendLine($"=== 検索フォルダ '{folderName}' の分析結果 ===");
                analysisResults.AppendLine($"メール件数: {items.Count}件");
                analysisResults.AppendLine();

                // 最大10件まで分析（パフォーマンス考慮）
                int analyzeCount = Math.Min(items.Count, 10);
                
                for (int i = 1; i <= analyzeCount; i++)
                {
                    if (items[i] is MailItem mailItem)
                    {
                        analysisResults.AppendLine($"--- メール {i} ---");
                        analysisResults.AppendLine($"件名: {mailItem.Subject}");
                        analysisResults.AppendLine($"送信者: {mailItem.SenderName}");
                        analysisResults.AppendLine($"受信日時: {mailItem.ReceivedTime}");
                        
                        try
                        {
                            var analysis = await AnalyzeEmailAsync(mailItem);
                            analysisResults.AppendLine("解析結果:");
                            analysisResults.AppendLine(analysis);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, $"メール {i} の解析に失敗しました");
                            analysisResults.AppendLine($"解析エラー: {ex.Message}");
                        }
                        
                        analysisResults.AppendLine();
                    }
                }

                if (items.Count > analyzeCount)
                {
                    analysisResults.AppendLine($"[注意] 残り{items.Count - analyzeCount}件のメールは表示されていません");
                }

                _logger.LogInformation("検索フォルダ分析完了");
                return analysisResults.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "検索フォルダ分析中にエラーが発生しました");
                throw new EmailAnalysisException("検索フォルダ分析に失敗しました", ex);
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// メールからテキスト content を抽出する
        /// </summary>
        /// <param name="mailItem">メールアイテム</param>
        /// <returns>抽出されたテキスト</returns>
        private string ExtractEmailContent(MailItem mailItem)
        {
            try
            {
                var content = new System.Text.StringBuilder();
                
                // ヘッダー情報
                content.AppendLine("=== メール情報 ===");
                content.AppendLine($"件名: {mailItem.Subject ?? "（件名なし）"}");
                content.AppendLine($"送信者: {mailItem.SenderName ?? "不明"} <{mailItem.SenderEmailAddress ?? "不明"}>");
                content.AppendLine($"受信日時: {mailItem.ReceivedTime}");
                content.AppendLine($"重要度: {GetImportanceName(mailItem.Importance)}");
                content.AppendLine();
                
                // 本文
                content.AppendLine("=== メール本文 ===");
                
                // HTMLメールの場合はプレーンテキストを優先
                if (!string.IsNullOrWhiteSpace(mailItem.Body))
                {
                    content.AppendLine(mailItem.Body);
                }
                else if (!string.IsNullOrWhiteSpace(mailItem.HTMLBody))
                {
                    // HTMLタグを簡単に除去（完璧ではないが基本的な解析には十分）
                    var plainText = System.Text.RegularExpressions.Regex.Replace(
                        mailItem.HTMLBody, 
                        "<[^>]*>", 
                        string.Empty
                    );
                    content.AppendLine(plainText);
                }
                else
                {
                    content.AppendLine("（本文なし）");
                }

                return content.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "メール内容抽出中にエラーが発生しました");
                return $"メール内容の抽出に失敗しました: {ex.Message}";
            }
        }

        /// <summary>
        /// 重要度を日本語名で取得
        /// </summary>
        /// <param name="importance">重要度</param>
        /// <returns>日本語の重要度名</returns>
        private string GetImportanceName(OlImportance importance)
        {
            return importance switch
            {
                OlImportance.olImportanceLow => "低",
                OlImportance.olImportanceNormal => "標準",
                OlImportance.olImportanceHigh => "高",
                _ => "不明"
            };
        }

        /// <summary>
        /// 検索フォルダを検索する
        /// </summary>
        /// <param name="nameSpace">MAPIネームスペース</param>
        /// <param name="folderName">フォルダ名</param>
        /// <returns>検索フォルダ（見つからない場合はnull）</returns>
        private MAPIFolder FindSearchFolder(NameSpace nameSpace, string folderName)
        {
            try
            {
                // 検索フォルダはルートの下にある
                var folders = nameSpace.Folders;
                
                foreach (MAPIFolder folder in folders)
                {
                    if (folder.Name.Equals(folderName, StringComparison.OrdinalIgnoreCase))
                    {
                        return folder;
                    }
                    
                    // サブフォルダも検索
                    var subFolder = SearchSubFolders(folder, folderName);
                    if (subFolder != null)
                    {
                        return subFolder;
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"検索フォルダ '{folderName}' の検索中にエラーが発生しました");
                return null;
            }
        }

        /// <summary>
        /// サブフォルダを再帰的に検索する
        /// </summary>
        /// <param name="parentFolder">親フォルダ</param>
        /// <param name="targetName">検索対象の名前</param>
        /// <returns>見つかったフォルダ（見つからない場合はnull）</returns>
        private MAPIFolder SearchSubFolders(MAPIFolder parentFolder, string targetName)
        {
            try
            {
                foreach (MAPIFolder subFolder in parentFolder.Folders)
                {
                    if (subFolder.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return subFolder;
                    }
                    
                    // さらにサブフォルダを検索
                    var result = SearchSubFolders(subFolder, targetName);
                    if (result != null)
                    {
                        return result;
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"サブフォルダ検索中にエラーが発生しました: {parentFolder.Name}");
                return null;
            }
        }

        #endregion
    }
}
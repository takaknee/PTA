using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Outlook;
using OutlookPTAAddin.Infrastructure.OpenAI;
using OutlookPTAAddin.Core.Models;

namespace OutlookPTAAddin.Core.Services
{
    /// <summary>
    /// メール作成サービス
    /// VBAのメール作成機能をVSTOで再実装
    /// </summary>
    public class EmailComposerService
    {
        #region フィールド

        private readonly ILogger<EmailComposerService> _logger;
        private readonly OpenAIService _openAIService;

        // VBA版と同様のプロンプトテンプレート
        private const string REJECTION_SYSTEM_PROMPT = @"あなたは丁寧で礼儀正しい日本語ビジネスメール作成の専門家です。
営業メールに対する断りの返信メールを作成してください。以下の点に注意してください：

1. 丁寧で礼儀正しい言葉遣い
2. 相手の時間を尊重する姿勢
3. 明確だが優しい断りの表現
4. 今後の関係性を損なわない配慮
5. 適切な敬語の使用

断りの理由は一般的なものを使用し、相手を不快にさせないよう配慮してください。";

        private const string ACCEPTANCE_SYSTEM_PROMPT = @"あなたは丁寧で礼儀正しい日本語ビジネスメール作成の専門家です。
提案やお誘いに対する承諾の返信メールを作成してください。以下の点に注意してください：

1. 感謝の気持ちを表現
2. 前向きで積極的な姿勢
3. 具体的な対応や次のステップの提示
4. 適切な敬語の使用
5. 相手への敬意と配慮

承諾の内容は提案に応じて適切に調整してください。";

        private const string CUSTOM_SYSTEM_PROMPT = @"あなたは日本語のビジネスメール作成の専門家です。
指定された要件に基づいて、適切で丁寧なビジネスメールを作成してください。以下の点に注意してください：

1. 適切な敬語の使用
2. ビジネスメールの基本構造（宛先、挨拶、本文、結び、署名）
3. 読みやすい文章構成
4. 相手への配慮と礼儀
5. 目的に応じた適切なトーン";

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="logger">ログサービス</param>
        /// <param name="openAIService">OpenAI サービス</param>
        public EmailComposerService(ILogger<EmailComposerService> logger, OpenAIService openAIService)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _openAIService = openAIService ?? throw new ArgumentNullException(nameof(openAIService));
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// 営業断りメールを作成する（VBA版のCreateRejectionEmailと同等）
        /// </summary>
        /// <param name="originalMail">元のメール（返信元）</param>
        /// <returns>作成されたメールアイテム</returns>
        public async Task<MailItem> CreateRejectionEmailAsync(MailItem originalMail = null)
        {
            try
            {
                _logger.LogInformation("営業断りメール作成開始");

                string originalContent = "";
                string replyTo = "";
                string subject = "お問い合わせの件について";

                // 元メールがある場合は内容を取得
                if (originalMail != null)
                {
                    originalContent = ExtractEmailContent(originalMail);
                    replyTo = originalMail.SenderEmailAddress ?? "";
                    subject = "Re: " + (originalMail.Subject ?? "");
                }

                // プロンプトの準備
                var userPrompt = string.IsNullOrEmpty(originalContent)
                    ? "一般的な営業メールに対する丁寧な断りメールを作成してください。"
                    : $"以下のメールに対する丁寧な断りの返信メールを作成してください：\\n\\n{originalContent}";

                // OpenAI APIでメール本文を生成
                var emailBody = await _openAIService.GenerateTextAsync(REJECTION_SYSTEM_PROMPT, userPrompt);

                // 新しいメールアイテムを作成
                var outlookApp = Globals.ThisAddIn.Application;
                var newMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                // メール情報を設定
                if (!string.IsNullOrEmpty(replyTo))
                {
                    newMail.To = replyTo;
                }
                newMail.Subject = subject;
                newMail.Body = emailBody;

                // メールを表示（送信はユーザーが手動で行う）
                newMail.Display(false);

                _logger.LogInformation("営業断りメール作成完了");
                return newMail;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "営業断りメール作成中にエラーが発生しました");
                throw new EmailCompositionException("営業断りメール作成に失敗しました", ex);
            }
        }

        /// <summary>
        /// 承諾メールを作成する（VBA版のCreateAcceptanceEmailと同等）
        /// </summary>
        /// <param name="originalMail">元のメール（返信元）</param>
        /// <returns>作成されたメールアイテム</returns>
        public async Task<MailItem> CreateAcceptanceEmailAsync(MailItem originalMail = null)
        {
            try
            {
                _logger.LogInformation("承諾メール作成開始");

                string originalContent = "";
                string replyTo = "";
                string subject = "ご提案の件について";

                // 元メールがある場合は内容を取得
                if (originalMail != null)
                {
                    originalContent = ExtractEmailContent(originalMail);
                    replyTo = originalMail.SenderEmailAddress ?? "";
                    subject = "Re: " + (originalMail.Subject ?? "");
                }

                // プロンプトの準備
                var userPrompt = string.IsNullOrEmpty(originalContent)
                    ? "一般的な提案やお誘いに対する承諾メールを作成してください。"
                    : $"以下のメールに対する承諾の返信メールを作成してください：\\n\\n{originalContent}";

                // OpenAI APIでメール本文を生成
                var emailBody = await _openAIService.GenerateTextAsync(ACCEPTANCE_SYSTEM_PROMPT, userPrompt);

                // 新しいメールアイテムを作成
                var outlookApp = Globals.ThisAddIn.Application;
                var newMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                // メール情報を設定
                if (!string.IsNullOrEmpty(replyTo))
                {
                    newMail.To = replyTo;
                }
                newMail.Subject = subject;
                newMail.Body = emailBody;

                // メールを表示（送信はユーザーが手動で行う）
                newMail.Display(false);

                _logger.LogInformation("承諾メール作成完了");
                return newMail;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "承諾メール作成中にエラーが発生しました");
                throw new EmailCompositionException("承諾メール作成に失敗しました", ex);
            }
        }

        /// <summary>
        /// カスタムメールを作成する（VBA版のCreateCustomBusinessEmailと同等）
        /// </summary>
        /// <param name="customPrompt">カスタムプロンプト</param>
        /// <param name="originalMail">元のメール（オプション）</param>
        /// <returns>作成されたメールアイテム</returns>
        public async Task<MailItem> CreateCustomEmailAsync(string customPrompt, MailItem originalMail = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(customPrompt))
                {
                    throw new ArgumentException("カスタムプロンプトが指定されていません", nameof(customPrompt));
                }

                _logger.LogInformation("カスタムメール作成開始");

                string originalContent = "";
                string replyTo = "";
                string subject = "お問い合わせの件";

                // 元メールがある場合は内容を取得
                if (originalMail != null)
                {
                    originalContent = ExtractEmailContent(originalMail);
                    replyTo = originalMail.SenderEmailAddress ?? "";
                    subject = "Re: " + (originalMail.Subject ?? "");
                }

                // プロンプトの準備
                var userPrompt = string.IsNullOrEmpty(originalContent)
                    ? customPrompt
                    : $"{customPrompt}\\n\\n参考となる元メール：\\n{originalContent}";

                // OpenAI APIでメール本文を生成
                var emailBody = await _openAIService.GenerateTextAsync(CUSTOM_SYSTEM_PROMPT, userPrompt);

                // 新しいメールアイテムを作成
                var outlookApp = Globals.ThisAddIn.Application;
                var newMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                // メール情報を設定
                if (!string.IsNullOrEmpty(replyTo))
                {
                    newMail.To = replyTo;
                }
                newMail.Subject = subject;
                newMail.Body = emailBody;

                // メールを表示（送信はユーザーが手動で行う）
                newMail.Display(false);

                _logger.LogInformation("カスタムメール作成完了");
                return newMail;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "カスタムメール作成中にエラーが発生しました");
                throw new EmailCompositionException("カスタムメール作成に失敗しました", ex);
            }
        }

        /// <summary>
        /// 選択されたメールへの営業断り返信を作成する
        /// </summary>
        /// <returns>作成されたメールアイテム</returns>
        public async Task<MailItem> CreateRejectionReplyToSelectedAsync()
        {
            try
            {
                var selectedMail = GetSelectedMail();
                if (selectedMail == null)
                {
                    throw new InvalidOperationException("メールが選択されていません。返信したいメールを選択してください。");
                }

                return await CreateRejectionEmailAsync(selectedMail);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "選択メールへの営業断り返信作成中にエラーが発生しました");
                throw;
            }
        }

        /// <summary>
        /// 選択されたメールへの承諾返信を作成する
        /// </summary>
        /// <returns>作成されたメールアイテム</returns>
        public async Task<MailItem> CreateAcceptanceReplyToSelectedAsync()
        {
            try
            {
                var selectedMail = GetSelectedMail();
                if (selectedMail == null)
                {
                    throw new InvalidOperationException("メールが選択されていません。返信したいメールを選択してください。");
                }

                return await CreateAcceptanceEmailAsync(selectedMail);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "選択メールへの承諾返信作成中にエラーが発生しました");
                throw;
            }
        }

        /// <summary>
        /// 選択されたメールへのカスタム返信を作成する
        /// </summary>
        /// <param name="customPrompt">カスタムプロンプト</param>
        /// <returns>作成されたメールアイテム</returns>
        public async Task<MailItem> CreateCustomReplyToSelectedAsync(string customPrompt)
        {
            try
            {
                var selectedMail = GetSelectedMail();
                return await CreateCustomEmailAsync(customPrompt, selectedMail);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "選択メールへのカスタム返信作成中にエラーが発生しました");
                throw;
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// 現在選択されているメールを取得する
        /// </summary>
        /// <returns>選択されたメール（選択されていない場合はnull）</returns>
        private MailItem GetSelectedMail()
        {
            try
            {
                var outlookApp = Globals.ThisAddIn.Application;
                var selection = outlookApp.ActiveExplorer()?.Selection;
                
                if (selection != null && selection.Count > 0)
                {
                    var selectedItem = selection[1];
                    return selectedItem as MailItem;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "選択メール取得中にエラーが発生しました");
                return null;
            }
        }

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
                content.AppendLine("=== 元メール情報 ===");
                content.AppendLine($"件名: {mailItem.Subject ?? "（件名なし）"}");
                content.AppendLine($"送信者: {mailItem.SenderName ?? "不明"}");
                content.AppendLine($"受信日時: {mailItem.ReceivedTime}");
                content.AppendLine();
                
                // 本文（最初の1000文字まで）
                content.AppendLine("=== 元メール本文 ===");
                var body = mailItem.Body ?? "";
                if (body.Length > 1000)
                {
                    body = body.Substring(0, 1000) + "...";
                }
                content.AppendLine(body);

                return content.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "メール内容抽出中にエラーが発生しました");
                return $"メール内容の抽出に失敗しました: {ex.Message}";
            }
        }

        #endregion
    }
}
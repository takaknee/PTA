using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OutlookPTAAddin.Core.Services;
using OutlookPTAAddin.UI.Dialogs;

namespace OutlookPTAAddin.UI.Ribbon
{
    /// <summary>
    /// PTA Outlook アドイン リボンUI
    /// </summary>
    [ComVisible(true)]
    public partial class PTARibbon : RibbonBase
    {
        #region フィールド

        private ILogger<PTARibbon> _logger;
        private EmailAnalysisService _emailAnalysisService;
        private EmailComposerService _emailComposerService;
        private AIServiceManager _aiServiceManager;

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        public PTARibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        #endregion

        #region イベントハンドラー

        /// <summary>
        /// リボン読み込み時の処理
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private void PTARibbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                // サービスの取得
                var serviceProvider = Globals.ThisAddIn.ServiceProvider;
                _logger = serviceProvider?.GetService<ILogger<PTARibbon>>();
                _emailAnalysisService = serviceProvider?.GetService<EmailAnalysisService>();
                _emailComposerService = serviceProvider?.GetService<EmailComposerService>();
                _aiServiceManager = serviceProvider?.GetService<AIServiceManager>();

                _logger?.LogInformation("PTA リボンUIが読み込まれました");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"リボンUI初期化エラー: {ex.Message}", "PTA Outlook アドイン", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// メール解析ボタンクリック
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private async void btnAnalyzeEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _logger?.LogInformation("メール解析ボタンがクリックされました");

                if (_emailAnalysisService == null)
                {
                    MessageBox.Show("メール解析サービスが利用できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // プログレス表示付きで解析実行
                await ProgressDialog.RunWithProgressAsync(
                    null,
                    async (dialog, cancellationToken) =>
                    {
                        dialog.UpdateStatus("選択されたメールを解析中...", "AI APIにリクエストを送信しています");
                        var result = await _emailAnalysisService.AnalyzeSelectedEmailAsync();
                        
                        dialog.UpdateStatus("解析結果を表示中...");
                        // 結果表示ダイアログを表示
                        var resultForm = new ResultDisplayForm("メール解析結果", result);
                        resultForm.ShowDialog();
                    },
                    "メール解析",
                    "メールの内容を解析しています...",
                    true
                );
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "メール解析処理中にエラーが発生しました");
                MessageBox.Show($"メール解析中にエラーが発生しました:\\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 営業断りメール作成ボタンクリック
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private async void btnCreateRejectionEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _logger?.LogInformation("営業断りメール作成ボタンがクリックされました");

                if (_emailComposerService == null)
                {
                    MessageBox.Show("メール作成サービスが利用できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // プログレス表示
                var progressForm = new ProgressForm("営業断りメールを作成中...");
                progressForm.Show();

                try
                {
                    await _emailComposerService.CreateRejectionReplyToSelectedAsync();
                    progressForm.Close();
                    
                    MessageBox.Show("営業断りメールが作成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    progressForm?.Close();
                }
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "営業断りメール作成中にエラーが発生しました");
                MessageBox.Show($"営業断りメール作成中にエラーが発生しました:\\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 承諾メール作成ボタンクリック
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private async void btnCreateAcceptanceEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _logger?.LogInformation("承諾メール作成ボタンがクリックされました");

                if (_emailComposerService == null)
                {
                    MessageBox.Show("メール作成サービスが利用できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // プログレス表示
                var progressForm = new ProgressForm("承諾メールを作成中...");
                progressForm.Show();

                try
                {
                    await _emailComposerService.CreateAcceptanceReplyToSelectedAsync();
                    progressForm.Close();
                    
                    MessageBox.Show("承諾メールが作成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    progressForm?.Close();
                }
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "承諾メール作成中にエラーが発生しました");
                MessageBox.Show($"承諾メール作成中にエラーが発生しました:\\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// カスタムメール作成ボタンクリック
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private async void btnCreateCustomEmail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _logger?.LogInformation("カスタムメール作成ボタンがクリックされました");

                if (_emailComposerService == null)
                {
                    MessageBox.Show("メール作成サービスが利用できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // カスタムプロンプトの入力
                var promptForm = new CustomPromptForm();
                if (promptForm.ShowDialog() == DialogResult.OK)
                {
                    var customPrompt = promptForm.CustomPrompt;
                    
                    if (string.IsNullOrWhiteSpace(customPrompt))
                    {
                        MessageBox.Show("カスタムプロンプトが入力されていません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // プログレス表示
                    var progressForm = new ProgressForm("カスタムメールを作成中...");
                    progressForm.Show();

                    try
                    {
                        await _emailComposerService.CreateCustomReplyToSelectedAsync(customPrompt);
                        progressForm.Close();
                        
                        MessageBox.Show("カスタムメールが作成されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    finally
                    {
                        progressForm?.Close();
                    }
                }
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "カスタムメール作成中にエラーが発生しました");
                MessageBox.Show($"カスタムメール作成中にエラーが発生しました:\\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 設定ボタンクリック
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _logger?.LogInformation("設定ボタンがクリックされました");

                if (_aiServiceManager == null)
                {
                    MessageBox.Show("AI サービス管理者が利用できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var settingsLogger = Globals.ThisAddIn.ServiceProvider?.GetService<ILogger<SettingsDialog>>();
                var settingsDialog = new SettingsDialog(settingsLogger, _aiServiceManager);
                
                if (settingsDialog.ShowDialog() == DialogResult.OK)
                {
                    MessageBox.Show("設定が保存されました。", "設定完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "設定画面表示中にエラーが発生しました");
                MessageBox.Show($"設定画面表示中にエラーが発生しました:\\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// バージョン情報ボタンクリック
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                _logger?.LogInformation("バージョン情報ボタンがクリックされました");

                var aboutForm = new AboutForm();
                aboutForm.ShowDialog();
            }
            catch (System.Exception ex)
            {
                _logger?.LogError(ex, "バージョン情報表示中にエラーが発生しました");
                MessageBox.Show($"バージョン情報表示中にエラーが発生しました:\\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// 部分的なメソッド実装（デザイナー用）
        /// </summary>
        partial void InitializeComponent();

        #endregion
    }

    #region 簡単なフォームクラス

    /// <summary>
    /// プログレス表示フォーム
    /// </summary>
    public class ProgressForm : Form
    {
        public ProgressForm(string message)
        {
            Text = "処理中";
            Size = new System.Drawing.Size(300, 100);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;

            var label = new Label
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            Controls.Add(label);
        }
    }

    /// <summary>
    /// 結果表示フォーム
    /// </summary>
    public class ResultDisplayForm : Form
    {
        public ResultDisplayForm(string title, string content)
        {
            Text = title;
            Size = new System.Drawing.Size(600, 400);
            StartPosition = FormStartPosition.CenterParent;

            var textBox = new TextBox
            {
                Text = content,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Dock = DockStyle.Fill
            };

            Controls.Add(textBox);
        }
    }

    /// <summary>
    /// カスタムプロンプト入力フォーム
    /// </summary>
    public class CustomPromptForm : Form
    {
        public string CustomPrompt { get; private set; }

        public CustomPromptForm()
        {
            Text = "カスタムプロンプト入力";
            Size = new System.Drawing.Size(500, 300);
            StartPosition = FormStartPosition.CenterParent;

            var label = new Label
            {
                Text = "作成したいメールの内容を詳しく記述してください:",
                Location = new System.Drawing.Point(10, 10),
                Size = new System.Drawing.Size(460, 20)
            };

            var textBox = new TextBox
            {
                Location = new System.Drawing.Point(10, 35),
                Size = new System.Drawing.Size(460, 150),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };

            var btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(315, 200),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new System.Drawing.Point(395, 200),
                DialogResult = DialogResult.Cancel
            };

            btnOK.Click += (s, e) => { CustomPrompt = textBox.Text; };

            Controls.AddRange(new Control[] { label, textBox, btnOK, btnCancel });
        }
    }

    /// <summary>
    /// 設定フォーム（プレースホルダー）
    /// </summary>
    public class SettingsForm : Form
    {
        public SettingsForm()
        {
            Text = "設定";
            Size = new System.Drawing.Size(400, 300);
            StartPosition = FormStartPosition.CenterParent;

            var label = new Label
            {
                Text = "設定画面（将来実装予定）",
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            Controls.Add(label);
        }
    }

    /// <summary>
    /// バージョン情報フォーム
    /// </summary>
    public class AboutForm : Form
    {
        public AboutForm()
        {
            Text = "バージョン情報";
            Size = new System.Drawing.Size(350, 200);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var info = $"PTA Outlook アドイン\\n\\nバージョン: 1.0.0 VSTO\\n作成日: 2024年\\n\\n" +
                      $"VSTOで再実装されたOutlook AI Helper\\n\\n主要機能:\\n" +
                      $"・メール内容解析\\n・営業断りメール作成\\n・承諾メール作成\\n・カスタムメール作成";

            var label = new Label
            {
                Text = info,
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(300, 120),
                TextAlign = System.Drawing.ContentAlignment.TopLeft
            };

            var btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(135, 150),
                DialogResult = DialogResult.OK
            };

            Controls.AddRange(new Control[] { label, btnOK });
        }
    }

    #endregion
}
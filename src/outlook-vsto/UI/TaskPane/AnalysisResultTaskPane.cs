using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;
using OutlookPTAAddin.Core.Services;

namespace OutlookPTAAddin.UI.TaskPane
{
    /// <summary>
    /// メール解析結果表示用タスクペイン
    /// AI分析結果を見やすい形で表示
    /// </summary>
    public partial class AnalysisResultTaskPane : UserControl
    {
        #region フィールド

        private readonly ILogger<AnalysisResultTaskPane> _logger;
        private readonly AIServiceManager _aiServiceManager;

        // UI コントロール
        private Panel _headerPanel;
        private Label _titleLabel;
        private Button _refreshButton;
        private Button _copyButton;
        private Button _clearButton;
        
        private Panel _contentPanel;
        private RichTextBox _resultTextBox;
        private Panel _statusPanel;
        private Label _statusLabel;
        private ProgressBar _progressBar;

        private Panel _actionPanel;
        private Button _analyzeCurrentButton;
        private Button _settingsButton;
        private ComboBox _aiProviderComboBox;

        private bool _isAnalyzing = false;

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="logger">ログサービス</param>
        /// <param name="aiServiceManager">AIサービス管理者</param>
        public AnalysisResultTaskPane(ILogger<AnalysisResultTaskPane> logger, AIServiceManager aiServiceManager)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _aiServiceManager = aiServiceManager ?? throw new ArgumentNullException(nameof(aiServiceManager));

            InitializeComponent();
            UpdateProviderComboBox();
            UpdateStatus("解析結果はここに表示されます");
        }

        #endregion

        #region コンポーネント初期化

        /// <summary>
        /// コンポーネントの初期化
        /// </summary>
        private void InitializeComponent()
        {
            // メインコントロールの設定
            Dock = DockStyle.Fill;
            BackColor = SystemColors.Control;
            Padding = new Padding(5);

            // ヘッダーパネルの作成
            CreateHeaderPanel();
            
            // コンテンツパネルの作成
            CreateContentPanel();
            
            // ステータスパネルの作成
            CreateStatusPanel();
            
            // アクションパネルの作成
            CreateActionPanel();

            // レイアウト調整
            AdjustLayout();
        }

        /// <summary>
        /// ヘッダーパネルの作成
        /// </summary>
        private void CreateHeaderPanel()
        {
            _headerPanel = new Panel
            {
                Height = 40,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(240, 240, 240),
                Padding = new Padding(5)
            };

            _titleLabel = new Label
            {
                Text = "🤖 AI メール解析",
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(50, 50, 50),
                Location = new Point(5, 10),
                AutoSize = true
            };
            _headerPanel.Controls.Add(_titleLabel);

            _refreshButton = new Button
            {
                Text = "🔄",
                Size = new Size(30, 25),
                Location = new Point(_headerPanel.Width - 100, 7),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                UseVisualStyleBackColor = true,
                ToolTipText = "再解析"
            };
            _refreshButton.Click += RefreshButton_Click;
            _headerPanel.Controls.Add(_refreshButton);

            _copyButton = new Button
            {
                Text = "📋",
                Size = new Size(30, 25),
                Location = new Point(_headerPanel.Width - 65, 7),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                UseVisualStyleBackColor = true,
                ToolTipText = "結果をコピー"
            };
            _copyButton.Click += CopyButton_Click;
            _headerPanel.Controls.Add(_copyButton);

            _clearButton = new Button
            {
                Text = "🗑️",
                Size = new Size(30, 25),
                Location = new Point(_headerPanel.Width - 35, 7),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                UseVisualStyleBackColor = true,
                ToolTipText = "クリア"
            };
            _clearButton.Click += ClearButton_Click;
            _headerPanel.Controls.Add(_clearButton);

            Controls.Add(_headerPanel);
        }

        /// <summary>
        /// コンテンツパネルの作成
        /// </summary>
        private void CreateContentPanel()
        {
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            _resultTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Meiryo UI", 9),
                WordWrap = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Text = "解析結果はここに表示されます。\n\n" +
                       "使用方法:\n" +
                       "1. 解析したいメールを選択\n" +
                       "2. 「現在のメールを解析」ボタンをクリック\n" +
                       "3. AI分析結果がここに表示されます\n\n" +
                       "複数のAIプロバイダーが利用可能な場合、\n" +
                       "自動でフォールバック機能が働きます。"
            };
            _contentPanel.Controls.Add(_resultTextBox);

            Controls.Add(_contentPanel);
        }

        /// <summary>
        /// ステータスパネルの作成
        /// </summary>
        private void CreateStatusPanel()
        {
            _statusPanel = new Panel
            {
                Height = 30,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(250, 250, 250),
                Padding = new Padding(5, 2, 5, 2)
            };

            _statusLabel = new Label
            {
                Text = "準備完了",
                Location = new Point(5, 5),
                AutoSize = true,
                ForeColor = Color.FromArgb(80, 80, 80),
                Font = new Font(Font.FontFamily, 8)
            };
            _statusPanel.Controls.Add(_statusLabel);

            _progressBar = new ProgressBar
            {
                Size = new Size(100, 15),
                Location = new Point(_statusPanel.Width - 110, 7),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 50,
                Visible = false
            };
            _statusPanel.Controls.Add(_progressBar);

            Controls.Add(_statusPanel);
        }

        /// <summary>
        /// アクションパネルの作成
        /// </summary>
        private void CreateActionPanel()
        {
            _actionPanel = new Panel
            {
                Height = 80,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(5)
            };

            // AIプロバイダー選択
            var lblProvider = new Label
            {
                Text = "AIプロバイダー:",
                Location = new Point(5, 8),
                Size = new Size(80, 23),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(Font.FontFamily, 8)
            };
            _actionPanel.Controls.Add(lblProvider);

            _aiProviderComboBox = new ComboBox
            {
                Location = new Point(90, 5),
                Size = new Size(120, 23),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font(Font.FontFamily, 8)
            };
            _aiProviderComboBox.SelectedIndexChanged += AiProviderComboBox_SelectedIndexChanged;
            _actionPanel.Controls.Add(_aiProviderComboBox);

            // 現在のメール解析ボタン
            _analyzeCurrentButton = new Button
            {
                Text = "📧 現在のメールを解析",
                Location = new Point(5, 35),
                Size = new Size(140, 30),
                UseVisualStyleBackColor = true,
                Font = new Font(Font.FontFamily, 9)
            };
            _analyzeCurrentButton.Click += AnalyzeCurrentButton_Click;
            _actionPanel.Controls.Add(_analyzeCurrentButton);

            // 設定ボタン
            _settingsButton = new Button
            {
                Text = "⚙️ 設定",
                Location = new Point(150, 35),
                Size = new Size(60, 30),
                UseVisualStyleBackColor = true,
                Font = new Font(Font.FontFamily, 9)
            };
            _settingsButton.Click += SettingsButton_Click;
            _actionPanel.Controls.Add(_settingsButton);

            Controls.Add(_actionPanel);
        }

        /// <summary>
        /// レイアウトの調整
        /// </summary>
        private void AdjustLayout()
        {
            // コンテンツパネルがボタンパネルとステータスパネルの間に配置されるように調整
            _contentPanel.Height = Height - _headerPanel.Height - _actionPanel.Height - _statusPanel.Height - Padding.Vertical;
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// 解析結果を表示する
        /// </summary>
        /// <param name="result">解析結果</param>
        /// <param name="provider">使用されたAIプロバイダー</param>
        public void DisplayAnalysisResult(string result, string provider = null)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<string, string>(DisplayAnalysisResult), result, provider);
                    return;
                }

                _resultTextBox.Clear();

                // ヘッダー情報
                var header = $"=== AI メール解析結果 ===\n";
                if (!string.IsNullOrWhiteSpace(provider))
                {
                    header += $"プロバイダー: {provider}\n";
                }
                header += $"解析日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}\n\n";

                // ヘッダーを太字で表示
                _resultTextBox.SelectionFont = new Font(_resultTextBox.Font, FontStyle.Bold);
                _resultTextBox.SelectionColor = Color.DarkBlue;
                _resultTextBox.AppendText(header);

                // 解析結果を通常フォントで表示
                _resultTextBox.SelectionFont = new Font(_resultTextBox.Font, FontStyle.Regular);
                _resultTextBox.SelectionColor = Color.Black;
                _resultTextBox.AppendText(result);

                // 最後に改行を追加
                _resultTextBox.AppendText("\n\n");

                // 先頭にスクロール
                _resultTextBox.SelectionStart = 0;
                _resultTextBox.ScrollToCaret();

                UpdateStatus($"解析完了 - {provider ?? "AI"}");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "解析結果表示中にエラーが発生しました");
                DisplayError("解析結果の表示中にエラーが発生しました", ex.Message);
            }
        }

        /// <summary>
        /// エラーメッセージを表示する
        /// </summary>
        /// <param name="message">エラーメッセージ</param>
        /// <param name="detail">詳細情報</param>
        public void DisplayError(string message, string detail = null)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<string, string>(DisplayError), message, detail);
                    return;
                }

                _resultTextBox.Clear();

                // エラーヘッダー
                _resultTextBox.SelectionFont = new Font(_resultTextBox.Font, FontStyle.Bold);
                _resultTextBox.SelectionColor = Color.Red;
                _resultTextBox.AppendText("❌ エラーが発生しました\n\n");

                // エラーメッセージ
                _resultTextBox.SelectionFont = new Font(_resultTextBox.Font, FontStyle.Regular);
                _resultTextBox.SelectionColor = Color.DarkRed;
                _resultTextBox.AppendText($"メッセージ: {message}\n");

                if (!string.IsNullOrWhiteSpace(detail))
                {
                    _resultTextBox.AppendText($"\n詳細:\n{detail}\n");
                }

                _resultTextBox.AppendText("\n対処方法:\n");
                _resultTextBox.AppendText("1. インターネット接続を確認してください\n");
                _resultTextBox.AppendText("2. API設定を確認してください\n");
                _resultTextBox.AppendText("3. しばらく時間をおいて再試行してください\n");

                UpdateStatus("エラー");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "エラー表示中にエラーが発生しました");
            }
        }

        /// <summary>
        /// 解析中の状態を表示する
        /// </summary>
        /// <param name="message">進行状況メッセージ</param>
        public void ShowAnalyzing(string message = "メールを解析中...")
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<string>(ShowAnalyzing), message);
                    return;
                }

                _isAnalyzing = true;
                _analyzeCurrentButton.Enabled = false;
                _progressBar.Visible = true;

                _resultTextBox.Clear();
                _resultTextBox.SelectionFont = new Font(_resultTextBox.Font, FontStyle.Italic);
                _resultTextBox.SelectionColor = Color.Gray;
                _resultTextBox.AppendText($"🔄 {message}\n\nしばらくお待ちください...");

                UpdateStatus(message);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "解析中表示でエラーが発生しました");
            }
        }

        /// <summary>
        /// 解析完了状態にリセットする
        /// </summary>
        public void ResetAnalyzingState()
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(ResetAnalyzingState));
                    return;
                }

                _isAnalyzing = false;
                _analyzeCurrentButton.Enabled = true;
                _progressBar.Visible = false;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "解析状態リセット中にエラーが発生しました");
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// ステータスを更新する
        /// </summary>
        /// <param name="message">ステータスメッセージ</param>
        private void UpdateStatus(string message)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<string>(UpdateStatus), message);
                    return;
                }

                if (_statusLabel != null)
                {
                    _statusLabel.Text = message;
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "ステータス更新中にエラーが発生しました");
            }
        }

        /// <summary>
        /// AIプロバイダーコンボボックスを更新する
        /// </summary>
        private void UpdateProviderComboBox()
        {
            try
            {
                if (_aiProviderComboBox == null) return;

                _aiProviderComboBox.Items.Clear();
                _aiProviderComboBox.Items.Add("自動選択");

                foreach (var service in _aiServiceManager.RegisteredServices)
                {
                    _aiProviderComboBox.Items.Add(service.ProviderName);
                }

                _aiProviderComboBox.SelectedIndex = 0; // 自動選択
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "プロバイダーコンボボックス更新中にエラーが発生しました");
            }
        }

        #endregion

        #region イベントハンドラー

        /// <summary>
        /// 現在のメール解析ボタンクリック
        /// </summary>
        private async void AnalyzeCurrentButton_Click(object sender, EventArgs e)
        {
            if (_isAnalyzing) return;

            try
            {
                ShowAnalyzing("選択されたメールを解析中...");

                // EmailAnalysisServiceを取得してメール解析を実行
                var emailAnalysisService = Globals.ThisAddIn.ServiceProvider?.GetService(typeof(EmailAnalysisService)) as EmailAnalysisService;
                
                if (emailAnalysisService == null)
                {
                    DisplayError("メール解析サービスが利用できません");
                    return;
                }

                var result = await emailAnalysisService.AnalyzeSelectedEmailAsync();
                var provider = _aiServiceManager.PrimaryService?.ProviderName ?? "AI";
                
                DisplayAnalysisResult(result, provider);
            }
            catch (InvalidOperationException ex)
            {
                DisplayError("メール選択エラー", ex.Message);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "メール解析中にエラーが発生しました");
                DisplayError("メール解析中にエラーが発生しました", ex.Message);
            }
            finally
            {
                ResetAnalyzingState();
            }
        }

        /// <summary>
        /// 再解析ボタンクリック
        /// </summary>
        private void RefreshButton_Click(object sender, EventArgs e)
        {
            AnalyzeCurrentButton_Click(sender, e);
        }

        /// <summary>
        /// コピーボタンクリック
        /// </summary>
        private void CopyButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(_resultTextBox.Text))
                {
                    Clipboard.SetText(_resultTextBox.Text);
                    UpdateStatus("結果をクリップボードにコピーしました");
                }
                else
                {
                    UpdateStatus("コピーする内容がありません");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "クリップボードコピー中にエラーが発生しました");
                UpdateStatus("コピーに失敗しました");
            }
        }

        /// <summary>
        /// クリアボタンクリック
        /// </summary>
        private void ClearButton_Click(object sender, EventArgs e)
        {
            try
            {
                _resultTextBox.Clear();
                _resultTextBox.Text = "解析結果はここに表示されます。";
                UpdateStatus("結果をクリアしました");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "結果クリア中にエラーが発生しました");
            }
        }

        /// <summary>
        /// AIプロバイダー選択変更
        /// </summary>
        private void AiProviderComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (_aiProviderComboBox.SelectedIndex <= 0) return; // 自動選択の場合は何もしない

                var selectedProvider = _aiProviderComboBox.SelectedItem.ToString();
                if (_aiServiceManager.SetPrimaryService(selectedProvider))
                {
                    UpdateStatus($"プライマリサービスを {selectedProvider} に変更しました");
                }
                else
                {
                    UpdateStatus($"{selectedProvider} の設定に失敗しました");
                    UpdateProviderComboBox(); // コンボボックスをリセット
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "プロバイダー変更中にエラーが発生しました");
                UpdateStatus("プロバイダーの変更に失敗しました");
            }
        }

        /// <summary>
        /// 設定ボタンクリック
        /// </summary>
        private void SettingsButton_Click(object sender, EventArgs e)
        {
            try
            {
                // 設定ダイアログを表示
                var logger = Globals.ThisAddIn.ServiceProvider?.GetService(typeof(ILogger<Dialogs.SettingsDialog>)) as ILogger<Dialogs.SettingsDialog>;
                var settingsDialog = new Dialogs.SettingsDialog(logger, _aiServiceManager);
                
                if (settingsDialog.ShowDialog() == DialogResult.OK)
                {
                    UpdateProviderComboBox();
                    UpdateStatus("設定が更新されました");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "設定ダイアログ表示中にエラーが発生しました");
                UpdateStatus("設定ダイアログの表示に失敗しました");
            }
        }

        /// <summary>
        /// サイズ変更時の処理
        /// </summary>
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            AdjustLayout();
        }

        #endregion
    }
}
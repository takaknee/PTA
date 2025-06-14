using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;
using OutlookPTAAddin.Core.Services;

namespace OutlookPTAAddin.UI.Dialogs
{
    /// <summary>
    /// 改良された設定ダイアログフォーム
    /// 日本語UIと直感的な操作を提供
    /// </summary>
    public partial class SettingsDialog : Form
    {
        #region フィールド

        private readonly ILogger<SettingsDialog> _logger;
        private readonly AIServiceManager _aiServiceManager;
        private bool _hasUnsavedChanges = false;

        // UI コントロール
        private TabControl tabControl;
        private TabPage tabAIServices;
        private TabPage tabGeneral;
        private TabPage tabAdvanced;
        
        // AIサービス設定
        private ComboBox cmbPrimaryService;
        private TextBox txtOpenAIKey;
        private TextBox txtOpenAIEndpoint;
        private TextBox txtClaudeKey;
        private Button btnTestConnection;
        private ListBox lstServiceStatus;
        
        // 一般設定
        private NumericUpDown numMaxTokens;
        private NumericUpDown numTimeout;
        private CheckBox chkShowProgress;
        private CheckBox chkAutoUpdate;
        
        // 高度な設定
        private CheckBox chkEnableLogging;
        private ComboBox cmbLogLevel;
        private TextBox txtLogPath;
        private Button btnResetSettings;

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="logger">ログサービス</param>
        /// <param name="aiServiceManager">AIサービス管理者</param>
        public SettingsDialog(ILogger<SettingsDialog> logger, AIServiceManager aiServiceManager)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _aiServiceManager = aiServiceManager ?? throw new ArgumentNullException(nameof(aiServiceManager));

            InitializeComponent();
            LoadSettings();
        }

        #endregion

        #region フォーム初期化

        /// <summary>
        /// コンポーネントの初期化
        /// </summary>
        private void InitializeComponent()
        {
            Text = "PTA Outlook アドイン - 設定";
            Size = new Size(600, 500);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;

            // タブコントロールの作成
            CreateTabControl();
            
            // ボタンの作成
            CreateButtons();

            // レイアウト調整
            AdjustLayout();
        }

        /// <summary>
        /// タブコントロールの作成
        /// </summary>
        private void CreateTabControl()
        {
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(10)
            };

            // AIサービス設定タブ
            CreateAIServicesTab();
            
            // 一般設定タブ
            CreateGeneralTab();
            
            // 高度な設定タブ
            CreateAdvancedTab();

            Controls.Add(tabControl);
        }

        /// <summary>
        /// AIサービス設定タブの作成
        /// </summary>
        private void CreateAIServicesTab()
        {
            tabAIServices = new TabPage("🤖 AIサービス");
            
            var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            
            var y = 10;
            
            // プライマリサービス選択
            var lblPrimary = new Label
            {
                Text = "プライマリAIサービス:",
                Location = new Point(10, y),
                Size = new Size(150, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblPrimary);
            
            cmbPrimaryService = new ComboBox
            {
                Location = new Point(170, y),
                Size = new Size(200, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbPrimaryService.SelectedIndexChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(cmbPrimaryService);
            
            y += 35;

            // OpenAI設定
            var lblOpenAI = new Label
            {
                Text = "OpenAI 設定",
                Location = new Point(10, y),
                Size = new Size(360, 23),
                Font = new Font(Font, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblOpenAI);
            y += 30;

            var lblOpenAIKey = new Label
            {
                Text = "APIキー:",
                Location = new Point(20, y),
                Size = new Size(80, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblOpenAIKey);
            
            txtOpenAIKey = new TextBox
            {
                Location = new Point(110, y),
                Size = new Size(260, 23),
                PasswordChar = '*',
                UseSystemPasswordChar = true
            };
            txtOpenAIKey.TextChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(txtOpenAIKey);
            y += 30;

            var lblOpenAIEndpoint = new Label
            {
                Text = "エンドポイント:",
                Location = new Point(20, y),
                Size = new Size(80, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblOpenAIEndpoint);
            
            txtOpenAIEndpoint = new TextBox
            {
                Location = new Point(110, y),
                Size = new Size(260, 23)
            };
            txtOpenAIEndpoint.TextChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(txtOpenAIEndpoint);
            y += 40;

            // Claude設定
            var lblClaude = new Label
            {
                Text = "Claude 設定",
                Location = new Point(10, y),
                Size = new Size(360, 23),
                Font = new Font(Font, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblClaude);
            y += 30;

            var lblClaudeKey = new Label
            {
                Text = "APIキー:",
                Location = new Point(20, y),
                Size = new Size(80, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblClaudeKey);
            
            txtClaudeKey = new TextBox
            {
                Location = new Point(110, y),
                Size = new Size(260, 23),
                PasswordChar = '*',
                UseSystemPasswordChar = true
            };
            txtClaudeKey.TextChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(txtClaudeKey);
            y += 40;

            // 接続テストボタン
            btnTestConnection = new Button
            {
                Text = "🔌 接続テスト",
                Location = new Point(10, y),
                Size = new Size(120, 30),
                UseVisualStyleBackColor = true
            };
            btnTestConnection.Click += BtnTestConnection_Click;
            panel.Controls.Add(btnTestConnection);
            y += 40;

            // サービス状態表示
            var lblStatus = new Label
            {
                Text = "サービス状態:",
                Location = new Point(10, y),
                Size = new Size(100, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblStatus);
            y += 25;

            lstServiceStatus = new ListBox
            {
                Location = new Point(10, y),
                Size = new Size(360, 100),
                Font = new Font("Courier New", 9)
            };
            panel.Controls.Add(lstServiceStatus);

            tabAIServices.Controls.Add(panel);
            tabControl.TabPages.Add(tabAIServices);
        }

        /// <summary>
        /// 一般設定タブの作成
        /// </summary>
        private void CreateGeneralTab()
        {
            tabGeneral = new TabPage("⚙️ 一般");
            
            var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            
            var y = 10;

            // 最大トークン数
            var lblMaxTokens = new Label
            {
                Text = "最大トークン数:",
                Location = new Point(10, y),
                Size = new Size(150, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblMaxTokens);
            
            numMaxTokens = new NumericUpDown
            {
                Location = new Point(170, y),
                Size = new Size(100, 23),
                Minimum = 100,
                Maximum = 8000,
                Value = 2000,
                Increment = 100
            };
            numMaxTokens.ValueChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(numMaxTokens);
            y += 35;

            // タイムアウト
            var lblTimeout = new Label
            {
                Text = "タイムアウト (秒):",
                Location = new Point(10, y),
                Size = new Size(150, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblTimeout);
            
            numTimeout = new NumericUpDown
            {
                Location = new Point(170, y),
                Size = new Size(100, 23),
                Minimum = 10,
                Maximum = 300,
                Value = 30,
                Increment = 5
            };
            numTimeout.ValueChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(numTimeout);
            y += 35;

            // プログレス表示
            chkShowProgress = new CheckBox
            {
                Text = "プログレス表示を有効にする",
                Location = new Point(10, y),
                Size = new Size(250, 23),
                Checked = true
            };
            chkShowProgress.CheckedChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(chkShowProgress);
            y += 30;

            // 自動更新
            chkAutoUpdate = new CheckBox
            {
                Text = "自動更新を有効にする",
                Location = new Point(10, y),
                Size = new Size(250, 23),
                Checked = true
            };
            chkAutoUpdate.CheckedChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(chkAutoUpdate);

            tabGeneral.Controls.Add(panel);
            tabControl.TabPages.Add(tabGeneral);
        }

        /// <summary>
        /// 高度な設定タブの作成
        /// </summary>
        private void CreateAdvancedTab()
        {
            tabAdvanced = new TabPage("🔧 高度な設定");
            
            var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            
            var y = 10;

            // ログ設定
            var lblLogging = new Label
            {
                Text = "ログ設定",
                Location = new Point(10, y),
                Size = new Size(360, 23),
                Font = new Font(Font, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblLogging);
            y += 30;

            chkEnableLogging = new CheckBox
            {
                Text = "ログ出力を有効にする",
                Location = new Point(20, y),
                Size = new Size(200, 23),
                Checked = true
            };
            chkEnableLogging.CheckedChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(chkEnableLogging);
            y += 30;

            var lblLogLevel = new Label
            {
                Text = "ログレベル:",
                Location = new Point(20, y),
                Size = new Size(100, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblLogLevel);
            
            cmbLogLevel = new ComboBox
            {
                Location = new Point(130, y),
                Size = new Size(120, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbLogLevel.Items.AddRange(new[] { "エラー", "警告", "情報", "デバッグ" });
            cmbLogLevel.SelectedIndex = 2; // 情報
            cmbLogLevel.SelectedIndexChanged += (s, e) => _hasUnsavedChanges = true;
            panel.Controls.Add(cmbLogLevel);
            y += 35;

            var lblLogPath = new Label
            {
                Text = "ログファイルパス:",
                Location = new Point(20, y),
                Size = new Size(100, 23),
                TextAlign = ContentAlignment.MiddleLeft
            };
            panel.Controls.Add(lblLogPath);
            
            txtLogPath = new TextBox
            {
                Location = new Point(130, y),
                Size = new Size(200, 23),
                ReadOnly = true,
                Text = "PTA_Outlook_Addin.log"
            };
            panel.Controls.Add(txtLogPath);
            y += 50;

            // リセットボタン
            btnResetSettings = new Button
            {
                Text = "🔄 設定をリセット",
                Location = new Point(10, y),
                Size = new Size(130, 30),
                UseVisualStyleBackColor = true
            };
            btnResetSettings.Click += BtnResetSettings_Click;
            panel.Controls.Add(btnResetSettings);

            tabAdvanced.Controls.Add(panel);
            tabControl.TabPages.Add(tabAdvanced);
        }

        /// <summary>
        /// ボタンの作成
        /// </summary>
        private void CreateButtons()
        {
            var buttonPanel = new Panel
            {
                Height = 50,
                Dock = DockStyle.Bottom,
                Padding = new Padding(10)
            };

            var btnOK = new Button
            {
                Text = "OK",
                Size = new Size(80, 30),
                Location = new Point(380, 10),
                DialogResult = DialogResult.OK,
                UseVisualStyleBackColor = true
            };
            btnOK.Click += BtnOK_Click;
            buttonPanel.Controls.Add(btnOK);

            var btnCancel = new Button
            {
                Text = "キャンセル",
                Size = new Size(80, 30),
                Location = new Point(470, 10),
                DialogResult = DialogResult.Cancel,
                UseVisualStyleBackColor = true
            };
            buttonPanel.Controls.Add(btnCancel);

            var btnApply = new Button
            {
                Text = "適用",
                Size = new Size(80, 30),
                Location = new Point(290, 10),
                UseVisualStyleBackColor = true
            };
            btnApply.Click += BtnApply_Click;
            buttonPanel.Controls.Add(btnApply);

            Controls.Add(buttonPanel);
        }

        /// <summary>
        /// レイアウトの調整
        /// </summary>
        private void AdjustLayout()
        {
            tabControl.Height = Height - 100; // ボタンパネル分を除く
        }

        #endregion

        #region イベントハンドラー

        /// <summary>
        /// 接続テストボタンクリック
        /// </summary>
        private async void BtnTestConnection_Click(object sender, EventArgs e)
        {
            try
            {
                btnTestConnection.Enabled = false;
                btnTestConnection.Text = "テスト中...";
                
                // 現在の設定で接続テスト
                await ApplySettings();
                
                var result = await _aiServiceManager.TestConnectionAsync();
                
                MessageBox.Show(result, "接続テスト結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                // サービス状態を更新
                UpdateServiceStatus();
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "接続テスト中にエラーが発生しました");
                MessageBox.Show($"接続テストでエラーが発生しました:\n{ex.Message}", 
                               "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnTestConnection.Enabled = true;
                btnTestConnection.Text = "🔌 接続テスト";
            }
        }

        /// <summary>
        /// 設定リセットボタンクリック
        /// </summary>
        private void BtnResetSettings_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("すべての設定をデフォルト値にリセットしますか？\nこの操作は元に戻せません。",
                                        "設定リセット確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                ResetToDefault();
                _hasUnsavedChanges = true;
                MessageBox.Show("設定がリセットされました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// OKボタンクリック
        /// </summary>
        private async void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                await ApplySettings();
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "設定保存中にエラーが発生しました");
                MessageBox.Show($"設定の保存でエラーが発生しました:\n{ex.Message}", 
                               "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 適用ボタンクリック
        /// </summary>
        private async void BtnApply_Click(object sender, EventArgs e)
        {
            try
            {
                await ApplySettings();
                _hasUnsavedChanges = false;
                MessageBox.Show("設定が適用されました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "設定適用中にエラーが発生しました");
                MessageBox.Show($"設定の適用でエラーが発生しました:\n{ex.Message}", 
                               "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// フォームクローズ時の処理
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_hasUnsavedChanges && DialogResult != DialogResult.OK)
            {
                var result = MessageBox.Show("未保存の変更があります。保存しますか？",
                                            "変更の保存", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        ApplySettings().Wait();
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "設定保存中にエラーが発生しました");
                        MessageBox.Show($"設定の保存でエラーが発生しました:\n{ex.Message}", 
                                       "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;
                        return;
                    }
                }
                else if (result == DialogResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
            }
            
            base.OnFormClosing(e);
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// 設定の読み込み
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                _logger?.LogDebug("設定の読み込み開始");

                // プライマリサービス選択肢の設定
                cmbPrimaryService.Items.Clear();
                foreach (var service in _aiServiceManager.RegisteredServices)
                {
                    cmbPrimaryService.Items.Add(service.ProviderName);
                }
                
                if (_aiServiceManager.PrimaryService != null)
                {
                    cmbPrimaryService.SelectedItem = _aiServiceManager.PrimaryService.ProviderName;
                }

                // 設定値の読み込み（ConfigurationManager経由で取得）
                // 実際の実装では、ConfigurationServiceから値を取得
                txtOpenAIKey.Text = "設定済み"; // マスク表示
                txtOpenAIEndpoint.Text = "https://api.openai.com/v1/chat/completions";
                txtClaudeKey.Text = "未設定";

                UpdateServiceStatus();
                
                _hasUnsavedChanges = false;
                _logger?.LogDebug("設定の読み込み完了");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "設定読み込み中にエラーが発生しました");
                MessageBox.Show($"設定の読み込みでエラーが発生しました:\n{ex.Message}", 
                               "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 設定の適用
        /// </summary>
        private async Task ApplySettings()
        {
            try
            {
                _logger?.LogDebug("設定の適用開始");

                // プライマリサービスの変更
                if (cmbPrimaryService.SelectedItem != null)
                {
                    var selectedProvider = cmbPrimaryService.SelectedItem.ToString();
                    _aiServiceManager.SetPrimaryService(selectedProvider);
                }

                // 実際の実装では、設定値をConfigurationServiceに保存
                // await _configurationService.SaveSettingsAsync(...);

                _hasUnsavedChanges = false;
                _logger?.LogDebug("設定の適用完了");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "設定適用中にエラーが発生しました");
                throw;
            }
        }

        /// <summary>
        /// デフォルト設定にリセット
        /// </summary>
        private void ResetToDefault()
        {
            txtOpenAIKey.Text = "";
            txtOpenAIEndpoint.Text = "https://api.openai.com/v1/chat/completions";
            txtClaudeKey.Text = "";
            numMaxTokens.Value = 2000;
            numTimeout.Value = 30;
            chkShowProgress.Checked = true;
            chkAutoUpdate.Checked = true;
            chkEnableLogging.Checked = true;
            cmbLogLevel.SelectedIndex = 2; // 情報
        }

        /// <summary>
        /// サービス状態の更新
        /// </summary>
        private void UpdateServiceStatus()
        {
            try
            {
                lstServiceStatus.Items.Clear();
                
                var status = _aiServiceManager.GetServiceStatus();
                var lines = status.Split('\n');
                
                foreach (var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        lstServiceStatus.Items.Add(line);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "サービス状態更新中にエラーが発生しました");
                lstServiceStatus.Items.Add("状態更新エラー");
            }
        }

        #endregion
    }
}
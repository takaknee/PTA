using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookPTAAddin.UI.Dialogs
{
    /// <summary>
    /// 改良されたプログレス表示ダイアログ
    /// 非同期処理の進行状況をユーザーフレンドリーに表示
    /// </summary>
    public partial class ProgressDialog : Form
    {
        #region フィールド

        private readonly CancellationTokenSource _cancellationTokenSource;
        private ProgressBar _progressBar;
        private Label _statusLabel;
        private Label _detailLabel;
        private Button _cancelButton;
        private PictureBox _iconPictureBox;
        private Timer _animationTimer;
        private int _animationFrame = 0;
        private bool _isIndeterminate = true;
        private bool _canCancel = true;

        // アニメーション用のアイコン
        private readonly string[] _animationIcons = { "🤖", "🔄", "⚡", "🔍", "📝", "✨" };

        #endregion

        #region プロパティ

        /// <summary>
        /// キャンセレーショントークン
        /// </summary>
        public CancellationToken CancellationToken => _cancellationTokenSource.Token;

        /// <summary>
        /// 処理がキャンセルされたかどうか
        /// </summary>
        public bool IsCancelled => _cancellationTokenSource.IsCancellationRequested;

        /// <summary>
        /// キャンセル可能かどうか
        /// </summary>
        public bool CanCancel
        {
            get => _canCancel;
            set
            {
                _canCancel = value;
                if (_cancelButton != null)
                {
                    _cancelButton.Enabled = value;
                    _cancelButton.Visible = value;
                }
            }
        }

        #endregion

        #region コンストラクター

        /// <summary>
        /// コンストラクター
        /// </summary>
        /// <param name="title">ダイアログタイトル</param>
        /// <param name="message">初期メッセージ</param>
        /// <param name="canCancel">キャンセル可能かどうか</param>
        public ProgressDialog(string title = "処理中", string message = "処理を実行しています...", bool canCancel = true)
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _canCancel = canCancel;
            
            InitializeComponent();
            
            Text = title;
            UpdateStatus(message);
            
            StartAnimation();
        }

        #endregion

        #region フォーム初期化

        /// <summary>
        /// コンポーネントの初期化
        /// </summary>
        private void InitializeComponent()
        {
            // フォームの基本設定
            Size = new Size(450, 200);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            TopMost = true;

            // レイアウトパネル
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                RowCount = 4,
                ColumnCount = 2
            };

            // 行と列のサイズ設定
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // アイコン行
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // ステータス行
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 詳細行
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // プログレスバー行
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // ボタン行
            
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // アイコン列
            mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // メイン列

            // アイコン
            _iconPictureBox = new PictureBox
            {
                Size = new Size(32, 32),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Margin = new Padding(0, 5, 10, 5)
            };
            mainPanel.Controls.Add(_iconPictureBox, 0, 0);

            // ステータスラベル
            _statusLabel = new Label
            {
                Text = "処理を実行しています...",
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 5, 0, 5)
            };
            mainPanel.Controls.Add(_statusLabel, 1, 0);

            // 詳細ラベル
            _detailLabel = new Label
            {
                Text = "",
                ForeColor = SystemColors.GrayText,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 0, 0, 10),
                MaximumSize = new Size(350, 0) // 自動改行用
            };
            mainPanel.SetColumnSpan(_detailLabel, 2);
            mainPanel.Controls.Add(_detailLabel, 0, 1);

            // プログレスバー
            _progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 50,
                Height = 20,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 0, 0, 15)
            };
            mainPanel.SetColumnSpan(_progressBar, 2);
            mainPanel.Controls.Add(_progressBar, 0, 2);

            // キャンセルボタン
            _cancelButton = new Button
            {
                Text = "キャンセル",
                Size = new Size(100, 30),
                Anchor = AnchorStyles.Right,
                UseVisualStyleBackColor = true,
                Enabled = _canCancel,
                Visible = _canCancel
            };
            _cancelButton.Click += CancelButton_Click;
            mainPanel.SetColumnSpan(_cancelButton, 2);
            mainPanel.Controls.Add(_cancelButton, 0, 3);

            Controls.Add(mainPanel);

            // タイマーの初期化
            _animationTimer = new Timer
            {
                Interval = 500,
                Enabled = false
            };
            _animationTimer.Tick += AnimationTimer_Tick;
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// ステータスメッセージを更新する
        /// </summary>
        /// <param name="message">メインメッセージ</param>
        /// <param name="detail">詳細メッセージ（オプション）</param>
        public void UpdateStatus(string message, string detail = null)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, string>(UpdateStatus), message, detail);
                return;
            }

            if (_statusLabel != null)
            {
                _statusLabel.Text = message ?? "処理中...";
            }

            if (_detailLabel != null)
            {
                _detailLabel.Text = detail ?? "";
                _detailLabel.Visible = !string.IsNullOrWhiteSpace(detail);
            }

            Application.DoEvents();
        }

        /// <summary>
        /// プログレスバーを確定的な値に設定する
        /// </summary>
        /// <param name="value">進行値（0-100）</param>
        /// <param name="maximum">最大値（デフォルト100）</param>
        public void SetProgress(int value, int maximum = 100)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<int, int>(SetProgress), value, maximum);
                return;
            }

            if (_progressBar != null)
            {
                _isIndeterminate = false;
                _progressBar.Style = ProgressBarStyle.Continuous;
                _progressBar.Maximum = maximum;
                _progressBar.Value = Math.Min(Math.Max(0, value), maximum);
            }
        }

        /// <summary>
        /// 不確定プログレスバーに戻す
        /// </summary>
        public void SetIndeterminate()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(SetIndeterminate));
                return;
            }

            if (_progressBar != null)
            {
                _isIndeterminate = true;
                _progressBar.Style = ProgressBarStyle.Marquee;
            }
        }

        /// <summary>
        /// 処理完了状態に更新する
        /// </summary>
        /// <param name="success">成功したかどうか</param>
        /// <param name="message">完了メッセージ</param>
        /// <param name="autoClose">自動で閉じるかどうか</param>
        /// <param name="autoCloseDelay">自動で閉じるまでの遅延（ミリ秒）</param>
        public void SetCompleted(bool success, string message = null, bool autoClose = true, int autoCloseDelay = 1500)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<bool, string, bool, int>(SetCompleted), success, message, autoClose, autoCloseDelay);
                return;
            }

            try
            {
                StopAnimation();

                var completionMessage = message ?? (success ? "処理が完了しました" : "処理が失敗しました");
                var icon = success ? "✅" : "❌";

                UpdateStatus(completionMessage);
                UpdateIcon(icon);

                if (_progressBar != null)
                {
                    _progressBar.Style = ProgressBarStyle.Continuous;
                    _progressBar.Value = success ? _progressBar.Maximum : 0;
                }

                if (_cancelButton != null)
                {
                    _cancelButton.Text = "閉じる";
                    _cancelButton.Enabled = true;
                    _cancelButton.Visible = true;
                }

                if (autoClose && success)
                {
                    var closeTimer = new Timer
                    {
                        Interval = autoCloseDelay,
                        Enabled = true
                    };
                    closeTimer.Tick += (s, e) =>
                    {
                        closeTimer.Dispose();
                        DialogResult = DialogResult.OK;
                        Close();
                    };
                }
            }
            catch (Exception)
            {
                // エラーが発生しても無視（フォームが既に閉じられている可能性）
            }
        }

        /// <summary>
        /// エラー状態に更新する
        /// </summary>
        /// <param name="errorMessage">エラーメッセージ</param>
        /// <param name="detail">詳細エラー情報</param>
        public void SetError(string errorMessage, string detail = null)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string, string>(SetError), errorMessage, detail);
                return;
            }

            StopAnimation();
            UpdateIcon("❌");
            UpdateStatus(errorMessage ?? "エラーが発生しました", detail);

            if (_progressBar != null)
            {
                _progressBar.Style = ProgressBarStyle.Continuous;
                _progressBar.Value = 0;
            }

            if (_cancelButton != null)
            {
                _cancelButton.Text = "閉じる";
                _cancelButton.Enabled = true;
                _cancelButton.Visible = true;
            }
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// アニメーションを開始する
        /// </summary>
        private void StartAnimation()
        {
            if (_animationTimer != null)
            {
                _animationTimer.Start();
            }
        }

        /// <summary>
        /// アニメーションを停止する
        /// </summary>
        private void StopAnimation()
        {
            if (_animationTimer != null)
            {
                _animationTimer.Stop();
            }
        }

        /// <summary>
        /// アイコンを更新する
        /// </summary>
        /// <param name="iconText">アイコンテキスト</param>
        private void UpdateIcon(string iconText)
        {
            if (_iconPictureBox == null) return;

            try
            {
                // シンプルなテキストベースのアイコン表示
                var bitmap = new Bitmap(32, 32);
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(BackColor);
                    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    
                    using (var font = new Font("Segoe UI Emoji", 16, FontStyle.Regular))
                    using (var brush = new SolidBrush(Color.Black))
                    {
                        var size = graphics.MeasureString(iconText, font);
                        var x = (32 - size.Width) / 2;
                        var y = (32 - size.Height) / 2;
                        graphics.DrawString(iconText, font, brush, x, y);
                    }
                }
                
                _iconPictureBox.Image?.Dispose();
                _iconPictureBox.Image = bitmap;
            }
            catch (Exception)
            {
                // アイコン更新に失敗しても無視
            }
        }

        #endregion

        #region イベントハンドラー

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void CancelButton_Click(object sender, EventArgs e)
        {
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _cancellationTokenSource.Cancel();
                UpdateStatus("キャンセル中...", "処理の停止を要求しています");
                _cancelButton.Enabled = false;
            }
            else
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        /// <summary>
        /// アニメーションタイマーティック
        /// </summary>
        private void AnimationTimer_Tick(object sender, EventArgs e)
        {
            if (_isIndeterminate && _animationIcons.Length > 0)
            {
                var iconText = _animationIcons[_animationFrame % _animationIcons.Length];
                UpdateIcon(iconText);
                _animationFrame++;
            }
        }

        /// <summary>
        /// フォームクローズ処理
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            StopAnimation();
            
            if (!_cancellationTokenSource.IsCancellationRequested && _canCancel)
            {
                _cancellationTokenSource.Cancel();
            }

            base.OnFormClosing(e);
        }

        /// <summary>
        /// リソース解放
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _animationTimer?.Dispose();
                _cancellationTokenSource?.Dispose();
                _iconPictureBox?.Image?.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion

        #region 静的ヘルパーメソッド

        /// <summary>
        /// 非同期タスクをプログレスダイアログ付きで実行する
        /// </summary>
        /// <typeparam name="T">戻り値の型</typeparam>
        /// <param name="parent">親ウィンドウ</param>
        /// <param name="task">実行するタスク</param>
        /// <param name="title">ダイアログタイトル</param>
        /// <param name="message">初期メッセージ</param>
        /// <param name="canCancel">キャンセル可能かどうか</param>
        /// <returns>タスクの実行結果</returns>
        public static async Task<T> RunWithProgressAsync<T>(
            IWin32Window parent,
            Func<ProgressDialog, CancellationToken, Task<T>> task,
            string title = "処理中",
            string message = "処理を実行しています...",
            bool canCancel = true)
        {
            using (var progressDialog = new ProgressDialog(title, message, canCancel))
            {
                var taskToRun = task(progressDialog, progressDialog.CancellationToken);
                
                // ダイアログを表示
                var dialogTask = Task.Run(() =>
                {
                    if (parent != null)
                    {
                        progressDialog.ShowDialog(parent);
                    }
                    else
                    {
                        progressDialog.ShowDialog();
                    }
                });

                try
                {
                    // タスクの完了を待機
                    var result = await taskToRun;
                    
                    // 成功時の表示
                    progressDialog.SetCompleted(true, "処理が完了しました");
                    
                    return result;
                }
                catch (OperationCanceledException)
                {
                    // キャンセルされた場合
                    progressDialog.SetCompleted(false, "処理がキャンセルされました", false);
                    throw;
                }
                catch (Exception ex)
                {
                    // エラーが発生した場合
                    progressDialog.SetError("処理中にエラーが発生しました", ex.Message);
                    throw;
                }
            }
        }

        /// <summary>
        /// 非同期タスク（戻り値なし）をプログレスダイアログ付きで実行する
        /// </summary>
        /// <param name="parent">親ウィンドウ</param>
        /// <param name="task">実行するタスク</param>
        /// <param name="title">ダイアログタイトル</param>
        /// <param name="message">初期メッセージ</param>
        /// <param name="canCancel">キャンセル可能かどうか</param>
        public static async Task RunWithProgressAsync(
            IWin32Window parent,
            Func<ProgressDialog, CancellationToken, Task> task,
            string title = "処理中",
            string message = "処理を実行しています...",
            bool canCancel = true)
        {
            await RunWithProgressAsync<object>(parent, async (dialog, token) =>
            {
                await task(dialog, token);
                return null;
            }, title, message, canCancel);
        }

        #endregion
    }
}
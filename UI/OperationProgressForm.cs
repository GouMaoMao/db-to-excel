using System;
using System.Drawing;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.UI
{
    /// <summary>显示长时间任务的阶段、进度和逐行日志，并向调用方发送取消请求。</summary>
    /// <remarks>
    /// 窗体保持置顶。成功或取消时自动关闭；发生错误时保留，供用户查看日志后手动关闭。
    /// 更新方法可从后台线程调用，内部会自动切换到 UI 线程。
    /// 日志列表按构造时传入的最低级别过滤；阶段标题、详情和进度条始终反映最新快照。
    /// </remarks>
    public sealed class OperationProgressForm : AppForm
    {
        private readonly Label _stageLabel;
        private readonly Label _detailLabel;
        private readonly ProgressBar _progressBar;
        private readonly RichTextBox _logBox;
        private readonly Button _cancelButton;
        private readonly Button _closeButton;
        private readonly LogSeverity _minimumSeverity;
        private bool _running = true;

        /// <summary>创建操作进度窗体。</summary>
        /// <param name="title">具体操作名称，会与产品名组合成窗口标题。</param>
        /// <param name="canCancel">是否允许用户请求取消。</param>
        /// <param name="minimumSeverity">日志列表的最低显示级别。阶段标题、详情和进度条不受此限制。</param>
        public OperationProgressForm(string title, bool canCancel, LogSeverity minimumSeverity)
        {
            _minimumSeverity = minimumSeverity;
            Text = AppPresentation.WindowTitle(title);
            Width = 680;
            Height = 400;
            MinimumSize = new Size(560, 320);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            TopMost = true;

            _stageLabel = new Label { AutoSize = true, Text = "正在准备…", Font = new Font(Font, FontStyle.Bold) };
            _detailLabel = new Label { AutoSize = true, Text = string.Empty };
            _progressBar = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee };
            _logBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font(AppPresentation.LogFontName, 8F),
                WordWrap = true,
                DetectUrls = false,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            _cancelButton = new Button { Text = "取消", AutoSize = true, Enabled = canCancel };
            _closeButton = new Button { Text = "关闭", AutoSize = true, Enabled = false, DialogResult = DialogResult.OK };

            _cancelButton.Click += (sender, args) =>
            {
                _cancelButton.Enabled = false;
                _stageLabel.Text = "正在取消…";
                CancelRequested?.Invoke(this, EventArgs.Empty);
            };
            _closeButton.Click += (sender, args) => Close();

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 5
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(_stageLabel, 0, 0);
            root.Controls.Add(_detailLabel, 0, 1);
            root.Controls.Add(_progressBar, 0, 2);
            root.Controls.Add(_logBox, 0, 3);

            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true
            };
            actions.Controls.Add(_closeButton);
            actions.Controls.Add(_cancelButton);
            root.Controls.Add(actions, 0, 4);
            Controls.Add(root);
        }

        /// <summary>用户点击取消按钮时触发；订阅方应取消对应令牌。</summary>
        public event EventHandler CancelRequested;

        /// <summary>使用最新进度快照更新阶段、日志和进度条。</summary>
        /// <param name="progress">后台操作报告的状态快照。</param>
        public void UpdateProgress(OperationProgress progress)
        {
            if (progress == null || IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<OperationProgress>(UpdateProgress), progress);
                return;
            }

            _stageLabel.Text = StageText(progress.Stage);
            _detailLabel.Text = progress.Message ?? string.Empty;
            AppendLog(progress);

            if (progress.IsIndeterminate || !progress.Percent.HasValue)
            {
                _progressBar.Style = ProgressBarStyle.Marquee;
            }
            else
            {
                _progressBar.Style = ProgressBarStyle.Continuous;
                _progressBar.Value = Math.Max(0, Math.Min(100, progress.Percent.Value));
            }

        }

        /// <summary>结束进度交互并根据结果关闭或保留窗体。</summary>
        /// <param name="error">任务异常；为空表示没有错误。</param>
        /// <param name="cancelled">任务是否被取消。</param>
        public void Complete(Exception error, bool cancelled)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<Exception, bool>(Complete), error, cancelled);
                return;
            }

            _running = false;
            _progressBar.Style = ProgressBarStyle.Continuous;
            _progressBar.Value = error == null && !cancelled ? 100 : 0;
            _stageLabel.Text = cancelled ? "已取消" : error == null ? "已完成" : "执行失败";
            if (error != null)
            {
                _detailLabel.Text = error.Message;
                AppendLog(LogSeverity.Error, "失败", null, error.Message);
            }
            _cancelButton.Enabled = false;

            if (error == null)
            {
                Close();
                return;
            }

            _closeButton.Enabled = true;
            AcceptButton = _closeButton;
        }

        /// <summary>运行期间把用户关闭窗口转换为取消请求，防止任务失去进度界面。</summary>
        /// <param name="e">窗体关闭事件参数。</param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_running)
            {
                e.Cancel = true;
                if (_cancelButton.Enabled) _cancelButton.PerformClick();
                return;
            }
            base.OnFormClosing(e);
        }

        private void AppendLog(OperationProgress progress)
        {
            string scope = !string.IsNullOrWhiteSpace(progress.TargetSheetName)
                ? progress.TargetSheetName
                : progress.ConnectionName;
            AppendLog(progress.Severity, StageText(progress.Stage), scope, progress.Message);
        }

        /// <summary>把一条进度写入日志列表。</summary>
        /// <param name="severity">该行级别。低于窗口最低级别时不追加，阶段标题和进度条仍由调用方更新。</param>
        /// <param name="stage">已转换为用户可见文字的阶段。</param>
        /// <param name="scope">工作表或连接名称，可为空。</param>
        /// <param name="message">状态消息，可为空。</param>
        /// <remarks><see cref="LogSeverity"/> 的声明顺序就是严重程度，因此用数值比较。</remarks>
        private void AppendLog(LogSeverity severity, string stage, string scope, string message)
        {
            if ((int)severity < (int)_minimumSeverity) return;
            string scopeText = string.IsNullOrWhiteSpace(scope) ? string.Empty : " [" + NormalizeLogText(scope) + "]";
            string messageText = string.IsNullOrWhiteSpace(message) ? string.Empty : "  " + NormalizeLogText(message);
            _logBox.AppendText(string.Format(
                "{0:HH:mm:ss}  {1}{2}{3}{4}",
                DateTime.Now,
                stage,
                scopeText,
                messageText,
                Environment.NewLine));
            _logBox.SelectionStart = _logBox.TextLength;
            _logBox.ScrollToCaret();
        }

        private static string NormalizeLogText(string value)
        {
            return value.Replace("\r", " ").Replace("\n", " ");
        }

        private static string StageText(OperationStage stage)
        {
            switch (stage)
            {
                case OperationStage.Preparing: return "准备";
                case OperationStage.Connecting: return "连接";
                case OperationStage.Executing: return "执行";
                case OperationStage.Reading: return "读取";
                case OperationStage.Writing: return "写入";
                case OperationStage.Cancelling: return "取消中";
                case OperationStage.Completed: return "完成";
                case OperationStage.Failed: return "失败";
                case OperationStage.Cancelled: return "已取消";
                default: return stage.ToString();
            }
        }
    }
}

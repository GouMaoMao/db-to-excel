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
    /// 日志列表按构造时传入的最低级别过滤。先按快照改条和标题，再写日志；写完日志就不再改界面，避免日志重入把旧快照画回去。
    /// 批量刷新的读取条只按已结束个数占总数绘制，分母取第一次出现双条时的总数。个数达到总数后读取条钉住，后到的读取快照只进日志。
    /// 写入条的序号和填充来自同一次快照，序号只向前。读取个数在阻塞写入之前已经画出。单次查询仍只有一条。两条都只向前，填充色都用系统高亮色。
    /// 长报错在详情区和日志里按窗体宽度换行，不能把列撑出客户区，以免挡住取消和关闭。
    /// </remarks>
    public sealed class OperationProgressForm : AppForm
    {
        private readonly Label _stageLabel;
        private readonly Label _detailLabel;
        private readonly Panel _meterHost;
        private readonly ProgressBar _progressBar;
        private readonly Panel _readRow;
        private readonly Panel _writeRow;
        private readonly Label _readCaption;
        private readonly Label _writeCaption;
        private readonly CountBar _readBar;
        private readonly CountBar _writeBar;
        private readonly RowStyle _meterRow;
        private readonly RichTextBox _logBox;
        private readonly Button _cancelButton;
        private readonly Button _closeButton;
        private readonly LogSeverity _minimumSeverity;
        private bool _running = true;
        private bool _dualMeters;
        private bool _readsPinned;
        private int _readTotal;
        private int _writePercentShown = -1;
        private int _readCountShown;
        private int _writeCountShown;

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
            _progressBar = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee, Height = 18 };
            _readRow = CreateMeterRow("读取", out _readCaption, out _readBar);
            _writeRow = CreateMeterRow("写入", out _writeCaption, out _writeBar);
            _readRow.Visible = false;
            _writeRow.Visible = false;
            _meterHost = new Panel { Dock = DockStyle.Fill, Height = 22 };
            _meterHost.Controls.Add(_progressBar);
            _meterHost.Controls.Add(_writeRow);
            _meterHost.Controls.Add(_readRow);
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
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _meterRow = new RowStyle(SizeType.Absolute, 22);
            root.RowStyles.Add(_meterRow);
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(_stageLabel, 0, 0);
            root.Controls.Add(_detailLabel, 0, 1);
            root.Controls.Add(_meterHost, 0, 2);
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
            Resize += (sender, args) => ApplyWrappedLabelWidth();
            ApplyWrappedLabelWidth();
        }

        /// <summary>把阶段标题、详情和日志限制在客户区宽度内，详情最多约三行。</summary>
        /// <remarks>宽度跟窗体走。详情超出三行后不再长高；日志在同一宽度内换行，完整内容仍可滚动查看。</remarks>
        private void ApplyWrappedLabelWidth()
        {
            int width = ClientSize.Width - 24;
            if (width < 80) return;
            _stageLabel.MaximumSize = new Size(width, 0);
            _detailLabel.MaximumSize = new Size(width, Math.Max(Font.Height * 3, 48));
            _logBox.MaximumSize = new Size(width, 0);
        }

        /// <summary>用户点击取消按钮时触发；订阅方应取消对应令牌。</summary>
        public event EventHandler CancelRequested;

        /// <summary>使用最新进度快照更新阶段、日志和进度条。</summary>
        /// <param name="progress">后台操作报告的状态快照。</param>
        /// <remarks>先改条和标题，再追加日志。日志控件会重入界面队列，因此本方法在写日志之后不再改界面。</remarks>
        public void UpdateProgress(OperationProgress progress)
        {
            if (progress == null || IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<OperationProgress>(UpdateProgress), progress);
                return;
            }

            ApplySnapshot(progress);
            AppendLog(progress);
        }

        /// <summary>按快照推进读取条、写入条、标题和详情。不写日志。</summary>
        /// <param name="progress">后台操作报告的状态快照。</param>
        /// <remarks>
        /// 读取个数达到总数后，阶段仍是读取的快照到此为止，只留给调用方写日志。
        /// 写入条只按当前快照的序号和填充向前画。
        /// </remarks>
        private void ApplySnapshot(OperationProgress progress)
        {
            if (progress.ReadTaskCount.HasValue || progress.WritePercent.HasValue)
            {
                ShowDualMeters(progress.TotalTasks);
                if (progress.ReadTaskCount.HasValue)
                    ApplyReadCount(progress.ReadTaskCount.Value, progress.TotalTasks);
                if (IsStaleReadSnapshot(progress)) return;
                if (progress.WritePercent.HasValue && !TryApplyWriteMeter(progress)) return;
                ApplyStage(progress);
                return;
            }

            if (_dualMeters)
            {
                if (_readsPinned && progress.Stage == OperationStage.Reading) return;
                if (!_readsPinned) _detailLabel.Text = progress.Message ?? string.Empty;
                return;
            }

            ApplyStage(progress);
            if (progress.TotalTasks > 1)
            {
                _progressBar.Style = ProgressBarStyle.Continuous;
                _progressBar.Value = OverallPercent(progress);
                return;
            }

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

        /// <summary>读取已经到总数，且这条快照仍是读取阶段、又不带写入百分比。</summary>
        /// <param name="progress">当前快照。</param>
        /// <returns>为 true 时不能改标题、详情和读取条。</returns>
        private bool IsStaleReadSnapshot(OperationProgress progress)
        {
            return _readsPinned
                && progress.Stage == OperationStage.Reading
                && !progress.WritePercent.HasValue;
        }

        /// <summary>把快照写进阶段标题和详情。</summary>
        /// <param name="progress">当前快照。</param>
        private void ApplyStage(OperationProgress progress)
        {
            _detailLabel.Text = progress.Message ?? string.Empty;
            _stageLabel.Text = progress.TotalTasks > 0 ? OverallStageText(progress) : StageText(progress.Stage);
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
            if (_dualMeters)
            {
                if (error == null && !cancelled)
                    _writeBar.SetPercent(100);
            }
            else
            {
                _progressBar.Style = ProgressBarStyle.Continuous;
                _progressBar.Value = error == null && !cancelled ? 100 : 0;
            }
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

        /// <summary>批量刷新改成读取、写入两条。单条进度条保持不动，直到出现读取个数或写入百分比。</summary>
        /// <remarks>总数只取第一次的值。后到快照不能把分母改掉。</remarks>
        private void ShowDualMeters(int totalTasks)
        {
            if (_readTotal < 1 && totalTasks > 0) _readTotal = totalTasks;
            if (_dualMeters) return;
            _dualMeters = true;
            _progressBar.Visible = false;
            _readRow.Visible = true;
            _writeRow.Visible = true;
            _meterHost.Height = 48;
            _meterRow.Height = 48;
            int total = _readTotal > 0 ? _readTotal : totalTasks;
            if (total > 0)
            {
                _readCaption.Text = "读取 0/" + total.ToString();
                _writeCaption.Text = "写入 0/" + total.ToString();
            }
        }

        /// <summary>用已结束个数同时更新读取文字和读取条。个数没有变大就不重画。</summary>
        /// <remarks>分母固定。个数达到总数时写成总数/总数并钉住。个数变化后立即重画，不等界面线程从阻塞写入里返回。</remarks>
        private void ApplyReadCount(int count, int totalTasks)
        {
            if (_readTotal < 1 && totalTasks > 0) _readTotal = totalTasks;
            int total = _readTotal > 0 ? _readTotal : totalTasks;
            if (total < 1 || count <= _readCountShown) return;
            if (count > total) count = total;
            _readCountShown = count;
            if (count >= total)
            {
                _readsPinned = true;
                _readCountShown = total;
                count = total;
            }

            _readCaption.Text = string.Format("读取 {0}/{1}", count, total);
            _readBar.SetCount(count, total);
            if (_readCaption.IsHandleCreated) _readCaption.Update();
            if (_readBar.IsHandleCreated) _readBar.Update();
        }

        /// <summary>按当前快照更新写入条和左侧序号。序号小于 1 时不画。</summary>
        /// <param name="progress">带写入百分比的快照。</param>
        /// <returns>已画出写入条时为 true。序号还没到 1 时为 false，调用方不要改标题。</returns>
        /// <remarks>序号和填充来自同一次快照，只向前。</remarks>
        private bool TryApplyWriteMeter(OperationProgress progress)
        {
            int? taskCount = progress.WriteTaskCount;
            if (!taskCount.HasValue || taskCount.Value < 1) return false;
            return PaintWriteMeter(taskCount.Value, progress.WritePercent ?? 0);
        }

        /// <summary>按同一次快照的序号和百分比画写入条。序号或百分比退回时不改。</summary>
        /// <param name="taskCount">从 1 开始的写入序号。</param>
        /// <param name="percent">0 到 100 的写入百分比。</param>
        /// <returns>这次快照推进了序号或填充时为 true。</returns>
        private bool PaintWriteMeter(int taskCount, int percent)
        {
            int total = _readTotal;
            if (taskCount < 1 || total < 1 || taskCount < _writeCountShown) return false;
            if (taskCount == _writeCountShown && percent < _writePercentShown) return false;
            if (percent >= _writePercentShown)
            {
                _writePercentShown = percent;
                _writeBar.SetPercent(percent);
            }

            _writeCountShown = taskCount;
            _writeCaption.Text = string.Format("写入 {0}/{1}", taskCount, total);
            return true;
        }

        /// <summary>左侧文字加一条系统高亮色填充条。先加入条再停靠文字，避免文字被盖住。</summary>
        private static Panel CreateMeterRow(string caption, out Label label, out CountBar bar)
        {
            label = new Label
            {
                Text = caption,
                Dock = DockStyle.Left,
                Width = 96,
                TextAlign = ContentAlignment.MiddleLeft
            };
            bar = new CountBar { Dock = DockStyle.Fill };
            Panel row = new Panel { Dock = DockStyle.Top, Height = 22 };
            row.Controls.Add(bar);
            row.Controls.Add(label);
            return row;
        }

        /// <summary>按分子分母宽度填充，填充色为 <see cref="SystemColors.Highlight"/>。</summary>
        /// <remarks>不用系统进度条，避免接近满格时被画成已经完成，也避免主题色与高亮色不一致。</remarks>
        private sealed class CountBar : Control
        {
            private int _count;
            private int _total;

            /// <summary>创建一条空白的填充条。</summary>
            public CountBar()
            {
                SetStyle(
                    ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint,
                    true);
                Height = 18;
            }

            /// <summary>记下个数和总数并重画。个数不会超过总数。</summary>
            /// <param name="count">已经结束的个数。</param>
            /// <param name="total">总数。小于 1 时不画填充。</param>
            public void SetCount(int count, int total)
            {
                if (count < 0) count = 0;
                if (total < 1) total = 0;
                if (count > total) count = total;
                _count = count;
                _total = total;
                Invalidate();
            }

            /// <summary>按 0 到 100 的百分比填充并重画。</summary>
            /// <param name="percent">完成百分比。</param>
            public void SetPercent(int percent)
            {
                if (percent < 0) percent = 0;
                if (percent > 100) percent = 100;
                SetCount(percent, 100);
            }

            /// <summary>背景留空，填充宽度等于个数乘以内宽再除以总数，颜色用系统高亮色。</summary>
            /// <param name="e">绘制参数。</param>
            protected override void OnPaint(PaintEventArgs e)
            {
                Rectangle bounds = ClientRectangle;
                if (bounds.Width <= 0 || bounds.Height <= 0) return;
                using (SolidBrush back = new SolidBrush(SystemColors.Control))
                    e.Graphics.FillRectangle(back, bounds);
                ControlPaint.DrawBorder(e.Graphics, bounds, SystemColors.ControlDark, ButtonBorderStyle.Solid);
                if (_total < 1 || _count <= 0) return;
                int innerWidth = bounds.Width - 2;
                if (innerWidth <= 0) return;
                int fill = (int)(_count * (long)innerWidth / _total);
                if (fill <= 0) return;
                Rectangle bar = new Rectangle(bounds.X + 1, bounds.Y + 1, fill, bounds.Height - 2);
                using (SolidBrush fore = new SolidBrush(SystemColors.Highlight))
                    e.Graphics.FillRectangle(fore, bar);
            }
        }

        /// <summary>批次进度写在进度条上方。日志仍只用阶段名，避免每一行都重复任务序号。</summary>
        private static string OverallStageText(OperationProgress progress)
        {
            string sheet = string.IsNullOrWhiteSpace(progress.TargetSheetName)
                ? string.Empty
                : progress.TargetSheetName.Trim() + "  ";
            return string.Format(
                "{0}/{1}  {2}{3}",
                Math.Max(0, progress.CurrentTask),
                progress.TotalTasks,
                sheet,
                StageText(progress.Stage));
        }

        /// <summary>用已完成任务数加上当前任务自己的百分比，得到整批进度。</summary>
        private static int OverallPercent(OperationProgress progress)
        {
            if (progress.Stage == OperationStage.Completed) return 100;
            int completed = progress.CurrentTask < 1 ? 0 : progress.CurrentTask - 1;
            if (completed > progress.TotalTasks) completed = progress.TotalTasks;
            int inner = progress.Percent.HasValue ? progress.Percent.Value : 0;
            if (inner < 0) inner = 0;
            if (inner > 100) inner = 100;
            long overall = (completed * 100L + inner) / progress.TotalTasks;
            if (overall < 0) overall = 0;
            if (overall > 100) overall = 100;
            return (int)overall;
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

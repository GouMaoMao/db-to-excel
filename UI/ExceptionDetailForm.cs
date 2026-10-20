using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>向用户展示完整错误详情，含异常类型、消息、堆栈和内部异常。</summary>
    /// <remarks>
    /// 用户操作失败时必须弹出本窗，不得只写状态栏或静默吞掉。
    /// 详情文本可复制，便于排查。
    /// </remarks>
    public sealed class ExceptionDetailForm : AppForm
    {
        private readonly TextBox _summary;
        private readonly TextBox _detail;
        private bool _askAi;

        /// <summary>创建错误详情窗体。</summary>
        /// <param name="summary">简短说明，通常是异常消息。</param>
        /// <param name="detail">完整详情文本。</param>
        /// <param name="offerAskAi">为 true 时显示「问 AI」。只有查询执行失败会打开。</param>
        public ExceptionDetailForm(string summary, string detail, bool offerAskAi)
        {
            Text = AppPresentation.WindowTitle("错误详情");
            Width = 720;
            Height = 480;
            MinimumSize = new Size(520, 360);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;

            _summary = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Window,
                Text = summary ?? string.Empty
            };
            _detail = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font(AppPresentation.LogFontName, AppPresentation.LogFontSize, FontStyle.Regular, GraphicsUnit.Point),
                Text = detail ?? string.Empty
            };

            Button copyButton = new Button { Text = "复制详情", AutoSize = true };
            Button closeButton = new Button { Text = "关闭", AutoSize = true, DialogResult = DialogResult.OK };
            copyButton.Click += (sender, args) =>
            {
                try
                {
                    Clipboard.SetText(string.IsNullOrEmpty(_detail.Text) ? (_summary.Text ?? string.Empty) : _detail.Text);
                }
                catch (Exception)
                {
                    // 剪贴板偶发占用；复制失败不阻断关闭。
                }
            };

            FlowLayoutPanel footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true
            };
            footer.Controls.Add(closeButton);
            footer.Controls.Add(copyButton);
            if (offerAskAi)
            {
                Button askButton = new Button { Text = "问 AI", AutoSize = true, UseVisualStyleBackColor = true };
                askButton.Click += (sender, args) =>
                {
                    _askAi = true;
                    DialogResult = DialogResult.OK;
                    Close();
                };
                footer.Controls.Add(askButton);
            }

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(new Label
            {
                Text = "发生错误。下面是完整信息，便于排查。",
                AutoSize = true,
                Padding = new Padding(0, 0, 0, 6)
            }, 0, 0);
            root.Controls.Add(_summary, 0, 1);
            root.Controls.Add(_detail, 0, 2);
            root.Controls.Add(footer, 0, 3);
            Controls.Add(root);

            AcceptButton = closeButton;
            CancelButton = closeButton;
        }

        /// <summary>弹出异常详情窗。</summary>
        /// <param name="owner">父窗体；可为 null。</param>
        /// <param name="exception">异常；为空时显示未知错误。</param>
        public static void Show(IWin32Window owner, Exception exception)
        {
            Show(owner, exception, false);
        }

        /// <summary>弹出异常详情窗，并在查询执行失败时提供「问 AI」。</summary>
        /// <param name="owner">父窗体；可为 null。</param>
        /// <param name="exception">异常；为空时显示未知错误。</param>
        /// <param name="offerAskAi">为 true 时显示「问 AI」。</param>
        /// <returns>用户点了「问 AI」时为 true。点关闭或没有该按钮时为 false。</returns>
        public static bool Show(IWin32Window owner, Exception exception, bool offerAskAi)
        {
            if (exception == null)
                return Show(owner, "未知错误", "没有可用的异常对象。", offerAskAi);

            return Show(owner, exception.Message, exception.ToString(), offerAskAi);
        }

        /// <summary>弹出错误详情窗（无异常对象时使用，例如连接测试失败消息）。</summary>
        /// <param name="owner">父窗体；可为 null。</param>
        /// <param name="summary">简短说明。</param>
        /// <param name="detail">完整详情。</param>
        public static void Show(IWin32Window owner, string summary, string detail)
        {
            Show(owner, summary, detail, false);
        }

        /// <summary>弹出错误详情窗。</summary>
        /// <param name="owner">父窗体；可为 null。</param>
        /// <param name="summary">简短说明。</param>
        /// <param name="detail">完整详情。</param>
        /// <param name="offerAskAi">为 true 时显示「问 AI」。</param>
        /// <returns>用户点了「问 AI」时为 true。</returns>
        public static bool Show(IWin32Window owner, string summary, string detail, bool offerAskAi)
        {
            using (ExceptionDetailForm form = new ExceptionDetailForm(
                string.IsNullOrWhiteSpace(summary) ? "操作失败。" : summary.Trim(),
                string.IsNullOrWhiteSpace(detail) ? (summary ?? string.Empty) : detail,
                offerAskAi))
            {
                if (owner == null)
                    form.ShowDialog();
                else
                    form.ShowDialog(owner);
                return form._askAi;
            }
        }
    }
}

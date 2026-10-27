using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>在功能窗体首开准备期间显示无边框提示，告知用户程序仍在打开中。</summary>
    /// <remarks>
    /// 该窗体只用于短时等待提示，不承载业务输入，也不抢占焦点。
    /// 调用方应在 UI 线程创建、显示和关闭该窗体；主窗体完成首次展示后立刻关闭。
    /// </remarks>
    internal sealed class OpeningHintForm : AppForm
    {
        private readonly Label _messageLabel;

        /// <summary>创建一个固定文案为“正在打开…”的轻量提示窗。</summary>
        public OpeningHintForm()
        {
            InitializeWindow(string.Empty, 220, 74, Size.Empty, FormStartPosition.Manual);
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(250, 250, 250);
            TopMost = true;

            _messageLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = AppPresentation.OpeningHintText,
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = false,
                Font = new Font(AppPresentation.DefaultFontName, 9F, FontStyle.Regular, GraphicsUnit.Point),
                ForeColor = Color.FromArgb(64, 64, 64)
            };

            Panel surface = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = BackColor,
                Padding = new Padding(8)
            };
            surface.Controls.Add(_messageLabel);
            Controls.Add(LightBorder.Wrap(surface));
        }

        /// <summary>按目标区域居中摆放提示窗。区域为空时按当前屏幕工作区居中。</summary>
        /// <param name="bounds">目标区域。传入空区域时回退到鼠标所在屏幕。</param>
        public void PlaceCenter(Rectangle bounds)
        {
            Rectangle area = bounds.Width > 0 && bounds.Height > 0
                ? bounds
                : FormSizeMemory.CursorWorkingArea();
            int left = area.Left + Math.Max(0, (area.Width - Width) / 2);
            int top = area.Top + Math.Max(0, (area.Height - Height) / 2);
            Location = new Point(left, top);
        }

        /// <summary>返回 true 使窗口显示时不抢占焦点。</summary>
        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        /// <summary>添加不激活扩展样式，防止打断 Excel 当前输入焦点。</summary>
        /// <param name="cp">窗口创建参数。</param>
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;
                return cp;
            }
        }
    }
}

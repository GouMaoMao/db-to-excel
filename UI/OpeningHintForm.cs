using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>在功能窗体首开准备期间显示无边框提示，告知用户程序仍在打开中。</summary>
    /// <remarks>
    /// 该窗体只用于短时等待提示，不承载业务输入，也不抢占焦点。
    /// 调用方应在 UI 线程创建、显示和关闭该窗体；主窗体完成首次展示后立刻关闭。
    /// SQL 查询和批量刷新共用这一提示。图标在构造时从嵌入资源解码一次，不在 Excel 启动时预加载。
    /// 图标交给图片框持有，窗体关闭时由图片框释放。提示字体在类型首次使用时创建，进程内复用。
    /// </remarks>
    internal sealed class OpeningHintForm : AppForm
    {
        private const int IconEdge = 32;

        private static readonly Font HintFont = new Font(
            AppPresentation.DefaultFontName,
            AppPresentation.OpeningHintFontSize,
            FontStyle.Regular,
            GraphicsUnit.Point);

        private readonly Control _content;

        /// <summary>创建一个固定文案为“正在打开…”的轻量提示窗。</summary>
        /// <remarks>图标放在文案左侧，两者作为一组在窗体中居中。不播放动画。</remarks>
        public OpeningHintForm()
        {
            InitializeWindow(string.Empty, 440, 148, Size.Empty, FormStartPosition.Manual);
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            ControlBox = false;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(250, 250, 250);
            TopMost = true;

            Image icon = DatabaseTreeImageCatalog.LoadActionIcon("icon_robot.png", IconEdge);
            Label message = new Label
            {
                AutoSize = true,
                Text = AppPresentation.OpeningHintText,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = HintFont,
                ForeColor = Color.FromArgb(64, 64, 64),
                BackColor = BackColor,
                Margin = Padding.Empty
            };
            Size textSize = message.PreferredSize;
            int rowHeight = Math.Max(IconEdge, textSize.Height);
            message.AutoSize = false;
            message.Size = new Size(Math.Max(1, textSize.Width), rowHeight);

            PictureBox iconBox = new PictureBox
            {
                Image = icon,
                SizeMode = PictureBoxSizeMode.CenterImage,
                Size = new Size(IconEdge, rowHeight),
                Margin = new Padding(0, 0, 12, 0),
                BackColor = BackColor
            };

            FlowLayoutPanel row = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = BackColor,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            row.Controls.Add(iconBox);
            row.Controls.Add(message);
            _content = row;

            Panel surface = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = BackColor
            };
            surface.Controls.Add(row);
            surface.Resize += (sender, args) => CenterContent(surface);
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

        /// <summary>句柄创建后把图标和文案摆到客户区中央。</summary>
        /// <param name="e">加载事件参数。</param>
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            _content.PerformLayout();
            CenterContent(_content.Parent);
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

        /// <summary>把图标和文案作为一组摆到提示表面的中央。</summary>
        /// <param name="surface">铺满客户区的表面。内容直接放在它上面。</param>
        private void CenterContent(Control surface)
        {
            if (surface == null || _content == null) return;
            _content.Left = Math.Max(0, (surface.ClientSize.Width - _content.Width) / 2);
            _content.Top = Math.Max(0, (surface.ClientSize.Height - _content.Height) / 2);
        }
    }
}

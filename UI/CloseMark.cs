using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>在文字末尾留出空位，并画出比正文大一号的「×」。</summary>
    /// <remarks>
    /// 页签标题和输入区引用都用这一处测量、绘制和命中判断，避免各写一个关闭按钮。
    /// 页签外观仍由系统绘制；标记在系统画完后补上。标记字体随进程保留，不随某个窗体释放。
    /// </remarks>
    internal static class CloseMark
    {
        /// <summary>画在空位上的关闭标记。</summary>
        public const string Mark = "×";

        private const float MarkPointSize = 13F;
        private const int MinWidth = 22;
        private const int RightMargin = 4;
        private const int WmPaint = 0x000F;
        private static readonly Font MarkFont = new Font(
            AppPresentation.DefaultFontName,
            MarkPointSize,
            FontStyle.Bold,
            GraphicsUnit.Point);

        /// <summary>创建会在系统画完页签后补上关闭标记的页签控件。</summary>
        /// <returns>尚未设置停靠和尺寸的页签控件。</returns>
        public static TabControl CreateTabs()
        {
            return new MarkedTabs();
        }

        /// <summary>创建会在文字右侧画出关闭标记的平按钮。</summary>
        /// <returns>尚未设置文字的按钮。文字用 <see cref="Caption"/>。</returns>
        public static Button CreateChip()
        {
            return new MarkedButton();
        }

        /// <summary>把表意文字空格接到文字末尾，供系统按这段文字撑开宽度。</summary>
        /// <param name="text">标记前面的文字。为空时只返回空位。</param>
        /// <param name="hostFont">宿主绘制正文的字体。用来把空位撑到能放下标记，并留出一个字的空隙。</param>
        /// <returns>末尾留有空位的文字。标记本身由 <see cref="Draw"/> 画在空位右侧，不盖住标题。</returns>
        /// <remarks>普通空格和不换行空格会被页签丢掉，标题后面就没有位置。表意文字空格会占宽度。</remarks>
        public static string Caption(string text, Font hostFont)
        {
            return (text ?? string.Empty) + Spacer(hostFont);
        }

        /// <summary>把标记画在宿主右侧，中线与正文字行对齐。</summary>
        /// <param name="graphics">目标画面。为空时不画。</param>
        /// <param name="hostFont">宿主绘制正文的字体。用来确定文字行的位置。</param>
        /// <param name="host">文字所在的客户区矩形。高度按文字行计算，不要把选中页签多出来的底边算进去。</param>
        /// <param name="color">标记颜色，跟宿主文字一致。</param>
        /// <remarks>
        /// 不用 <see cref="TextFormatFlags.VerticalCenter"/>。它和 <see cref="TextFormatFlags.NoPadding"/> 一起用时会把字贴在矩形底边。
        /// </remarks>
        public static void Draw(Graphics graphics, Font hostFont, Rectangle host, Color color)
        {
            if (graphics == null || host.Width <= 0 || host.Height <= 0) return;
            Rectangle slot = GlyphBounds(host);
            int line = LineHeight(hostFont ?? MarkFont);
            int mark = LineHeight(MarkFont);
            int lineTop = host.Top + Math.Max(0, (host.Height - line) / 2);
            int y = lineTop + (line - mark) / 2;
            int width = Measure(Mark, MarkFont);
            int x = slot.Left + Math.Max(0, (slot.Width - width) / 2);
            TextRenderer.DrawText(
                graphics,
                Mark,
                MarkFont,
                new Rectangle(x, y, Math.Max(1, width), Math.Max(1, mark)),
                color,
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
        }

        /// <summary>计算标记在宿主矩形右侧所占的区域。宽度至少 22 像素，高度与宿主相同。</summary>
        /// <param name="host">文字所在的客户区矩形。</param>
        /// <returns>标记的点击区域。</returns>
        public static Rectangle GlyphBounds(Rectangle host)
        {
            int width = MarkWidth();
            int left = host.Right - RightMargin - width;
            if (left < host.Left) left = host.Left;
            int fit = Math.Max(1, host.Right - left);
            return new Rectangle(left, host.Top, Math.Min(width, fit), host.Height);
        }

        /// <summary>判断点击是否落在标记上。</summary>
        /// <param name="host">文字所在的客户区矩形。</param>
        /// <param name="location">相对宿主的点击位置。</param>
        /// <returns>落在标记上时为 true。</returns>
        public static bool Hit(Rectangle host, Point location)
        {
            return GlyphBounds(host).Contains(location);
        }

        private static int MarkWidth()
        {
            return Math.Max(MinWidth, Measure(Mark, MarkFont));
        }

        /// <summary>用表意文字空格撑出标记和大约一个字的空隙。页签计算宽度时会保留它。</summary>
        private static string Spacer(Font hostFont)
        {
            Font font = hostFont ?? MarkFont;
            int unit = Measure("\u3000", font);
            if (unit < 1) unit = 4;
            int needed = unit + MarkWidth() + RightMargin;
            int count = Math.Max(2, (needed + unit - 1) / unit);
            return new string('\u3000', count);
        }

        private static int Measure(string text, Font font)
        {
            return MeasureSize(text, font).Width;
        }

        /// <summary>量出一行文字的高度，供标记和标题共用同一条中线。</summary>
        private static int LineHeight(Font font)
        {
            return MeasureSize("国", font).Height;
        }

        private static Size MeasureSize(string text, Font font)
        {
            return TextRenderer.MeasureText(
                text ?? string.Empty,
                font,
                new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
        }

        /// <summary>页签里真正放标题的矩形。选中页更高时按较矮的那一档，避免底边把标记带下去。</summary>
        /// <param name="tabs">正在画的页签控件。</param>
        /// <param name="index">页签序号。</param>
        /// <returns>顶边与原页签相同、高度不超过其他页签的矩形。</returns>
        private static Rectangle TabTextRect(TabControl tabs, int index)
        {
            Rectangle rect = tabs.GetTabRect(index);
            int shortest = rect.Height;
            for (int other = 0; other < tabs.TabCount; other++)
            {
                if (other == index) continue;
                int height = tabs.GetTabRect(other).Height;
                if (height > 0 && height < shortest)
                    shortest = height;
            }

            rect.Height = shortest;
            return rect;
        }

        /// <summary>在系统画完的页签上补标记。绘制失败时留下空位，不把绘制过程当成用户操作失败。</summary>
        private static void PaintTabs(TabControl tabs)
        {
            if (tabs == null || tabs.IsDisposed || !tabs.IsHandleCreated) return;
            try
            {
                using (Graphics graphics = Graphics.FromHwnd(tabs.Handle))
                {
                    for (int index = 0; index < tabs.TabCount; index++)
                        Draw(graphics, tabs.Font, TabTextRect(tabs, index), tabs.ForeColor);
                }
            }
            catch (Exception)
            {
                // 句柄正在销毁时取画面可能失败。这不是用户操作，不能在绘制里弹错误窗。
            }
        }

        /// <summary>系统页签。画完后只补关闭标记，不改页签底和边。</summary>
        private sealed class MarkedTabs : TabControl
        {
            private bool _painting;

            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);
                if (m.Msg != WmPaint || _painting) return;
                _painting = true;
                try
                {
                    PaintTabs(this);
                }
                finally
                {
                    _painting = false;
                }
            }
        }

        /// <summary>引用芯片。文字由系统画出，标记画在右侧空位上。</summary>
        private sealed class MarkedButton : Button
        {
            /// <summary>系统画完文字后，把标记画在右侧空位上。</summary>
            /// <param name="e">按钮的绘制参数。</param>
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                Draw(e.Graphics, Font, ClientRectangle, ForeColor);
            }
        }
    }
}

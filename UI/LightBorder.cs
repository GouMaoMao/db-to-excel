using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>用 1 像素较浅的系统色标出区域。比 <see cref="BorderStyle.FixedSingle"/> 浅，并跟着系统主题。</summary>
    internal static class LightBorder
    {
        /// <summary>在内容四周留出 1 像素，露出底色作为边框。内容铺满内区，边线不会被盖住。</summary>
        /// <param name="content">要框住的控件。停靠改为铺满。</param>
        /// <returns>调用方负责加入父容器的外框。</returns>
        public static Panel Wrap(Control content)
        {
            Panel frame = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(1),
                BackColor = SystemColors.ControlDark
            };
            if (content != null)
            {
                content.Dock = DockStyle.Fill;
                frame.Controls.Add(content);
            }

            return frame;
        }
    }
}

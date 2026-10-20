using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>空白区中央的可点文字。外观与底色相接，没有系统按钮的凹凸边。</summary>
    internal static class HintButton
    {
        /// <summary>做一颗平按钮。鼠标移上去只把底色略加深，按下时不另画立体边。</summary>
        /// <param name="text">按钮上的字。</param>
        /// <returns>调用方负责加入容器、写悬浮说明并接上点击。</returns>
        public static Button Create(string text)
        {
            Button button = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.GrayText,
                Cursor = Cursors.Hand,
                TabStop = false
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.BorderColor = SystemColors.Window;
            button.FlatAppearance.MouseOverBackColor = SystemColors.Control;
            button.FlatAppearance.MouseDownBackColor = SystemColors.Control;
            return button;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>补齐 WinForms 下拉菜单的「点外面即收起」。</summary>
    /// <remarks>
    /// WinForms 靠本线程的鼠标消息判断是否点在菜单外面。WebView2 正文属于另一个进程，点在那里本线程收不到消息，
    /// 所以由宿主在收到页面通知时调用 <see cref="CloseAll"/>。另提供按钮与下拉配对的开合切换。
    /// 所有方法都必须在界面线程调用。
    /// </remarks>
    internal static class DropDownCloser
    {
        private delegate bool EnumWindowsProc(IntPtr handle, IntPtr state);

        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(int threadId, EnumWindowsProc callback, IntPtr state);

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();

        /// <summary>收起本线程上所有可见的 WinForms 下拉菜单，关闭原因记为点击程序其他位置。</summary>
        /// <remarks>只遍历本线程顶层窗口，窗口数很少，开销可忽略。</remarks>
        public static void CloseAll()
        {
            List<ToolStripDropDown> open = new List<ToolStripDropDown>();
            EnumThreadWindows(GetCurrentThreadId(), (handle, state) =>
            {
                ToolStripDropDown dropDown = Control.FromHandle(handle) as ToolStripDropDown;
                if (dropDown != null && dropDown.Visible)
                    open.Add(dropDown);
                return true;
            }, IntPtr.Zero);
            foreach (ToolStripDropDown dropDown in open)
                dropDown.Close(ToolStripDropDownCloseReason.AppClicked);
        }

        /// <summary>让按钮再点一次时收起它打开的下拉，而不是关掉后又立刻打开。</summary>
        /// <param name="button">打开下拉的按钮。</param>
        /// <param name="menu">按钮对应的下拉。</param>
        /// <param name="show">点按钮且下拉未打开时执行的打开动作。</param>
        /// <remarks>
        /// 鼠标按下时菜单先因点在外面被收起，按钮随后才收到按下和单击。
        /// 这里在收起时记下鼠标是否落在按钮上，按钮这次单击就不再打开。
        /// </remarks>
        public static void BindToggle(Button button, ToolStripDropDown menu, Action show)
        {
            bool closedOnButton = false;
            bool skipClick = false;
            menu.Closed += (sender, args) =>
            {
                closedOnButton = args.CloseReason == ToolStripDropDownCloseReason.AppClicked
                    && !button.IsDisposed
                    && button.ClientRectangle.Contains(button.PointToClient(Cursor.Position));
            };
            button.MouseDown += (sender, args) =>
            {
                skipClick = closedOnButton;
                closedOnButton = false;
            };
            button.Click += (sender, args) =>
            {
                if (skipClick)
                {
                    skipClick = false;
                    return;
                }

                show();
            };
            // 单击先于 MouseUp 事件；松开在按钮外面时不会有单击，这里清掉以免吞掉下一次键盘单击。
            button.MouseUp += (sender, args) => skipClick = false;
        }
    }
}

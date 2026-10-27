using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using DB2Sheet.Contracts;

namespace DB2Sheet.UI
{
    /// <summary>为插件所有 WinForms 窗体统一应用品牌字体和嵌入式图标。</summary>
    /// <remarks>功能窗体应继承此类而不是直接继承 <see cref="Form"/>，否则公共视觉配置不会生效。</remarks>
    public abstract class AppForm : Form
    {
        private static readonly Icon ApplicationIcon = LoadApplicationIcon();
        private bool _preparedFirstShow;

        /// <summary>初始化公共字体和窗体图标。</summary>
        protected AppForm()
        {
            Font = new Font(
                AppPresentation.DefaultFontName,
                AppPresentation.DefaultFontSize,
                FontStyle.Regular,
                GraphicsUnit.Point);
            Icon = ApplicationIcon;
        }

        /// <summary>
        /// 句柄已按最终尺寸创建，布局已完成。
        /// 子类在这里摆分割条或填第一帧数据，避免窗口画出后再改布局留下残影。
        /// </summary>
        /// <remarks>
        /// 只在 UI 线程、首次 <see cref="SetVisibleCore"/> 变为可见时调用一次。
        /// <see cref="DeferFirstShow"/> 为 true 时，窗口已经显示但完全透明，客户区是最终尺寸。
        /// </remarks>
        protected virtual void PrepareFirstShow()
        {
        }

        /// <summary>为 true 时，首次显示先保持完全透明，由 <see cref="RevealFirstShow"/> 再画出来。</summary>
        /// <remarks>只有打开时会改分栏或填表的窗体需要。进度窗和模态小窗保持 false，否则它们自己也会晚一拍。</remarks>
        protected virtual bool DeferFirstShow
        {
            get { return false; }
        }

        /// <summary>透明显示之前。子类可在这里挡住紧接着的激活，避免把准备阶段再做一遍。</summary>
        /// <remarks>只在界面线程、推迟显示的窗体首次显示时调用一次。</remarks>
        protected virtual void OnFirstShowOpening()
        {
        }

        /// <summary>窗口已经不透明并完成一次同步重画。子类在这里弹出准备阶段延后的提示。</summary>
        /// <remarks>只在界面线程由 <see cref="RevealFirstShow"/> 调用。此时进度窗应当已经关掉。</remarks>
        protected virtual void OnFirstShowRevealed()
        {
        }

        /// <summary>按标题、初始尺寸、最小尺寸和居中方式摆好窗体骨架。</summary>
        /// <param name="title">完整窗口标题。需要品牌前缀时由调用方先经 <see cref="AppPresentation.WindowTitle"/> 组合。</param>
        /// <param name="width">首次显示的宽度。</param>
        /// <param name="height">首次显示的高度。</param>
        /// <param name="minimumSize">用户可缩到的最小尺寸。</param>
        /// <param name="startPosition">居中方式。需要记住尺寸时改用带存储参数的重载。</param>
        /// <remarks>必须在创建句柄之前调用。只设置窗体属性，不创建控件。</remarks>
        protected void InitializeWindow(
            string title, int width, int height, Size minimumSize, FormStartPosition startPosition)
        {
            Text = title;
            Width = width;
            Height = height;
            MinimumSize = minimumSize;
            StartPosition = startPosition;
        }

        /// <summary>摆好会记住宽高的功能窗骨架，并在句柄创建前恢复上次尺寸。</summary>
        /// <param name="title">完整窗口标题。</param>
        /// <param name="width">没有历史尺寸时的宽度。</param>
        /// <param name="height">没有历史尺寸时的高度。</param>
        /// <param name="minimumSize">用户可缩到的最小尺寸。</param>
        /// <param name="store">会话设置存储。</param>
        /// <param name="sizeKey">该窗体在 <c>session.formBounds</c> 里的稳定键。</param>
        /// <remarks>
        /// 内部调用 <see cref="FormSizeMemory.Attach"/>，因此窗体位置改为按上次保存的值或工作区居中。
        /// 必须在创建句柄之前、添加控件之前调用。
        /// </remarks>
        protected void InitializeWindow(
            string title, int width, int height, Size minimumSize, ISettingsStore store, string sizeKey)
        {
            InitializeWindow(title, width, height, minimumSize, FormStartPosition.Manual);
            FormSizeMemory.Attach(this, store, sizeKey);
        }

        /// <summary>建立统一内边距的单列根表格，并按顺序加入行样式。</summary>
        /// <param name="rows">各行的高。省略时只留一行铺满。</param>
        /// <returns>可直接加入窗体的根表格。调用方继续添加行样式和控件。</returns>
        /// <remarks>自动加一列铺满可用宽度，窗口缩小时内容跟着收窄。</remarks>
        protected static TableLayoutPanel CreateRootTable(params RowStyle[] rows)
        {
            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = rows == null || rows.Length == 0 ? 1 : rows.Length
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            if (rows == null || rows.Length == 0)
            {
                root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
                return root;
            }

            foreach (RowStyle row in rows)
                root.RowStyles.Add(row);
            return root;
        }

        /// <summary>建立统一留白的底部按钮栏，按钮按传入顺序从右往左排列。</summary>
        /// <param name="buttons">按钮或夹在按钮之间的说明控件。第一个显示在最右端。</param>
        /// <returns>可直接放进根表格最后一行的按钮栏。</returns>
        /// <remarks>栏高随按钮自适应；上方留 8 像素，与内容分开。</remarks>
        protected static FlowLayoutPanel CreateActionBar(params Control[] buttons)
        {
            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Margin = new Padding(0),
                Padding = new Padding(0, 8, 0, 0)
            };
            if (buttons != null)
                actions.Controls.AddRange(buttons);
            return actions;
        }

        /// <summary>结束首次显示的透明阶段，把窗口画出来。</summary>
        /// <remarks>
        /// 必须在界面线程、关掉打开用的进度窗之后调用。会立刻重画，再调用 <see cref="OnFirstShowRevealed"/>。
        /// 透明度留在 1，分层窗口样式会去掉，避免编辑器和浏览器花屏。
        /// </remarks>
        internal void RevealFirstShow()
        {
            if (IsDisposed) return;
            if (Opacity < 1)
                Opacity = 1;
            Refresh();
            OnFirstShowRevealed();
        }

        /// <summary>
        /// 首次显示前先创建句柄并完成布局，再调用 <see cref="PrepareFirstShow"/>。
        /// 推迟显示的窗体这次保持完全透明，调用方再调用 <see cref="RevealFirstShow"/>。
        /// </summary>
        /// <param name="value">为 true 时显示窗体。</param>
        /// <remarks>
        /// 推迟显示时要先进入可见状态，最大化后的客户区才会确定。透明期间完成分栏和填表，用户看不到半成品。
        /// </remarks>
        protected override void SetVisibleCore(bool value)
        {
            if (value && !_preparedFirstShow)
            {
                _preparedFirstShow = true;
                // 句柄还没建：先把窗口收进当前屏幕的工作区，小屏和高缩放下也看得到底部按钮。
                FormSizeMemory.ClampToWorkingArea(this);
                if (DeferFirstShow)
                {
                    Opacity = 0;
                    if (!IsHandleCreated)
                        CreateHandle();
                    OnFirstShowOpening();
                    base.SetVisibleCore(true);
                    PerformLayout();
                    PrepareFirstShow();
                    Refresh();
                    return;
                }

                if (!IsHandleCreated)
                    CreateHandle();
                PerformLayout();
                PrepareFirstShow();
            }

            base.SetVisibleCore(value);
        }

        private static Icon LoadApplicationIcon()
        {
            using (Stream stream = typeof(AppForm).Assembly.GetManifestResourceStream(AppPresentation.IconResourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException("找不到嵌入的窗体图标资源：" + AppPresentation.IconResourceName);

                using (Icon icon = new Icon(stream))
                {
                    return (Icon)icon.Clone();
                }
            }
        }
    }
}

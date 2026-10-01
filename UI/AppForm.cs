using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

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
        /// 句柄已按最终尺寸创建，布局已完成，窗口仍未显示。
        /// 子类在这里摆分割条或填第一帧数据，避免窗口画出后再改布局留下残影。
        /// </summary>
        /// <remarks>只在 UI 线程、首次 <see cref="SetVisibleCore"/> 变为可见时调用一次。</remarks>
        protected virtual void PrepareFirstShow()
        {
        }

        /// <summary>
        /// 首次显示前先创建仍隐藏的句柄并完成布局，再调用 <see cref="PrepareFirstShow"/>。
        /// 窗口画到屏幕上时已经是最终尺寸。
        /// </summary>
        /// <param name="value">为 true 时显示窗体。</param>
        protected override void SetVisibleCore(bool value)
        {
            if (value && !_preparedFirstShow)
            {
                _preparedFirstShow = true;
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

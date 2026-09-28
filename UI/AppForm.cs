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

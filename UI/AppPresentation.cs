using System;
using System.Drawing;
using System.Reflection;

namespace DB2Sheet.UI
{
    /// <summary>集中定义所有用户可见品牌名称、版本展示和公共 UI 字体。</summary>
    /// <remarks>
    /// 更换插件展示名称时修改 <see cref="DisplayName"/> 即可。程序集名、持久化目录和 Ribbon 控件 ID 是独立的技术标识。
    /// </remarks>
    internal static class AppPresentation
    {
        public const string DisplayName = "SQL查询";
        public const string IconResourceName = "DB2Sheet.Resources.rocket.ico";
        public const string DefaultFontName = "Microsoft YaHei UI";
        public const float DefaultFontSize = 9F;
        public const string LogFontName = "Consolas";
        public const float LogFontSize = 9F;
        public const string CodeFontName = "Microsoft YaHei UI Light";
        public const float CodeFontSize = 9.5F;

        /// <summary>读取当前程序集的文件版本，供 Ribbon 等界面展示。</summary>
        /// <remarks>优先使用 AssemblyFileVersion，不读取程序集文件路径，避免 VSTO 缓存路径读不到版本。特性缺失时回退程序集版本号。</remarks>
        public static string FileVersion
        {
            get
            {
                Assembly assembly = typeof(AppPresentation).Assembly;
                AssemblyFileVersionAttribute attribute = Attribute.GetCustomAttribute(
                    assembly,
                    typeof(AssemblyFileVersionAttribute)) as AssemblyFileVersionAttribute;
                if (attribute != null && !string.IsNullOrWhiteSpace(attribute.Version))
                    return attribute.Version;

                return assembly.GetName().Version.ToString();
            }
        }

        /// <summary>为 SQL 编辑器创建中英文同一字面高度的代码字体；优先细体雅黑，缺失时回退。</summary>
        /// <param name="style">字重样式；关键字使用粗体。</param>
        /// <returns>已解析到本机已安装字体族的字体实例。</returns>
        public static Font CreateCodeFont(FontStyle style)
        {
            string[] candidates = { "Microsoft YaHei UI Light", "DengXian Light", "Microsoft YaHei UI" };
            foreach (string name in candidates)
            {
                using (Font probe = new Font(name, CodeFontSize, style, GraphicsUnit.Point))
                {
                    if (!string.Equals(probe.FontFamily.Name, name, StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                return new Font(name, CodeFontSize, style, GraphicsUnit.Point);
            }

            return new Font(DefaultFontName, CodeFontSize, style, GraphicsUnit.Point);
        }

        /// <summary>生成统一格式的功能窗体标题。</summary>
        /// <param name="featureName">功能名称；为空时只返回产品展示名称。</param>
        /// <returns>“产品名 - 功能名”格式的标题。</returns>
        public static string WindowTitle(string featureName)
        {
            return string.IsNullOrWhiteSpace(featureName)
                ? DisplayName
                : DisplayName + " - " + featureName;
        }
    }
}

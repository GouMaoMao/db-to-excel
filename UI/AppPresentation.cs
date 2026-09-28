namespace DB2Sheet.UI
{
    /// <summary>集中定义所有用户可见品牌名称和公共 UI 字体。</summary>
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
        public const string CodeFontName = "Consolas";
        public const float CodeFontSize = 10F;

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

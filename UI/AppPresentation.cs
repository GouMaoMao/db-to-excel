using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>集中定义所有用户可见品牌名称、版本展示和公共 UI 字体。</summary>
    /// <remarks>
    /// 更换插件展示名称时修改 <see cref="DisplayName"/> 即可。程序集名、持久化目录和 Ribbon 控件 ID 是独立的技术标识。
    /// 连接编辑窗的类型标题、说明、悬停和空框提示也放在此处，避免散落在各 Provider。
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

        /// <summary>连接编辑窗下拉的固定顺序。新建默认取第一项 SQL Server。</summary>
        public static readonly string[] ConnectionProviderOrder =
        {
            "sqlserver",
            "mysql",
            "postgresql",
            "sqlite",
            "duckdb",
            "jdbc"
        };

        /// <summary>显示名称空框提示。</summary>
        public const string ConnectionNameCue = "例如：ERP正式环境、数据中台XX空间";

        /// <summary>服务器地址空框提示。</summary>
        public const string ConnectionHostCue = "主机名或 IP";

        /// <summary>JDBC URL 下方的示例行。</summary>
        public const string JdbcUrlExample = "示例：jdbc:oracle:thin:@主机:1521:库名";

        /// <summary>驱动 jar 旁的简短说明。</summary>
        public const string DriverJarHint = "向管理员或厂商要 .jar";

        /// <summary>未测试成功时保存按钮旁的说明。</summary>
        public const string SaveRequiresTestHint = "测试成功后才能保存";

        /// <summary>选中 JDBC 且未找到 Java 时的横幅文案。</summary>
        public const string JdbcJavaBanner =
            "其他数据库（JDBC）需要本机 Java（JDK 8 或更高）。请先配置 JDBC 环境，再测试连接。";

        private const int EmSetCueBanner = 0x1501;

        /// <summary>读取当前程序集的产品版本，供 Ribbon 等界面展示。</summary>
        /// <remarks>优先使用三段的 AssemblyFileVersion，不读取程序集文件路径，避免 VSTO 缓存路径读不到版本。特性缺失时回退程序集版本号。</remarks>
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

        /// <summary>为 SQL 编辑器创建中英文同一字面高度的常规代码字体；优先细体雅黑，缺失时回退。</summary>
        /// <returns>已解析到本机已安装字体族的字体实例。调用方负责释放。</returns>
        public static Font CreateCodeFont()
        {
            string[] candidates = { "Microsoft YaHei UI Light", "DengXian Light", "Microsoft YaHei UI" };
            foreach (string name in candidates)
            {
                using (Font probe = new Font(name, CodeFontSize, FontStyle.Regular, GraphicsUnit.Point))
                {
                    if (!string.Equals(probe.FontFamily.Name, name, StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                return new Font(name, CodeFontSize, FontStyle.Regular, GraphicsUnit.Point);
            }

            return new Font(DefaultFontName, CodeFontSize, FontStyle.Regular, GraphicsUnit.Point);
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

        /// <summary>连接类型在界面上的标题。JDBC 显示为「其他数据库（JDBC）」。</summary>
        /// <param name="providerId">Provider 标识。</param>
        /// <returns>用户可见标题；未知标识时回退原值。</returns>
        public static string ConnectionProviderTitle(string providerId)
        {
            if (string.IsNullOrWhiteSpace(providerId)) return string.Empty;
            switch (providerId.Trim().ToLowerInvariant())
            {
                case "sqlserver": return "SQL Server";
                case "mysql": return "MySQL";
                case "postgresql": return "PostgreSQL";
                case "sqlite": return "SQLite";
                case "duckdb": return "DuckDB";
                case "jdbc": return "其他数据库（JDBC）";
                default: return providerId;
            }
        }

        /// <summary>选中某连接类型时显示在下拉下方的一句说明。</summary>
        /// <param name="providerId">Provider 标识。</param>
        /// <returns>白话说明。服务器型共用同一句。</returns>
        public static string ConnectionProviderHint(string providerId)
        {
            if (string.IsNullOrWhiteSpace(providerId)) return string.Empty;
            switch (providerId.Trim().ToLowerInvariant())
            {
                case "sqlserver":
                case "mysql":
                case "postgresql":
                    return "填服务器、库名和账号。";
                case "sqlite":
                case "duckdb":
                    return "本机一个文件，点浏览即可。";
                case "jdbc":
                    return "上面没有的库，需要驱动 jar 和本机 Java。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>连接编辑字段的悬停说明。</summary>
        /// <param name="fieldKey">字段键，如 host、jdbcUrl，或 name、provider、test、save。</param>
        /// <returns>一句白话；未知键为空。</returns>
        public static string ConnectionFieldTip(string fieldKey)
        {
            if (string.IsNullOrWhiteSpace(fieldKey)) return string.Empty;
            switch (fieldKey.Trim())
            {
                case "name":
                    return "在连接列表里显示的名字，方便认出环境。";
                case "provider":
                    return "要连接的数据库种类；常见库选上面几项，没有再选其他（JDBC）。";
                case "host":
                    return "数据库所在机器的主机名或 IP。";
                case "port":
                    return "一般用默认即可，除非管理员另有要求。";
                case "database":
                    return "要查询的库名。";
                case "userName":
                    return "登录该库的账号；勾选 Windows 集成认证时可留空。";
                case "password":
                    return "登录该库的密码；勾选 Windows 集成认证时可留空。";
                case "integratedSecurity":
                    return "用当前 Windows 账号登录，不必再填用户名密码。";
                case "filePath":
                    return "本机数据库文件路径，可点浏览选择。";
                case "encrypt":
                    return "是否加密传输。一般保持默认，除非管理员有要求。";
                case "trustServerCertificate":
                    return "开发或内网环境可勾选，以信任服务器自签证书。";
                case "jdbcUrl":
                    return "厂商要求的连接串，向管理员索取。";
                case "driverJars":
                    return "厂商提供的驱动文件，向管理员或厂商索取。";
                case "driverClass":
                    return "一般点「检测驱动类」即可；检测不到再手工填。";
                case "test":
                    return "用当前填写内容试连；通过后才能保存。";
                case "save":
                    return "测试通过后才能保存；保存后可在查询里使用。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>给文本框设置灰色空框提示。句柄尚未创建时等创建后再设。</summary>
        /// <param name="textBox">目标文本框。</param>
        /// <param name="cue">提示文字。</param>
        public static void SetCueBanner(TextBox textBox, string cue)
        {
            if (textBox == null || string.IsNullOrEmpty(cue)) return;
            if (textBox.IsHandleCreated)
            {
                SendMessage(textBox.Handle, EmSetCueBanner, (IntPtr)1, cue);
                return;
            }

            EventHandler onCreated = null;
            onCreated = (sender, args) =>
            {
                textBox.HandleCreated -= onCreated;
                if (!textBox.IsDisposed && textBox.IsHandleCreated)
                {
                    SendMessage(textBox.Handle, EmSetCueBanner, (IntPtr)1, cue);
                }
            };
            textBox.HandleCreated += onCreated;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
    }
}

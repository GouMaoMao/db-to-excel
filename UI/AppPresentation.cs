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
    /// 连接编辑窗的类型标题、类型介绍、框下说明、悬停和空框提示，以及参数设置窗的框下说明，也放在此处。
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

        /// <summary>JDBC URL 下方说明的第一句。</summary>
        public const string JdbcUrlNote =
            "厂商规定的完整连接地址，写明连哪台机器、哪个库。向管理员索取后原样粘贴。";

        /// <summary>JDBC URL 下方的示例行。</summary>
        public const string JdbcUrlExample = "示例：jdbc:oracle:thin:@主机:1521:库名";

        /// <summary>驱动 jar 下方的说明。</summary>
        public const string DriverJarNote =
            "JDBC 驱动程序，数据库厂商提供的 .jar 文件。Java 加载它，才能按该库的协议建立连接。驱动含多个 jar 时全部添加。";

        /// <summary>驱动类下方的说明。</summary>
        public const string DriverClassNote =
            "该驱动程序中供 JDBC 加载的类名，例如 oracle.jdbc.OracleDriver。点「检测驱动类」从 jar 中读取；jar 未声明时按厂商文档填写。";

        /// <summary>JDBC 用户名下方的说明。</summary>
        public const string JdbcUserNameNote = "登录账号。有的库把账号写在 JDBC URL 里，这里可以留空。";

        /// <summary>jar 中读不到驱动类时的提示。</summary>
        public const string DriverClassMissingMessage = "jar 未声明驱动类，请按厂商文档填写。";

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

        /// <summary>选中某连接类型时显示在下拉下方的一句介绍，只说这种库的特点。</summary>
        /// <param name="providerId">Provider 标识。</param>
        /// <returns>一句介绍；未知标识为空。</returns>
        public static string ConnectionProviderHint(string providerId)
        {
            if (string.IsNullOrWhiteSpace(providerId)) return string.Empty;
            switch (providerId.Trim().ToLowerInvariant())
            {
                case "sqlserver":
                    return "微软的企业数据库，常见于公司业务系统，可用当前 Windows 账号登录。";
                case "mysql":
                    return "使用最广的开源数据库，多见于网站和一般业务系统。";
                case "postgresql":
                    return "开源数据库，擅长复杂查询和较严格的数据约束。";
                case "sqlite":
                    return "不需要数据库服务，整个库是本机上的一个文件。";
                case "duckdb":
                    return "本机上的一个文件，专为分析查询设计，适合统计和汇总。";
                case "jdbc":
                    return "通用连接方式，用来连接这里没有单独列出的数据库，例如 Oracle，以及 Hive 等数据仓库。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>输入框下方的灰色说明。没有说明的字段返回空，窗体不另占一行。</summary>
        /// <param name="providerId">当前数据源类型。用户名说明只在 JDBC 下出现。</param>
        /// <param name="fieldKey">字段键，如 jdbcUrl、driverJars、integratedSecurity。</param>
        /// <returns>框下说明；不需要说明时为空。</returns>
        public static string ConnectionFieldNote(string providerId, string fieldKey)
        {
            if (string.IsNullOrWhiteSpace(fieldKey)) return string.Empty;
            bool jdbc = string.Equals(providerId, "jdbc", StringComparison.OrdinalIgnoreCase);
            switch (fieldKey.Trim())
            {
                case "jdbcUrl":
                    return JdbcUrlNote + Environment.NewLine + JdbcUrlExample;
                case "driverJars":
                    return DriverJarNote;
                case "driverClass":
                    return DriverClassNote;
                case "userName":
                    return jdbc ? JdbcUserNameNote : string.Empty;
                case "integratedSecurity":
                    return "勾选后用当前 Windows 账号登录，用户名和密码可留空。";
                case "encrypt":
                    return "传输是否加密。没有管理员要求时保持默认。";
                case "trustServerCertificate":
                    return "内网或自签证书连不上时再勾选。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>连接编辑字段的悬停说明。已有框下说明的字段与说明使用同一句话。</summary>
        /// <param name="providerId">当前数据源类型，用于区分 JDBC 用户名和其他库的用户名。</param>
        /// <param name="fieldKey">字段键，如 host、jdbcUrl，或 name、provider、test、save。</param>
        /// <returns>悬停说明；未知键为空。</returns>
        public static string ConnectionFieldTip(string providerId, string fieldKey)
        {
            string note = ConnectionFieldNote(providerId, fieldKey);
            if (!string.IsNullOrEmpty(note)) return note;
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
                case "filePath":
                    return "本机数据库文件路径，可点浏览选择。";
                case "test":
                    return "用当前填写内容试连；通过后才能保存。";
                case "save":
                    return "测试通过后才能保存；保存后可在查询里使用。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>参数设置窗里某一项下方的灰色说明。会话状态不出现在设置窗，这里也不写。</summary>
        /// <param name="key">设置键，与 <c>CoreSettings</c> 中的持久化键相同。</param>
        /// <returns>一句定义；未知键为空，窗体不另占一行。</returns>
        public static string SettingNote(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;
            switch (key.Trim())
            {
                case "query.maxPreviewRows":
                    return "SQL 查询窗预览时最多读取并显示的行数。";
                case "excel.maxExportRows":
                    return "写入 Excel 的结果数据行上限，不含标题行。超出的行被截断。";
                case "batch.mode":
                    return "Serial 为串行，一项查询结束后再查下一项。Parallel 为并行，多项同时查询。";
                case "batch.maxParallelism":
                    return "并行时同时查询的任务个数。串行时不起作用。";
                case "query.timeoutSeconds":
                    return "单条查询允许运行的最长时间，超时后该查询失败。";
                case "query.resultBlockSize":
                    return "从数据库读取结果时每次取回的行数。只影响读取节奏，不改变最终行数。";
                case "logging.level":
                    return "进度窗口只显示不低于所选级别的日志。Debug 最详细，Error 只显示错误。写入文件的日志仍保留全部级别。";
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

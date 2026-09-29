using System;

using System.Collections.Generic;

using System.Linq;

using DB2Sheet.Models;



namespace DB2Sheet.Providers

{

    /// <summary>描述一条 JDBC 连接在执行前解析出的 URL、驱动 jar、驱动类和凭据。</summary>

    internal sealed class JdbcConnectionTarget

    {

        /// <summary>创建执行用的连接目标。</summary>

        /// <param name="url">JDBC URL。</param>

        /// <param name="driverJars">厂商驱动 jar 的绝对路径。</param>

        /// <param name="driverClass">驱动类名。</param>

        /// <param name="userName">用户名或 AccessKey ID。</param>

        /// <param name="password">密码或 AccessKey Secret。</param>

        /// <param name="catalog">要切换到的目录或架构；为空表示保持 URL 中的默认库。</param>

        public JdbcConnectionTarget(

            string url,

            IReadOnlyList<string> driverJars,

            string driverClass,

            string userName,

            string password,

            string catalog)

        {

            Url = url ?? string.Empty;

            DriverJars = (driverJars ?? new string[0]).Where(item => !string.IsNullOrWhiteSpace(item)).Select(item => item.Trim()).ToList().AsReadOnly();

            DriverClass = driverClass ?? string.Empty;

            UserName = userName ?? string.Empty;

            Password = password ?? string.Empty;

            Catalog = catalog ?? string.Empty;

        }



        /// <summary>获取 JDBC URL。</summary>

        public string Url { get; }



        /// <summary>获取厂商驱动 jar 路径列表。</summary>

        public IReadOnlyList<string> DriverJars { get; }



        /// <summary>获取驱动类名。</summary>

        public string DriverClass { get; }



        /// <summary>获取用户名或 AccessKey ID。</summary>

        public string UserName { get; }



        /// <summary>获取密码或 AccessKey Secret。</summary>

        public string Password { get; }



        /// <summary>获取目录或架构名称。</summary>

        public string Catalog { get; }

    }



    /// <summary>把 JDBC 连接参数整理成执行目标。</summary>

    /// <remarks>连接编辑填写完整 JDBC URL、驱动 jar、驱动类和凭据。</remarks>

    internal static class JdbcConnectionComposer

    {

        /// <summary>JDBC 提供程序的稳定标识。</summary>

        public const string ProviderId = "jdbc";



        /// <summary>连接参数中多条 jar 路径的分隔符。</summary>

        public const char DriverJarSeparator = '|';



        private static readonly IReadOnlyList<ParameterDefinition> Definitions = new List<ParameterDefinition>

        {

            new ParameterDefinition(JdbcParameterKeys.JdbcUrl, "JDBC URL", ParameterValueType.Text, true),

            new ParameterDefinition(JdbcParameterKeys.DriverJars, "驱动 jar", ParameterValueType.Text, true),

            new ParameterDefinition(JdbcParameterKeys.DriverClass, "驱动类", ParameterValueType.Text, true),

            new ParameterDefinition(DatabaseParameterKeys.UserName, "用户名", ParameterValueType.Text, false),

            new ParameterDefinition(DatabaseParameterKeys.Password, "密码", ParameterValueType.Password, false, null, true)

        }.AsReadOnly();



        /// <summary>获取连接编辑窗使用的参数定义。</summary>

        public static IReadOnlyList<ParameterDefinition> Parameters => Definitions;



        /// <summary>检查必填字段。</summary>

        /// <param name="profile">连接快照。</param>

        /// <returns>错误消息；空列表表示可以通过。</returns>

        public static IReadOnlyList<string> Validate(ConnectionProfileSnapshot profile)

        {

            List<string> errors = new List<string>();

            if (profile == null)

            {

                errors.Add("连接方案不能为空。");

                return errors;

            }



            if (!string.Equals(profile.ProviderId, ProviderId, StringComparison.OrdinalIgnoreCase))

            {

                errors.Add("连接方案与数据源类型不匹配。");

            }



            Require(errors, profile, JdbcParameterKeys.JdbcUrl, "JDBC URL");

            if (SplitDriverJars(profile.GetValue(JdbcParameterKeys.DriverJars)).Count == 0)

            {

                errors.Add("请至少添加一个驱动 jar。");

            }



            Require(errors, profile, JdbcParameterKeys.DriverClass, "驱动类");

            return errors;

        }



        /// <summary>生成执行目标。</summary>

        /// <param name="profile">连接快照。</param>

        /// <returns>URL、驱动 jar、驱动类和凭据。</returns>

        public static JdbcConnectionTarget Compose(ConnectionProfileSnapshot profile)

        {

            if (profile == null) throw new ArgumentNullException(nameof(profile));

            return new JdbcConnectionTarget(

                profile.GetValue(JdbcParameterKeys.JdbcUrl).Trim(),

                SplitDriverJars(profile.GetValue(JdbcParameterKeys.DriverJars)),

                profile.GetValue(JdbcParameterKeys.DriverClass).Trim(),

                profile.GetValue(DatabaseParameterKeys.UserName),

                profile.GetValue(DatabaseParameterKeys.Password),

                profile.GetValue(JdbcParameterKeys.Catalog));

        }



        /// <summary>把参数中的 jar 文本拆成路径列表。</summary>

        /// <param name="value">以竖线分隔的路径文本。</param>

        /// <returns>非空路径列表。</returns>

        public static IReadOnlyList<string> SplitDriverJars(string value)

        {

            if (string.IsNullOrWhiteSpace(value))

            {

                return new string[0];

            }



            return value.Split(new[] { DriverJarSeparator }, StringSplitOptions.RemoveEmptyEntries)

                .Select(item => item.Trim())

                .Where(item => item.Length > 0)

                .Distinct(StringComparer.OrdinalIgnoreCase)

                .ToList()

                .AsReadOnly();

        }



        /// <summary>把 jar 路径列表编码为连接参数文本。</summary>

        /// <param name="jars">路径列表。</param>

        /// <returns>竖线分隔的文本。</returns>

        public static string JoinDriverJars(IEnumerable<string> jars)

        {

            if (jars == null) return string.Empty;

            return string.Join(

                DriverJarSeparator.ToString(),

                jars.Where(item => !string.IsNullOrWhiteSpace(item)).Select(item => item.Trim()).Distinct(StringComparer.OrdinalIgnoreCase));

        }



        private static void Require(List<string> errors, ConnectionProfileSnapshot profile, string key, string displayName)

        {

            if (string.IsNullOrWhiteSpace(profile.GetValue(key)))

            {

                errors.Add(displayName + "不能为空。");

            }

        }

    }



    /// <summary>集中定义 JDBC 连接参数键。这些字符串会写入连接方案。</summary>

    internal static class JdbcParameterKeys

    {

        /// <summary>完整 JDBC URL。</summary>

        public const string JdbcUrl = "jdbcUrl";



        /// <summary>竖线分隔的驱动 jar 绝对路径。</summary>

        public const string DriverJars = "driverJars";



        /// <summary>驱动类名。</summary>

        public const string DriverClass = "driverClass";



        /// <summary>从数据库树选择后临时写入快照的目录或架构名，不在编辑窗中显示。</summary>

        public const string Catalog = "jdbcCatalog";

    }

}


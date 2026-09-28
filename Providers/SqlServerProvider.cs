using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using DB2Sheet.Models;

namespace DB2Sheet.Providers
{
    /// <summary>使用 System.Data.SqlClient 为 SQL Server 提供只读意图连接和查询能力。</summary>
    /// <remarks>支持 Windows 集成认证或用户名密码认证，并可配置传输加密和证书信任。</remarks>
    public sealed class SqlServerProvider : DatabaseProviderBase
    {
        private static readonly IReadOnlyList<ParameterDefinition> Parameters = new List<ParameterDefinition>
        {
            new ParameterDefinition(DatabaseParameterKeys.Host, "服务器", ParameterValueType.Text, true, "localhost"),
            new ParameterDefinition(DatabaseParameterKeys.Port, "端口", ParameterValueType.Integer, false, "1433"),
            new ParameterDefinition(DatabaseParameterKeys.Database, "数据库", ParameterValueType.Text, true),
            new ParameterDefinition(DatabaseParameterKeys.IntegratedSecurity, "Windows 集成认证", ParameterValueType.Boolean, false, "false"),
            new ParameterDefinition(DatabaseParameterKeys.UserName, "用户名", ParameterValueType.Text),
            new ParameterDefinition(DatabaseParameterKeys.Password, "密码", ParameterValueType.Password, false, null, true),
            new ParameterDefinition(DatabaseParameterKeys.Encrypt, "加密连接", ParameterValueType.Boolean, false, "true"),
            new ParameterDefinition(DatabaseParameterKeys.TrustServerCertificate, "信任服务器证书", ParameterValueType.Boolean, false, "false")
        }.AsReadOnly();

        /// <inheritdoc/>
        public override string ProviderId => "sqlserver";
        /// <inheritdoc/>
        public override string DisplayName => "SQL Server";
        /// <inheritdoc/>
        public override IReadOnlyList<ParameterDefinition> ConnectionParameters => Parameters;

        /// <inheritdoc/>
        public override DbConnection CreateConnection(ConnectionProfileSnapshot profile)
        {
            int port = GetPort(profile, 1433);
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder
            {
                DataSource = port == 1433 ? profile.GetValue(DatabaseParameterKeys.Host) : profile.GetValue(DatabaseParameterKeys.Host) + "," + port,
                InitialCatalog = profile.GetValue(DatabaseParameterKeys.Database),
                IntegratedSecurity = GetBoolean(profile, DatabaseParameterKeys.IntegratedSecurity),
                Encrypt = GetBoolean(profile, DatabaseParameterKeys.Encrypt, true),
                TrustServerCertificate = GetBoolean(profile, DatabaseParameterKeys.TrustServerCertificate),
                ApplicationName = "DB2Sheet",
                ApplicationIntent = ApplicationIntent.ReadOnly
            };

            if (!builder.IntegratedSecurity)
            {
                builder.UserID = profile.GetValue(DatabaseParameterKeys.UserName);
                builder.Password = profile.GetValue(DatabaseParameterKeys.Password);
            }

            return new SqlConnection(builder.ConnectionString);
        }
    }
}

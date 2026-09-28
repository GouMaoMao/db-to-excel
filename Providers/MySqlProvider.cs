using System.Collections.Generic;
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;
using MySqlConnector;

namespace DB2Sheet.Providers
{
    /// <summary>使用 MySqlConnector 为 MySQL 提供只读查询和连接测试能力。</summary>
    /// <remarks>查询前通过会话命令启用只读事务约束。</remarks>
    public sealed class MySqlProvider : DatabaseProviderBase
    {
        private static readonly IReadOnlyList<ParameterDefinition> Parameters = new List<ParameterDefinition>
        {
            new ParameterDefinition(DatabaseParameterKeys.Host, "服务器", ParameterValueType.Text, true, "localhost"),
            new ParameterDefinition(DatabaseParameterKeys.Port, "端口", ParameterValueType.Integer, false, "3306"),
            new ParameterDefinition(DatabaseParameterKeys.Database, "数据库", ParameterValueType.Text, true),
            new ParameterDefinition(DatabaseParameterKeys.UserName, "用户名", ParameterValueType.Text, true),
            new ParameterDefinition(DatabaseParameterKeys.Password, "密码", ParameterValueType.Password, false, null, true)
        }.AsReadOnly();

        /// <inheritdoc/>
        public override string ProviderId => "mysql";
        /// <inheritdoc/>
        public override string DisplayName => "MySQL";
        /// <inheritdoc/>
        public override IReadOnlyList<ParameterDefinition> ConnectionParameters => Parameters;
        /// <inheritdoc/>
        public override ProviderCapabilities Capabilities => base.Capabilities | ProviderCapabilities.ReadOnlySession;

        /// <inheritdoc/>
        public override DbConnection CreateConnection(ConnectionProfileSnapshot profile)
        {
            MySqlConnectionStringBuilder builder = new MySqlConnectionStringBuilder
            {
                Server = profile.GetValue(DatabaseParameterKeys.Host),
                Port = (uint)GetPort(profile, 3306),
                Database = profile.GetValue(DatabaseParameterKeys.Database),
                UserID = profile.GetValue(DatabaseParameterKeys.UserName),
                Password = profile.GetValue(DatabaseParameterKeys.Password),
                ApplicationName = "DB2Sheet"
            };
            return new MySqlConnection(builder.ConnectionString);
        }

        /// <inheritdoc/>
        public override Task ConfigureReadOnlySessionAsync(DbConnection connection, CancellationToken cancellationToken)
        {
            return ExecuteSessionCommandAsync(connection, "SET SESSION TRANSACTION READ ONLY", cancellationToken);
        }
    }
}

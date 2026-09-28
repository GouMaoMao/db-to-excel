using System.Collections.Generic;
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;
using Npgsql;

namespace DB2Sheet.Providers
{
    /// <summary>使用 Npgsql 为 PostgreSQL 提供只读查询和连接测试能力。</summary>
    /// <remarks>查询前将当前会话的默认事务设置为只读。</remarks>
    public sealed class PostgreSqlProvider : DatabaseProviderBase
    {
        private static readonly IReadOnlyList<ParameterDefinition> Parameters = new List<ParameterDefinition>
        {
            new ParameterDefinition(DatabaseParameterKeys.Host, "服务器", ParameterValueType.Text, true, "localhost"),
            new ParameterDefinition(DatabaseParameterKeys.Port, "端口", ParameterValueType.Integer, false, "5432"),
            new ParameterDefinition(DatabaseParameterKeys.Database, "数据库", ParameterValueType.Text, true),
            new ParameterDefinition(DatabaseParameterKeys.UserName, "用户名", ParameterValueType.Text, true),
            new ParameterDefinition(DatabaseParameterKeys.Password, "密码", ParameterValueType.Password, false, null, true)
        }.AsReadOnly();

        /// <inheritdoc/>
        public override string ProviderId => "postgresql";
        /// <inheritdoc/>
        public override string DisplayName => "PostgreSQL";
        /// <inheritdoc/>
        public override IReadOnlyList<ParameterDefinition> ConnectionParameters => Parameters;
        /// <inheritdoc/>
        public override ProviderCapabilities Capabilities => base.Capabilities | ProviderCapabilities.ReadOnlySession;

        /// <inheritdoc/>
        public override DbConnection CreateConnection(ConnectionProfileSnapshot profile)
        {
            NpgsqlConnectionStringBuilder builder = new NpgsqlConnectionStringBuilder
            {
                Host = profile.GetValue(DatabaseParameterKeys.Host),
                Port = GetPort(profile, 5432),
                Database = profile.GetValue(DatabaseParameterKeys.Database),
                Username = profile.GetValue(DatabaseParameterKeys.UserName),
                Password = profile.GetValue(DatabaseParameterKeys.Password),
                ApplicationName = "DB2Sheet"
            };
            return new NpgsqlConnection(builder.ConnectionString);
        }

        /// <inheritdoc/>
        public override Task ConfigureReadOnlySessionAsync(DbConnection connection, CancellationToken cancellationToken)
        {
            return ExecuteSessionCommandAsync(connection, "SET default_transaction_read_only = on", cancellationToken);
        }
    }
}

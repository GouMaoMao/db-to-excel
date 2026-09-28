using System.Collections.Generic;
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SQLite;
using DB2Sheet.Models;

namespace DB2Sheet.Providers
{
    /// <summary>使用 System.Data.SQLite 以只读方式查询本地 SQLite 数据库文件。</summary>
    /// <remarks>连接字符串禁止创建缺失文件，并通过 PRAGMA 再次启用只读查询模式。</remarks>
    public sealed class SqliteProvider : DatabaseProviderBase
    {
        private static readonly IReadOnlyList<ParameterDefinition> Parameters = new List<ParameterDefinition>
        {
            new ParameterDefinition(DatabaseParameterKeys.FilePath, "数据库文件", ParameterValueType.FilePath, true)
        }.AsReadOnly();

        /// <inheritdoc/>
        public override string ProviderId => "sqlite";
        /// <inheritdoc/>
        public override string DisplayName => "SQLite";
        /// <inheritdoc/>
        public override IReadOnlyList<ParameterDefinition> ConnectionParameters => Parameters;
        /// <inheritdoc/>
        public override ProviderCapabilities Capabilities => base.Capabilities | ProviderCapabilities.ReadOnlySession;

        /// <inheritdoc/>
        public override DbConnection CreateConnection(ConnectionProfileSnapshot profile)
        {
            SQLiteConnectionStringBuilder builder = new SQLiteConnectionStringBuilder
            {
                DataSource = profile.GetValue(DatabaseParameterKeys.FilePath),
                ReadOnly = true,
                FailIfMissing = true
            };
            return new SQLiteConnection(builder.ConnectionString);
        }

        /// <inheritdoc/>
        public override Task ConfigureReadOnlySessionAsync(DbConnection connection, CancellationToken cancellationToken)
        {
            return ExecuteSessionCommandAsync(connection, "PRAGMA query_only = ON", cancellationToken);
        }
    }
}

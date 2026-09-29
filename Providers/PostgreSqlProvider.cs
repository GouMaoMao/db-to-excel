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
        public override async Task<IReadOnlyList<DatabaseMetadata>> GetDatabasesAsync(
            ConnectionProfileSnapshot profile,
            CancellationToken cancellationToken)
        {
            List<DatabaseMetadata> databases = new List<DatabaseMetadata>();
            using (DbConnection connection = CreateConnection(profile))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT datname FROM pg_database WHERE datallowconn AND has_database_privilege(datname, 'CONNECT') ORDER BY datname";
                    using (DbDataReader reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                    {
                        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                            databases.Add(new DatabaseMetadata(reader.GetString(0)));
                    }
                }
            }
            return databases.AsReadOnly();
        }

        /// <inheritdoc/>
        public override async Task<IReadOnlyList<DatabaseObjectMetadata>> GetDatabaseObjectsAsync(
            ConnectionProfileSnapshot profile,
            string databaseName,
            CancellationToken cancellationToken)
        {
            List<DatabaseObjectMetadata> objects = new List<DatabaseObjectMetadata>();
            using (DbConnection connection = CreateConnection(CreateDatabaseSnapshot(profile, databaseName)))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT table_schema, table_name, table_type FROM information_schema.tables WHERE table_catalog = current_database() ORDER BY table_schema, table_type, table_name";
                    using (DbDataReader reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                    {
                        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                        {
                            objects.Add(new DatabaseObjectMetadata(
                                reader.GetString(0),
                                reader.GetString(1),
                                reader.GetString(2).Equals("VIEW", System.StringComparison.OrdinalIgnoreCase)
                                    ? DatabaseObjectKind.View
                                    : DatabaseObjectKind.Table));
                        }
                    }
                }
            }
            return objects.AsReadOnly();
        }

        /// <inheritdoc/>
        public override Task ConfigureReadOnlySessionAsync(DbConnection connection, CancellationToken cancellationToken)
        {
            return ExecuteSessionCommandAsync(connection, "SET default_transaction_read_only = on", cancellationToken);
        }
    }
}

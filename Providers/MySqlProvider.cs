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
                    command.CommandText = "SELECT SCHEMA_NAME FROM INFORMATION_SCHEMA.SCHEMATA ORDER BY SCHEMA_NAME";
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
            using (DbConnection connection = CreateConnection(profile))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = @database ORDER BY TABLE_TYPE, TABLE_NAME";
                    DbParameter parameter = command.CreateParameter();
                    parameter.ParameterName = "@database";
                    parameter.Value = databaseName;
                    command.Parameters.Add(parameter);
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
            return ExecuteSessionCommandAsync(connection, "SET SESSION TRANSACTION READ ONLY", cancellationToken);
        }

        /// <inheritdoc/>
        protected override Task ApplyPreviewRowLimitAsync(
            DbConnection connection,
            int rowLimit,
            CancellationToken cancellationToken)
        {
            return ExecuteSessionCommandAsync(
                connection,
                "SET SQL_SELECT_LIMIT = " + rowLimit.ToString(System.Globalization.CultureInfo.InvariantCulture),
                cancellationToken);
        }
    }
}

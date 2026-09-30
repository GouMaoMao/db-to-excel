using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
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

        /// <summary>创建 SQL Server 提供程序。</summary>
        /// <param name="logger">注释缺口写入的文件日志。为空时不记录。</param>
        public SqlServerProvider(ILogger logger = null) : base(logger)
        {
        }

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
                    command.CommandText = "SELECT [name] FROM sys.databases WHERE [state] = 0 AND HAS_DBACCESS([name]) = 1 ORDER BY [name]";
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
            using (DbConnection connection = CreateConnection(CreateDatabaseSnapshot(profile, databaseName)))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                try
                {
                    List<DatabaseObjectMetadata> objects = await ReadSqlServerObjectsAsync(connection, true, cancellationToken).ConfigureAwait(false);
                    LogIfTableCommentsEmpty(objects);
                    return objects.AsReadOnly();
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception exception)
                {
                    LogComment(CommentDiagnostics.TableStage, exception.Message, 0, true, exception);
                    List<DatabaseObjectMetadata> objects = await ReadSqlServerObjectsAsync(connection, false, cancellationToken).ConfigureAwait(false);
                    return objects.AsReadOnly();
                }
            }
        }

        private async Task<List<DatabaseObjectMetadata>> ReadSqlServerObjectsAsync(
            DbConnection connection,
            bool includeComment,
            CancellationToken cancellationToken)
        {
            List<DatabaseObjectMetadata> objects = new List<DatabaseObjectMetadata>();
            using (DbCommand command = connection.CreateCommand())
            {
                command.CommandText = includeComment
                    ? "SELECT s.[name], o.[name], o.[type], CAST(ep.[value] AS nvarchar(4000)) FROM sys.objects o INNER JOIN sys.schemas s ON s.schema_id = o.schema_id LEFT JOIN sys.extended_properties ep ON ep.major_id = o.object_id AND ep.minor_id = 0 AND ep.[name] = N'MS_Description' WHERE o.[type] IN ('U', 'V') AND o.is_ms_shipped = 0 ORDER BY s.[name], o.[type], o.[name]"
                    : "SELECT s.[name], o.[name], o.[type] FROM sys.objects o INNER JOIN sys.schemas s ON s.schema_id = o.schema_id WHERE o.[type] IN ('U', 'V') AND o.is_ms_shipped = 0 ORDER BY s.[name], o.[type], o.[name]";
                using (DbDataReader reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        string comment = includeComment ? CommentDiagnostics.Normalize(ReadText(reader, 3)) : string.Empty;
                        objects.Add(new DatabaseObjectMetadata(
                            ReadText(reader, 0),
                            ReadText(reader, 1),
                            ReadText(reader, 2) == "V" ? DatabaseObjectKind.View : DatabaseObjectKind.Table,
                            comment));
                    }
                }
            }

            return objects;
        }

        /// <inheritdoc/>
        protected override Task ApplyPreviewRowLimitAsync(
            DbConnection connection,
            int rowLimit,
            CancellationToken cancellationToken)
        {
            return ExecuteSessionCommandAsync(
                connection,
                "SET ROWCOUNT " + rowLimit.ToString(System.Globalization.CultureInfo.InvariantCulture),
                cancellationToken);
        }
    }
}

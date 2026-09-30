using System.Collections.Generic;
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.IO;
using DB2Sheet.Contracts;
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

        /// <summary>创建 SQLite 提供程序。</summary>
        /// <param name="logger">注释缺口写入的文件日志。为空时不记录。</param>
        public SqliteProvider(ILogger logger = null) : base(logger)
        {
        }

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
        public override Task<IReadOnlyList<DatabaseMetadata>> GetDatabasesAsync(
            ConnectionProfileSnapshot profile,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string path = profile.GetValue(DatabaseParameterKeys.FilePath);
            IReadOnlyList<DatabaseMetadata> databases = new List<DatabaseMetadata>
            {
                new DatabaseMetadata(Path.GetFileNameWithoutExtension(path))
            }.AsReadOnly();
            return Task.FromResult(databases);
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
                    command.CommandText = "SELECT name, type FROM sqlite_master WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite_%' ORDER BY type, name";
                    using (DbDataReader reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                    {
                        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                        {
                            objects.Add(new DatabaseObjectMetadata(
                                "main",
                                reader.GetString(0),
                                reader.GetString(1) == "view" ? DatabaseObjectKind.View : DatabaseObjectKind.Table));
                        }
                    }
                }
            }
            IReadOnlyList<DatabaseObjectMetadata> listed = objects.AsReadOnly();
            LogIfTableCommentsEmpty(listed, "SQLite 不提供表注释。");
            return listed;
        }

        /// <inheritdoc/>
        public override ConnectionProfileSnapshot CreateDatabaseSnapshot(
            ConnectionProfileSnapshot profile,
            string databaseName)
        {
            if (profile == null) throw new System.ArgumentNullException(nameof(profile));
            Dictionary<string, string> parameters = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, string> parameter in profile.Parameters)
                parameters[parameter.Key] = parameter.Value;
            return new ConnectionProfileSnapshot(profile.Id, profile.Name, profile.ProviderId, parameters);
        }

        /// <inheritdoc/>
        public override Task ConfigureReadOnlySessionAsync(DbConnection connection, CancellationToken cancellationToken)
        {
            return ExecuteSessionCommandAsync(connection, "PRAGMA query_only = ON", cancellationToken);
        }
    }
}

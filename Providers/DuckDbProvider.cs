using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Common;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DuckDB.NET.Data;

namespace DB2Sheet.Providers
{
    /// <summary>使用 DuckDB.NET 以只读方式查询本地 DuckDB 数据库文件。</summary>
    /// <remarks>
    /// 只读在打开连接时通过 <c>access_mode=READ_ONLY</c> 生效，不会新建缺失文件。
    /// 本机库随插件输出放在 <c>DB2Sheet.dll</c> 旁边，首次建连前按该路径加载。DuckDB 没有 32 位 Windows 库。
    /// 不关闭外部文件访问，因此 <c>SELECT ... FROM read_parquet</c> 等只读表函数可以读本地文件。
    /// </remarks>
    public sealed class DuckDbProvider : DatabaseProviderBase
    {
        private const string AccessModeKey = "access_mode";
        private const string ReadOnlyMode = "READ_ONLY";
        private static readonly object NativeGate = new object();
        private static bool _nativeLoaded;

        private static readonly IReadOnlyList<ParameterDefinition> Parameters = new List<ParameterDefinition>
        {
            new ParameterDefinition(DatabaseParameterKeys.FilePath, "数据库文件", ParameterValueType.FilePath, true)
        }.AsReadOnly();

        /// <summary>创建 DuckDB 提供程序。</summary>
        /// <param name="logger">注释缺口写入的文件日志。为空时不记录。</param>
        public DuckDbProvider(ILogger logger = null) : base(logger)
        {
        }

        /// <inheritdoc/>
        public override string ProviderId => "duckdb";
        /// <inheritdoc/>
        public override string DisplayName => "DuckDB";
        /// <inheritdoc/>
        public override IReadOnlyList<ParameterDefinition> ConnectionParameters => Parameters;
        /// <inheritdoc/>
        public override ProviderCapabilities Capabilities => base.Capabilities | ProviderCapabilities.ReadOnlySession;

        /// <inheritdoc/>
        public override DbConnection CreateConnection(ConnectionProfileSnapshot profile)
        {
            EnsureNativeLibrary();
            DuckDBConnectionStringBuilder builder = new DuckDBConnectionStringBuilder
            {
                DataSource = profile.GetValue(DatabaseParameterKeys.FilePath)
            };
            builder[AccessModeKey] = ReadOnlyMode;
            return new DuckDBConnection(builder.ConnectionString);
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
            using (DbConnection connection = CreateConnection(profile))
            {
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                try
                {
                    List<DatabaseObjectMetadata> objects = await ReadObjectsAsync(connection, true, cancellationToken).ConfigureAwait(false);
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
                    List<DatabaseObjectMetadata> objects = await ReadObjectsAsync(connection, false, cancellationToken).ConfigureAwait(false);
                    return objects.AsReadOnly();
                }
            }
        }

        /// <inheritdoc/>
        public override ConnectionProfileSnapshot CreateDatabaseSnapshot(
            ConnectionProfileSnapshot profile,
            string databaseName)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            Dictionary<string, string> parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, string> parameter in profile.Parameters)
                parameters[parameter.Key] = parameter.Value;
            return new ConnectionProfileSnapshot(profile.Id, profile.Name, profile.ProviderId, parameters);
        }

        /// <summary>按插件原始目录加载 64 位 duckdb.dll。同一进程只加载一次。</summary>
        /// <remarks>
        /// Excel 会把插件复制到临时目录再加载，<c>Assembly.Location</c> 指向那份副本，旁边没有本机库。
        /// <c>Assembly.CodeBase</c> 仍是编译输出或安装目录。必须在任何 DuckDB.NET 类型首次使用之前调用。
        /// </remarks>
        /// <exception cref="InvalidOperationException">当前进程不是 64 位。</exception>
        /// <exception cref="FileNotFoundException">原始目录和加载目录里都没有 duckdb.dll。</exception>
        /// <exception cref="Win32Exception">Windows 拒绝加载该本机库。</exception>
        private static void EnsureNativeLibrary()
        {
            lock (NativeGate)
            {
                if (_nativeLoaded) return;
                if (IntPtr.Size != 8)
                    throw new InvalidOperationException("DuckDB 只提供 64 位本机库，当前 Excel 不是 64 位进程。");

                string[] directories = CandidateDirectories();
                string path = null;
                for (int index = 0; index < directories.Length; index++)
                {
                    if (string.IsNullOrEmpty(directories[index])) continue;
                    string candidate = Path.Combine(directories[index], "duckdb.dll");
                    if (!File.Exists(candidate)) continue;
                    path = candidate;
                    break;
                }

                if (path == null)
                    throw new FileNotFoundException("找不到 DuckDB 本机库。已查找：" + string.Join("；", directories));
                if (LoadLibrary(path) == IntPtr.Zero)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "无法加载 DuckDB 本机库：" + path);
                _nativeLoaded = true;
            }
        }

        /// <summary>先取插件原始目录，再取当前加载目录。两者相同或为空时不重复。</summary>
        /// <returns>用于查找 duckdb.dll 的目录。</returns>
        private static string[] CandidateDirectories()
        {
            string codeBaseDirectory = DirectoryFromCodeBase();
            string locationDirectory = Path.GetDirectoryName(typeof(DuckDbProvider).Assembly.Location);
            if (string.IsNullOrEmpty(codeBaseDirectory))
                return new[] { locationDirectory ?? string.Empty };
            if (string.IsNullOrEmpty(locationDirectory) ||
                string.Equals(codeBaseDirectory, locationDirectory, StringComparison.OrdinalIgnoreCase))
                return new[] { codeBaseDirectory };
            return new[] { codeBaseDirectory, locationDirectory };
        }

        /// <summary>从程序集 CodeBase 解析插件原来所在的目录。</summary>
        /// <returns>本地目录；CodeBase 缺失或不是文件路径时返回 null。</returns>
        private static string DirectoryFromCodeBase()
        {
            string codeBase = typeof(DuckDbProvider).Assembly.CodeBase;
            if (string.IsNullOrEmpty(codeBase)) return null;
            Uri uri;
            if (!Uri.TryCreate(codeBase, UriKind.Absolute, out uri) || !uri.IsFile) return null;
            return Path.GetDirectoryName(uri.LocalPath);
        }

        private async Task<List<DatabaseObjectMetadata>> ReadObjectsAsync(
            DbConnection connection,
            bool includeComment,
            CancellationToken cancellationToken)
        {
            List<DatabaseObjectMetadata> objects = new List<DatabaseObjectMetadata>();
            using (DbCommand command = connection.CreateCommand())
            {
                command.CommandText = includeComment
                    ? "SELECT schema_name, table_name, 'BASE TABLE', comment FROM duckdb_tables() WHERE database_name = current_database() AND NOT internal AND schema_name NOT IN ('information_schema', 'pg_catalog') UNION ALL SELECT schema_name, view_name, 'VIEW', comment FROM duckdb_views() WHERE database_name = current_database() AND NOT internal AND schema_name NOT IN ('information_schema', 'pg_catalog')"
                    : "SELECT table_schema, table_name, table_type FROM information_schema.tables WHERE table_catalog = current_database() AND table_schema NOT IN ('information_schema', 'pg_catalog') ORDER BY table_schema, table_type, table_name";
                using (DbDataReader reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false))
                {
                    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
                    {
                        string comment = includeComment ? CommentDiagnostics.Normalize(ReadText(reader, 3)) : string.Empty;
                        string kind = ReadText(reader, 2);
                        objects.Add(new DatabaseObjectMetadata(
                            ReadText(reader, 0),
                            ReadText(reader, 1),
                            kind.Equals("VIEW", StringComparison.OrdinalIgnoreCase) ? DatabaseObjectKind.View : DatabaseObjectKind.Table,
                            comment));
                    }
                }
            }

            return objects;
        }

        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);
    }
}

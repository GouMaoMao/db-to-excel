using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;

namespace DB2Sheet.Providers
{
    /// <summary>通过用户本机的 Java 和连接方案中的 JDBC 驱动执行只读查询。</summary>
    /// <remarks>
    /// 插件不附带 JDK，也不附带厂商驱动。Java 在「JDBC 环境」中配置；驱动 jar 与驱动类写在连接方案里。
    /// 查询在独立 Java 进程中执行。本提供程序实例可被多个操作复用，进程由 <see cref="JdbcBridgeHost"/> 管理。
    /// </remarks>
    public sealed class JdbcProvider : IDataSourceProvider, IDatabaseMetadataProvider
    {
        private readonly JdbcBridgeHost _host;
        private readonly IJdbcEnvironmentStore _environment;
        private readonly ILogger _logger;
        private readonly ISqlReadOnlyValidator _readOnlyValidator;

        /// <summary>创建 JDBC 提供程序。构造函数对程序集内部可见，因为宿主类型不对外公开。</summary>
        /// <param name="host">转接进程宿主。调用方拥有其生命周期。</param>
        /// <param name="environment">JDBC 环境存储。</param>
        /// <param name="logger">注释缺口写入的文件日志。为空时不记录。字段注释由结果流在读到列信息时记录。</param>
        /// <param name="readOnlyValidator">只读 SQL 校验器；为空时使用默认实现。</param>
        internal JdbcProvider(
            JdbcBridgeHost host,
            IJdbcEnvironmentStore environment,
            ILogger logger = null,
            ISqlReadOnlyValidator readOnlyValidator = null)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _environment = environment ?? throw new ArgumentNullException(nameof(environment));
            _logger = logger;
            _readOnlyValidator = readOnlyValidator ?? new SqlReadOnlyValidator();
        }

        /// <inheritdoc/>
        public string ProviderId => JdbcConnectionComposer.ProviderId;

        /// <inheritdoc/>
        public string DisplayName => "JDBC";

        /// <inheritdoc/>
        public ProviderCapabilities Capabilities =>
            ProviderCapabilities.Query |
            ProviderCapabilities.Preview |
            ProviderCapabilities.Streaming |
            ProviderCapabilities.Cancellation;

        /// <inheritdoc/>
        public IReadOnlyList<ParameterDefinition> ConnectionParameters => JdbcConnectionComposer.Parameters;

        /// <inheritdoc/>
        public IReadOnlyList<string> ValidateConnection(ConnectionProfileSnapshot profile)
        {
            return JdbcConnectionComposer.Validate(profile);
        }

        /// <inheritdoc/>
        public IReadOnlyList<string> ValidateRequest(DataSourceRequest request)
        {
            List<string> errors = new List<string>();
            if (request == null)
            {
                errors.Add("查询请求不能为空。");
                return errors;
            }

            errors.AddRange(_readOnlyValidator.Validate(request.QueryText, ProviderId));
            return errors;
        }

        /// <inheritdoc/>
        public async Task<ConnectionTestResult> TestConnectionAsync(
            ConnectionProfileSnapshot profile,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            IReadOnlyList<string> errors = ValidateConnection(profile);
            if (errors.Count > 0)
            {
                return ConnectionTestResult.Failure(string.Join(Environment.NewLine, errors), TimeSpan.Zero);
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            try
            {
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Connecting,
                    ConnectionName = profile.Name,
                    Message = "正在测试 JDBC 连接…",
                    IsIndeterminate = true
                });
                JdbcConnectionTarget target = JdbcConnectionComposer.Compose(profile);
                await _host.RequestAsync(CurrentLaunch(), CreateRequest("test", target, null, 0), cancellationToken).ConfigureAwait(false);
                return ConnectionTestResult.Success(stopwatch.Elapsed);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                return ConnectionTestResult.Failure(exception.Message, stopwatch.Elapsed);
            }
        }

        /// <inheritdoc/>
        public async Task<IDataResultStream> ExecuteAsync(
            DataSourceRequest request,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            IReadOnlyList<string> errors = ValidateRequest(request);
            if (errors.Count > 0)
            {
                throw new InvalidOperationException(string.Join(Environment.NewLine, errors));
            }

            errors = ValidateConnection(request.Connection);
            if (errors.Count > 0)
            {
                throw new InvalidOperationException(string.Join(Environment.NewLine, errors));
            }

            progress?.Report(new OperationProgress
            {
                Stage = OperationStage.Connecting,
                ConnectionName = request.Connection.Name,
                Message = "正在连接 JDBC 数据源…",
                IsIndeterminate = true
            });
            JdbcConnectionTarget target = JdbcConnectionComposer.Compose(request.Connection);
            return await _host.QueryAsync(
                CurrentLaunch(),
                target,
                request.QueryText,
                request.TimeoutSeconds,
                request.RowLimit,
                request.Connection.Name,
                progress,
                cancellationToken).ConfigureAwait(false);
        }

        /// <inheritdoc/>
        public async Task<IReadOnlyList<DatabaseMetadata>> GetDatabasesAsync(
            ConnectionProfileSnapshot profile,
            CancellationToken cancellationToken)
        {
            IReadOnlyList<string> errors = ValidateConnection(profile);
            if (errors.Count > 0)
            {
                throw new InvalidOperationException(string.Join(Environment.NewLine, errors));
            }

            JdbcConnectionTarget target = JdbcConnectionComposer.Compose(profile);
            JdbcBridgeReply reply = await _host.RequestAsync(
                CurrentLaunch(),
                CreateRequest("catalogs", target, null, 0),
                cancellationToken).ConfigureAwait(false);
            return (reply.Names ?? new List<string>())
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(name => new DatabaseMetadata(name))
                .ToList()
                .AsReadOnly();
        }

        /// <inheritdoc/>
        public async Task<IReadOnlyList<DatabaseObjectMetadata>> GetDatabaseObjectsAsync(
            ConnectionProfileSnapshot profile,
            string databaseName,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(databaseName)) throw new ArgumentException("数据库名称不能为空。", nameof(databaseName));
            IReadOnlyList<string> errors = ValidateConnection(profile);
            if (errors.Count > 0)
            {
                throw new InvalidOperationException(string.Join(Environment.NewLine, errors));
            }

            JdbcConnectionTarget target = JdbcConnectionComposer.Compose(profile);
            JdbcBridgeRequest request = CreateRequest("tables", target, null, 0);
            request.Catalog = databaseName;
            JdbcBridgeReply reply = await _host.RequestAsync(CurrentLaunch(), request, cancellationToken).ConfigureAwait(false);
            List<DatabaseObjectMetadata> objects = new List<DatabaseObjectMetadata>();
            foreach (JdbcObjectInfo item in reply.Objects ?? new List<JdbcObjectInfo>())
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Name)) continue;
                DatabaseObjectKind kind = string.Equals(item.Kind, "view", StringComparison.OrdinalIgnoreCase)
                    ? DatabaseObjectKind.View
                    : DatabaseObjectKind.Table;
                objects.Add(new DatabaseObjectMetadata(item.Schema, item.Name, kind, CommentDiagnostics.Normalize(item.Comment)));
            }

            CommentDiagnostics.Write(
                _logger,
                ProviderId,
                CommentDiagnostics.TableStage,
                reply.CommentNote,
                objects.Count,
                reply.CommentFailed);
            return objects.AsReadOnly();
        }

        /// <inheritdoc/>
        public ConnectionProfileSnapshot CreateDatabaseSnapshot(ConnectionProfileSnapshot profile, string databaseName)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (string.IsNullOrWhiteSpace(databaseName)) throw new ArgumentException("数据库名称不能为空。", nameof(databaseName));
            Dictionary<string, string> parameters = profile.Parameters.ToDictionary(
                item => item.Key,
                item => item.Value,
                StringComparer.OrdinalIgnoreCase);
            parameters[JdbcParameterKeys.Catalog] = databaseName;
            return new ConnectionProfileSnapshot(profile.Id, profile.Name, profile.ProviderId, parameters);
        }

        private JdbcBridgeLaunch CurrentLaunch()
        {
            JdbcEnvironmentSettings settings = _environment.Load();
            return new JdbcBridgeLaunch(settings.JavaExecutable);
        }

        private static JdbcBridgeRequest CreateRequest(string operation, JdbcConnectionTarget target, string sql, int timeoutSeconds)
        {
            return new JdbcBridgeRequest
            {
                Op = operation,
                Url = target.Url,
                User = target.UserName,
                Password = target.Password,
                Driver = target.DriverClass,
                Jars = target.DriverJars.ToList(),
                Sql = sql ?? string.Empty,
                Catalog = target.Catalog,
                TimeoutSeconds = timeoutSeconds
            };
        }
    }
}

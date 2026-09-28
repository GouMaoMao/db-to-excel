using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;

namespace DB2Sheet.Providers
{
    /// <summary>
    /// 为关系型数据库提供程序实现公共校验、连接测试和流式查询流程。
    /// </summary>
    /// <remarks>
    /// 派生类只需定义连接参数、创建连接，并按需配置数据库级只读会话。
    /// 查询成功后，连接和命令的所有权会转交给返回的 <see cref="IDataResultStream"/>。
    /// </remarks>
    public abstract class DatabaseProviderBase : IDatabaseQueryProvider
    {
        private readonly ISqlReadOnlyValidator _readOnlyValidator;

        /// <summary>初始化数据库提供程序并注入只读 SQL 校验器。</summary>
        /// <param name="readOnlyValidator">校验器；为空时创建默认实现。</param>
        protected DatabaseProviderBase(ISqlReadOnlyValidator readOnlyValidator = null)
        {
            _readOnlyValidator = readOnlyValidator ?? new SqlReadOnlyValidator();
        }

        /// <inheritdoc/>
        public abstract string ProviderId { get; }
        /// <inheritdoc/>
        public abstract string DisplayName { get; }
        /// <inheritdoc/>
        public virtual ProviderCapabilities Capabilities =>
            ProviderCapabilities.Query |
            ProviderCapabilities.Preview |
            ProviderCapabilities.Streaming |
            ProviderCapabilities.Cancellation;
        /// <inheritdoc/>
        public abstract IReadOnlyList<ParameterDefinition> ConnectionParameters { get; }

        /// <inheritdoc/>
        public virtual IReadOnlyList<string> ValidateConnection(ConnectionProfileSnapshot profile)
        {
            List<string> errors = new List<string>();
            if (profile == null)
            {
                errors.Add("连接方案不能为空。");
                return errors;
            }

            if (!string.Equals(profile.ProviderId, ProviderId, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("连接方案与数据源类型不匹配。");
            }

            foreach (ParameterDefinition definition in ConnectionParameters.Where(item => item.IsRequired))
            {
                if (string.IsNullOrWhiteSpace(profile.GetValue(definition.Key)))
                {
                    errors.Add(definition.DisplayName + "不能为空。");
                }
            }

            return errors;
        }

        /// <inheritdoc/>
        public virtual IReadOnlyList<string> ValidateRequest(DataSourceRequest request)
        {
            List<string> errors = new List<string>();
            if (request == null) errors.Add("查询请求不能为空。");
            else
            {
                errors.AddRange(_readOnlyValidator.Validate(request.QueryText, ProviderId));
            }
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
                    Message = "正在测试连接…",
                    IsIndeterminate = true
                });

                using (DbConnection connection = CreateConnection(profile))
                {
                    await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                }

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
        public abstract DbConnection CreateConnection(ConnectionProfileSnapshot profile);

        /// <inheritdoc/>
        public virtual DbCommand CreateCommand(DbConnection connection, string queryText, int timeoutSeconds)
        {
            DbCommand command = connection.CreateCommand();
            command.CommandText = queryText;
            command.CommandTimeout = timeoutSeconds;
            return command;
        }

        /// <inheritdoc/>
        public virtual Task ConfigureReadOnlySessionAsync(DbConnection connection, CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }

        /// <summary>在已打开连接上执行一条会话级非查询命令。</summary>
        /// <param name="connection">已打开的连接。</param>
        /// <param name="commandText">会话配置命令。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>表示命令完成的任务。</returns>
        protected static async Task ExecuteSessionCommandAsync(
            DbConnection connection,
            string commandText,
            CancellationToken cancellationToken)
        {
            using (DbCommand command = connection.CreateCommand())
            {
                command.CommandText = commandText;
                await command.ExecuteNonQueryAsync(cancellationToken).ConfigureAwait(false);
            }
        }

        /// <inheritdoc/>
        public virtual async Task<IDataResultStream> ExecuteAsync(
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

            DbConnection connection = null;
            DbCommand command = null;
            try
            {
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Connecting,
                    ConnectionName = request.Connection.Name,
                    Message = "正在连接数据源…",
                    IsIndeterminate = true
                });
                connection = CreateConnection(request.Connection);
                await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
                await ConfigureReadOnlySessionAsync(connection, cancellationToken).ConfigureAwait(false);

                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Executing,
                    ConnectionName = request.Connection.Name,
                    Message = "正在执行只读查询…",
                    IsIndeterminate = true
                });
                command = CreateCommand(connection, request.QueryText, request.TimeoutSeconds);
                DbDataReader reader = await command.ExecuteReaderAsync(
                    CommandBehavior.SequentialAccess,
                    cancellationToken).ConfigureAwait(false);
                return new DatabaseResultStream(
                    connection,
                    command,
                    reader,
                    request.RowLimit,
                    progress,
                    cancellationToken);
            }
            catch
            {
                command?.Dispose();
                connection?.Dispose();
                throw;
            }
        }

        /// <summary>从连接参数读取布尔值，格式无效时使用默认值。</summary>
        protected static bool GetBoolean(ConnectionProfileSnapshot profile, string key, bool defaultValue = false)
        {
            return bool.TryParse(profile.GetValue(key), out bool value) ? value : defaultValue;
        }

        /// <summary>从连接参数读取端口，格式无效时使用默认端口。</summary>
        protected static int GetPort(ConnectionProfileSnapshot profile, int defaultPort)
        {
            return int.TryParse(profile.GetValue(DatabaseParameterKeys.Port), out int value) ? value : defaultPort;
        }
    }
}

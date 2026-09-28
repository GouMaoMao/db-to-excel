using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>按连接中的提供程序标识解析数据源，并统一记录查询和连接测试日志。</summary>
    /// <remarks>日志只记录查询哈希，不记录 SQL 正文，以减少敏感信息泄露。</remarks>
    public sealed class DataSourceExecutionService : IDataSourceExecutionService
    {
        private readonly IProviderRegistry _providers;
        private readonly ILogger _logger;

        /// <summary>创建数据源执行服务。</summary>
        /// <param name="providers">提供程序注册表。</param>
        /// <param name="logger">结构化日志记录器。</param>
        public DataSourceExecutionService(IProviderRegistry providers, ILogger logger)
        {
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<IDataResultStream> ExecuteAsync(
            DataSourceRequest request,
            string operationId,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            IDataSourceProvider provider = Resolve(request.Connection.ProviderId);
            Stopwatch stopwatch = Stopwatch.StartNew();
            try
            {
                IDataResultStream stream = await provider.ExecuteAsync(request, progress, cancellationToken).ConfigureAwait(false);
                _logger.Write(LogSeverity.Information, "数据源查询已启动。", operationId, properties: Properties(request, stopwatch.Elapsed));
                return stream;
            }
            catch (Exception exception)
            {
                _logger.Write(LogSeverity.Error, "数据源查询启动失败。", operationId, exception, Properties(request, stopwatch.Elapsed));
                throw;
            }
        }

        /// <inheritdoc/>
        public async Task<ConnectionTestResult> TestConnectionAsync(
            ConnectionProfileSnapshot profile,
            string operationId,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            IDataSourceProvider provider = Resolve(profile.ProviderId);
            ConnectionTestResult result = await provider.TestConnectionAsync(profile, progress, cancellationToken).ConfigureAwait(false);
            _logger.Write(
                result.Succeeded ? LogSeverity.Information : LogSeverity.Warning,
                result.Succeeded ? "连接测试成功。" : "连接测试失败。",
                operationId,
                properties: new Dictionary<string, string>
                {
                    ["provider"] = provider.ProviderId,
                    ["profile"] = profile.Name,
                    ["elapsedMs"] = ((long)result.Elapsed.TotalMilliseconds).ToString()
                });
            return result;
        }

        private IDataSourceProvider Resolve(string providerId)
        {
            return _providers.GetById(providerId) ??
                throw new InvalidOperationException("未注册的数据源 Provider：" + providerId);
        }

        private static IReadOnlyDictionary<string, string> Properties(DataSourceRequest request, TimeSpan elapsed)
        {
            return new Dictionary<string, string>
            {
                ["provider"] = request.Connection.ProviderId,
                ["profile"] = request.Connection.Name,
                ["purpose"] = request.Purpose.ToString(),
                ["queryHash"] = ComputeStableHash(request.QueryText),
                ["elapsedMs"] = ((long)elapsed.TotalMilliseconds).ToString()
            };
        }

        private static string ComputeStableHash(string value)
        {
            unchecked
            {
                uint hash = 2166136261;
                foreach (char character in value ?? string.Empty)
                {
                    hash ^= character;
                    hash *= 16777619;
                }
                return hash.ToString("X8");
            }
        }
    }
}

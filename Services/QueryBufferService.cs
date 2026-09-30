using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>读取数据源流的所有结果块并组成可供 Excel 写入的内存结果。</summary>
    /// <remarks>
    /// 结果量越大占用内存越多；应通过请求行数上限控制峰值。结果流始终在方法结束时释放。
    /// 读取结束、取消或失败时写入文件汇总。查询启动失败仍由数据源执行服务记录，此处不重复记。
    /// </remarks>
    public sealed class QueryBufferService : IQueryBufferService
    {
        private readonly IDataSourceExecutionService _executionService;
        private readonly ILogger _logger;

        /// <summary>创建查询缓冲服务。</summary>
        /// <param name="executionService">统一数据源执行入口。</param>
        /// <param name="logger">读取阶段的文件日志。不按进度日志级别过滤。</param>
        public QueryBufferService(IDataSourceExecutionService executionService, ILogger logger)
        {
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<BufferedQueryResult> ExecuteAsync(
            DataSourceRequest request,
            string operationId,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            List<object[,]> blocks = new List<object[,]>();
            using (IDataResultStream stream = await _executionService.ExecuteAsync(
                request,
                operationId,
                progress,
                cancellationToken).ConfigureAwait(false))
            {
                Stopwatch readWatch = Stopwatch.StartNew();
                try
                {
                    while (!stream.IsCompleted)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        ResultBlock block = await stream.ReadBlockAsync(request.BlockSize, cancellationToken).ConfigureAwait(false);
                        if (block.RowCount > 0) blocks.Add(block.Values);
                    }

                    _logger.Write(
                        LogSeverity.Information,
                        "结果读取完成。",
                        operationId,
                        properties: ReadProperties(request, stream, blocks.Count, readWatch.Elapsed, includeRate: true));
                    return new BufferedQueryResult(stream.Columns, blocks.AsReadOnly(), stream.RowsRead, stream.IsTruncated);
                }
                catch (OperationCanceledException)
                {
                    _logger.Write(
                        LogSeverity.Information,
                        "结果读取已取消。",
                        operationId,
                        properties: ReadOutcome(stream.RowsRead, blocks.Count, readWatch.Elapsed, includeBlockCount: false));
                    throw;
                }
                catch (Exception exception)
                {
                    _logger.Write(
                        LogSeverity.Error,
                        "结果读取失败。",
                        operationId,
                        exception,
                        ReadOutcome(stream.RowsRead, blocks.Count, readWatch.Elapsed, includeBlockCount: true));
                    throw;
                }
            }
        }

        /// <summary>组装读取完成日志的属性。耗时不大于 0 时不写行每秒。</summary>
        private static Dictionary<string, string> ReadProperties(
            DataSourceRequest request,
            IDataResultStream stream,
            int blockCount,
            TimeSpan elapsed,
            bool includeRate)
        {
            Dictionary<string, string> properties = ReadOutcome(stream.RowsRead, blockCount, elapsed, includeBlockCount: true);
            properties["columnCount"] = (stream.Columns == null ? 0 : stream.Columns.Count).ToString();
            properties["blockSize"] = request.BlockSize.ToString();
            properties["rowLimit"] = request.RowLimit.ToString();
            properties["truncated"] = stream.IsTruncated ? "true" : "false";
            if (includeRate) AddRowsPerSecond(properties, stream.RowsRead, elapsed);
            return properties;
        }

        /// <summary>组装读取取消或失败时已经知道的行数和耗时。</summary>
        private static Dictionary<string, string> ReadOutcome(long rowsRead, int blockCount, TimeSpan elapsed, bool includeBlockCount)
        {
            Dictionary<string, string> properties = new Dictionary<string, string>
            {
                ["rowsRead"] = rowsRead.ToString(),
                ["elapsedMs"] = ((long)elapsed.TotalMilliseconds).ToString()
            };
            if (includeBlockCount) properties["blockCount"] = blockCount.ToString();
            return properties;
        }

        /// <summary>在耗时大于 0 时写入每秒行数。</summary>
        private static void AddRowsPerSecond(Dictionary<string, string> properties, long rows, TimeSpan elapsed)
        {
            long elapsedMs = (long)elapsed.TotalMilliseconds;
            if (elapsedMs <= 0) return;
            properties["rowsPerSecond"] = (rows * 1000L / elapsedMs).ToString();
        }
    }
}

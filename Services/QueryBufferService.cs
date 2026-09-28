using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>读取数据源流的所有结果块并组成可供 Excel 写入的内存结果。</summary>
    /// <remarks>结果量越大占用内存越多；应通过请求行数上限控制峰值。结果流始终在方法结束时释放。</remarks>
    public sealed class QueryBufferService : IQueryBufferService
    {
        private readonly IDataSourceExecutionService _executionService;

        /// <summary>创建查询缓冲服务。</summary>
        /// <param name="executionService">统一数据源执行入口。</param>
        public QueryBufferService(IDataSourceExecutionService executionService)
        {
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
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
                while (!stream.IsCompleted)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ResultBlock block = await stream.ReadBlockAsync(request.BlockSize, cancellationToken).ConfigureAwait(false);
                    if (block.RowCount > 0) blocks.Add(block.Values);
                }

                return new BufferedQueryResult(stream.Columns, blocks.AsReadOnly(), stream.RowsRead, stream.IsTruncated);
            }
        }
    }
}

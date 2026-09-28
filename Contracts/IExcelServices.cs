using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.Contracts
{
    /// <summary>从当前 Excel 工作簿的 SQL Sheet 中读取批量刷新任务。</summary>
    public interface ISqlSheetTaskReader
    {
        /// <summary>扫描工作簿并生成可执行的刷新任务。</summary>
        /// <param name="workbook">要扫描的 Excel 工作簿 COM 对象。</param>
        /// <returns>按源列顺序生成的任务列表。</returns>
        /// <remarks>调用方必须保证在 Excel UI 线程访问传入的 COM 对象。</remarks>
        IReadOnlyList<RefreshTaskDefinition> ReadTasks(ExcelInterop.Workbook workbook);
    }

    /// <summary>将已缓冲的查询结果写入 Excel 工作表。</summary>
    public interface IExcelResultWriter
    {
        /// <summary>创建或清空目标工作表，并批量写入列名和数据。</summary>
        /// <param name="workbook">目标工作簿。</param>
        /// <param name="targetSheetName">目标工作表名称。</param>
        /// <param name="result">已在内存中缓冲的查询结果。</param>
        /// <param name="progress">可选的进度接收器。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>实际写入的数据行数，不包含标题行。</returns>
        /// <remarks>该方法会修改工作簿内容，应在 Excel UI 线程调用。</remarks>
        long WriteResult(
            ExcelInterop.Workbook workbook,
            string targetSheetName,
            BufferedQueryResult result,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken);
    }

    /// <summary>执行流式查询并将全部结果转换为内存块。</summary>
    public interface IQueryBufferService
    {
        /// <summary>执行请求并依次读取结果块。</summary>
        /// <param name="request">数据源查询请求。</param>
        /// <param name="operationId">日志和进度关联标识。</param>
        /// <param name="progress">可选的进度接收器。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>包含列、数据块、总行数和截断状态的缓冲结果。</returns>
        Task<BufferedQueryResult> ExecuteAsync(
            DataSourceRequest request,
            string operationId,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken);
    }

    /// <summary>编排多个 SQL Sheet 任务的查询、缓冲和 Excel 写入。</summary>
    public interface IBatchRefreshService
    {
        /// <summary>按指定模式刷新一组工作表任务。</summary>
        /// <param name="workbook">要写入的 Excel 工作簿。</param>
        /// <param name="connection">所有任务共享的连接快照。</param>
        /// <param name="tasks">待执行任务列表。</param>
        /// <param name="mode">串行或并行查询模式。</param>
        /// <param name="maximumParallelism">并行模式下允许同时查询的最大任务数。</param>
        /// <param name="rowLimit">每个任务允许读取的最大行数。</param>
        /// <param name="blockSize">每次从数据源读取的行数。</param>
        /// <param name="timeoutSeconds">单个数据库命令超时秒数。</param>
        /// <param name="operationId">日志和进度关联标识。</param>
        /// <param name="progress">可选的总体进度接收器。</param>
        /// <param name="cancellationToken">取消整个批次的令牌。</param>
        /// <returns>每个任务的成功、跳过、写入行数和消息汇总。</returns>
        /// <remarks>数据库查询可并行执行，但 Excel COM 写入仍由服务串行完成。</remarks>
        Task<BatchRefreshResult> RefreshAsync(
            ExcelInterop.Workbook workbook,
            ConnectionProfileSnapshot connection,
            IReadOnlyList<RefreshTaskDefinition> tasks,
            BatchExecutionMode mode,
            int maximumParallelism,
            int rowLimit,
            int blockSize,
            int timeoutSeconds,
            string operationId,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken);
    }
}

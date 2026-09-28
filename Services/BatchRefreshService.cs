using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.Services
{
    /// <summary>编排 SQL Sheet 批量任务的查询缓冲、错误隔离和 Excel 写入。</summary>
    /// <remarks>
    /// 并行模式只并发执行数据库查询；所有 Excel COM 写入都会回到启动调用的 UI 同步上下文并按任务顺序执行。
    /// 目标 Sheet 名重复的后续任务会被跳过，单个任务失败不会终止其他任务，用户取消会终止整个批次。
    /// </remarks>
    public sealed class BatchRefreshService : IBatchRefreshService
    {
        private readonly IQueryBufferService _queryBuffer;
        private readonly IExcelResultWriter _writer;
        private readonly ILogger _logger;

        /// <summary>创建批量刷新服务。</summary>
        /// <param name="queryBuffer">负责执行并缓冲查询结果的服务。</param>
        /// <param name="writer">负责写入 Excel 的服务。</param>
        /// <param name="logger">错误日志记录器。</param>
        public BatchRefreshService(IQueryBufferService queryBuffer, IExcelResultWriter writer, ILogger logger)
        {
            _queryBuffer = queryBuffer ?? throw new ArgumentNullException(nameof(queryBuffer));
            _writer = writer ?? throw new ArgumentNullException(nameof(writer));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<BatchRefreshResult> RefreshAsync(
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
            CancellationToken cancellationToken)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (connection == null) throw new ArgumentNullException(nameof(connection));
            if (tasks == null) throw new ArgumentNullException(nameof(tasks));

            SynchronizationContext excelContext = SynchronizationContext.Current ??
                throw new InvalidOperationException("批量刷新必须从 Excel UI 线程启动。");
            List<RefreshTaskDefinition> executable = RemoveDuplicateTargets(tasks, out List<RefreshTaskResult> skipped);
            Stopwatch stopwatch = Stopwatch.StartNew();
            List<RefreshTaskResult> results = new List<RefreshTaskResult>(skipped);

            if (mode == BatchExecutionMode.Parallel && executable.Count > 1)
            {
                IReadOnlyList<BufferedTask> buffered = await BufferParallelAsync(
                    connection, executable, Math.Max(1, maximumParallelism), rowLimit, blockSize,
                    timeoutSeconds, operationId, progress, skipped.Count, tasks.Count, stopwatch,
                    cancellationToken).ConfigureAwait(false);
                foreach (BufferedTask item in buffered.OrderBy(value => value.Index))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (item.Error != null)
                    {
                        results.Add(Failed(item.Task, item.Error));
                        continue;
                    }
                    results.Add(WriteOnContext(
                        excelContext, workbook, item.Task, item.Result, progress,
                        results, tasks.Count, stopwatch, cancellationToken));
                }
            }
            else
            {
                for (int index = 0; index < executable.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    RefreshTaskDefinition task = executable[index];
                    try
                    {
                        BufferedQueryResult buffered = await BufferTaskAsync(
                            connection, task, rowLimit, blockSize, timeoutSeconds, operationId,
                            progress, results.Count + 1, tasks.Count, stopwatch, cancellationToken).ConfigureAwait(false);
                        results.Add(WriteOnContext(
                            excelContext, workbook, task, buffered, progress,
                            results, tasks.Count, stopwatch, cancellationToken));
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        _logger.Write(LogSeverity.Error, "批量刷新任务失败。", operationId, exception);
                        results.Add(Failed(task, exception));
                    }
                }
            }

            progress?.Report(CreateProgress(
                OperationStage.Completed,
                null,
                "批量刷新完成。",
                results.Count,
                tasks.Count,
                results,
                stopwatch.Elapsed));
            return new BatchRefreshResult(OrderResults(tasks, results));
        }

        private async Task<IReadOnlyList<BufferedTask>> BufferParallelAsync(
            ConnectionProfileSnapshot connection,
            IReadOnlyList<RefreshTaskDefinition> tasks,
            int maximumParallelism,
            int rowLimit,
            int blockSize,
            int timeoutSeconds,
            string operationId,
            IProgress<OperationProgress> progress,
            int initialCompleted,
            int totalTasks,
            Stopwatch stopwatch,
            CancellationToken cancellationToken)
        {
            SemaphoreSlim semaphore = new SemaphoreSlim(maximumParallelism, maximumParallelism);
            try
            {
                int completed = initialCompleted;
                Task<BufferedTask>[] operations = tasks.Select((task, index) => Task.Run(async () =>
                {
                    await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
                    try
                    {
                        BufferedQueryResult result = await BufferTaskAsync(
                            connection, task, rowLimit, blockSize, timeoutSeconds, operationId,
                            progress, Volatile.Read(ref completed) + 1, totalTasks, stopwatch,
                            cancellationToken).ConfigureAwait(false);
                        Interlocked.Increment(ref completed);
                        return new BufferedTask(index, task, result, null);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        Interlocked.Increment(ref completed);
                        _logger.Write(LogSeverity.Error, "并行读取任务失败。", operationId, exception);
                        return new BufferedTask(index, task, null, exception);
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }, cancellationToken)).ToArray();
                return await Task.WhenAll(operations).ConfigureAwait(false);
            }
            finally
            {
                semaphore.Dispose();
            }
        }

        private Task<BufferedQueryResult> BufferTaskAsync(
            ConnectionProfileSnapshot connection,
            RefreshTaskDefinition task,
            int rowLimit,
            int blockSize,
            int timeoutSeconds,
            string operationId,
            IProgress<OperationProgress> progress,
            int currentTask,
            int totalTasks,
            Stopwatch stopwatch,
            CancellationToken cancellationToken)
        {
            DataSourceRequest request = new DataSourceRequest(
                connection,
                task.QueryText,
                ExecutionPurpose.Export,
                rowLimit,
                blockSize,
                timeoutSeconds);
            IProgress<OperationProgress> taskProgress = new ForwardingProgress(value =>
            {
                value.TaskId = task.Id;
                value.TargetSheetName = task.TargetSheetName;
                value.CurrentTask = currentTask;
                value.TotalTasks = totalTasks;
                value.Elapsed = stopwatch.Elapsed;
                progress?.Report(value);
            });
            return _queryBuffer.ExecuteAsync(request, operationId + "-" + task.Id, taskProgress, cancellationToken);
        }

        private RefreshTaskResult WriteOnContext(
            SynchronizationContext context,
            ExcelInterop.Workbook workbook,
            RefreshTaskDefinition task,
            BufferedQueryResult buffered,
            IProgress<OperationProgress> progress,
            IReadOnlyList<RefreshTaskResult> completed,
            int totalTasks,
            Stopwatch stopwatch,
            CancellationToken cancellationToken)
        {
            RefreshTaskResult result = null;
            Exception error = null;
            context.Send(state =>
            {
                try
                {
                    IProgress<OperationProgress> writeProgress = new ForwardingProgress(value =>
                    {
                        value.TaskId = task.Id;
                        value.TargetSheetName = task.TargetSheetName;
                        value.CurrentTask = completed.Count + 1;
                        value.TotalTasks = totalTasks;
                        value.SucceededTasks = completed.Count(item => item.Succeeded);
                        value.FailedTasks = completed.Count(item => !item.Succeeded && !item.Skipped);
                        value.SkippedTasks = completed.Count(item => item.Skipped);
                        value.Elapsed = stopwatch.Elapsed;
                        progress?.Report(value);
                    });
                    long rows = _writer.WriteResult(workbook, task.TargetSheetName, buffered, writeProgress, cancellationToken);
                    result = new RefreshTaskResult(
                        task, true, false, rows, buffered.IsTruncated,
                        buffered.IsTruncated ? "刷新成功，结果已截断。" : "刷新成功。");
                }
                catch (Exception exception)
                {
                    error = exception;
                }
            }, null);
            if (error != null)
            {
                _logger.Write(LogSeverity.Error, "写入 Excel 失败。", null, error);
                return Failed(task, error);
            }
            return result;
        }

        private static List<RefreshTaskDefinition> RemoveDuplicateTargets(
            IReadOnlyList<RefreshTaskDefinition> tasks,
            out List<RefreshTaskResult> skipped)
        {
            HashSet<string> targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<RefreshTaskDefinition> executable = new List<RefreshTaskDefinition>();
            skipped = new List<RefreshTaskResult>();
            foreach (RefreshTaskDefinition task in tasks)
            {
                if (targets.Add(task.TargetSheetName)) executable.Add(task);
                else skipped.Add(new RefreshTaskResult(task, false, true, 0, false, "目标 Sheet 重复，已跳过。"));
            }
            return executable;
        }

        private static RefreshTaskResult Failed(RefreshTaskDefinition task, Exception exception)
        {
            return new RefreshTaskResult(task, false, false, 0, false, exception.Message);
        }

        private static OperationProgress CreateProgress(
            OperationStage stage,
            RefreshTaskDefinition task,
            string message,
            int currentTask,
            int totalTasks,
            IReadOnlyList<RefreshTaskResult> results,
            TimeSpan elapsed)
        {
            return new OperationProgress
            {
                TaskId = task?.Id,
                TargetSheetName = task?.TargetSheetName,
                Stage = stage,
                Message = message,
                CurrentTask = currentTask,
                TotalTasks = totalTasks,
                SucceededTasks = results.Count(item => item.Succeeded),
                FailedTasks = results.Count(item => !item.Succeeded && !item.Skipped),
                SkippedTasks = results.Count(item => item.Skipped),
                RowsWritten = results.Sum(item => item.RowsWritten),
                Elapsed = elapsed,
                IsIndeterminate = stage != OperationStage.Completed
            };
        }

        private static IReadOnlyList<RefreshTaskResult> OrderResults(
            IReadOnlyList<RefreshTaskDefinition> tasks,
            IReadOnlyList<RefreshTaskResult> results)
        {
            Dictionary<string, RefreshTaskResult> byId = results.ToDictionary(item => item.Task.Id, StringComparer.OrdinalIgnoreCase);
            return tasks.Where(task => byId.ContainsKey(task.Id)).Select(task => byId[task.Id]).ToList().AsReadOnly();
        }

        /// <summary>保存并行查询阶段的任务顺序、结果或异常，供后续串行写入。</summary>
        private sealed class BufferedTask
        {
            public BufferedTask(int index, RefreshTaskDefinition task, BufferedQueryResult result, Exception error)
            {
                Index = index;
                Task = task;
                Result = result;
                Error = error;
            }

            public int Index { get; }
            public RefreshTaskDefinition Task { get; }
            public BufferedQueryResult Result { get; }
            public Exception Error { get; }
        }

        /// <summary>把子查询进度转交给批次级进度转换函数。</summary>
        private sealed class ForwardingProgress : IProgress<OperationProgress>
        {
            private readonly Action<OperationProgress> _report;

            public ForwardingProgress(Action<OperationProgress> report)
            {
                _report = report;
            }

            public void Report(OperationProgress value)
            {
                if (value != null) _report(value);
            }
        }
    }
}

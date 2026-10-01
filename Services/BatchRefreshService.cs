using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Excel;
using DB2Sheet.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.Services
{
    /// <summary>编排 SQL Sheet 批量任务的查询缓冲、错误隔离和 Excel 写入。</summary>
    /// <remarks>
    /// 并行模式只并发执行数据库查询；所有 Excel COM 写入都会回到启动调用的 UI 同步上下文并按任务顺序执行。
    /// 读取条只在查询结束时按已完成个数前进，进行中的读取没有百分比，写入快照也不改读取条。
    /// 批次开始只报告读取个数 0 和总数，不带写入百分比；写入百分比只在真正写表时由转发填写。
    /// 进入阻塞写入之前，在界面线程上同步报告当前读取个数并画出读取条。并行在写第一张表前把个数钉到总数。
    /// 目标表名重复时直接失败，不写入任何表。单个任务失败不会终止其他任务，用户取消会终止整个批次。
    /// </remarks>
    public sealed class BatchRefreshService : IBatchRefreshService
    {
        private readonly IQueryBufferService _queryBuffer;
        private readonly IExcelResultWriter _writer;
        private readonly ILogger _logger;

        /// <summary>创建批量刷新服务。</summary>
        /// <param name="queryBuffer">负责执行并缓冲查询结果的服务。</param>
        /// <param name="writer">负责写入 Excel 的服务。</param>
        /// <param name="logger">任务失败和批次汇总的文件日志。不按进度日志级别过滤。</param>
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
            string duplicateMessage = SqlSheetHeader.DescribeDuplicateTargets(tasks);
            if (!string.IsNullOrEmpty(duplicateMessage))
            {
                throw new InvalidOperationException(duplicateMessage);
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            List<RefreshTaskResult> results = new List<RefreshTaskResult>();
            List<RefreshTaskDefinition> executable = new List<RefreshTaskDefinition>(tasks);
            if (tasks.Count > 1)
            {
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Reading,
                    Message = "开始读取查询。",
                    TotalTasks = tasks.Count,
                    ReadTaskCount = 0,
                    Elapsed = stopwatch.Elapsed,
                    IsIndeterminate = true
                });
            }

            if (mode == BatchExecutionMode.Parallel && executable.Count > 1)
            {
                IReadOnlyList<BufferedTask> buffered = await BufferParallelAsync(
                    connection, executable, Math.Max(1, maximumParallelism), rowLimit, blockSize,
                    timeoutSeconds, operationId, progress, tasks.Count, stopwatch,
                    cancellationToken).ConfigureAwait(false);
                ReportReadFinishedOnUi(excelContext, progress, null, tasks.Count, tasks.Count, stopwatch);
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
                        results, tasks.Count, stopwatch, operationId, cancellationToken, false));
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
                            progress, () => results.Count, tasks.Count, stopwatch,                             cancellationToken).ConfigureAwait(false);
                        results.Add(WriteOnContext(
                            excelContext, workbook, task, buffered, progress,
                            results, tasks.Count, stopwatch, operationId, cancellationToken, true));
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        _logger.Write(LogSeverity.Error, "批量刷新任务失败。", operationId, exception, SheetProperty(task));
                        ReportReadFinished(progress, task, results.Count + 1, tasks.Count, stopwatch);
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
            _logger.Write(
                LogSeverity.Information,
                "批量刷新完成。",
                operationId,
                properties: BatchSummary(mode, maximumParallelism, rowLimit, blockSize, timeoutSeconds, tasks.Count, results, stopwatch.Elapsed));
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
            int totalTasks,
            Stopwatch stopwatch,
            CancellationToken cancellationToken)
        {
            SemaphoreSlim semaphore = new SemaphoreSlim(maximumParallelism, maximumParallelism);
            try
            {
                int completed = 0;
                Task<BufferedTask>[] operations = tasks.Select((task, index) => Task.Run(async () =>
                {
                    await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
                    try
                    {
                        BufferedQueryResult result = await BufferTaskAsync(
                            connection, task, rowLimit, blockSize, timeoutSeconds, operationId,
                            progress, () => Volatile.Read(ref completed), totalTasks, stopwatch,
                            cancellationToken).ConfigureAwait(false);
                        int finished = Interlocked.Increment(ref completed);
                        ReportReadFinished(progress, task, finished, totalTasks, stopwatch);
                        return new BufferedTask(index, task, result, null);
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception)
                    {
                        int finished = Interlocked.Increment(ref completed);
                        ReportReadFinished(progress, task, finished, totalTasks, stopwatch);
                        _logger.Write(LogSeverity.Error, "并行读取任务失败。", operationId, exception, SheetProperty(task));
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
            Func<int> finishedReads,
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
                value.CurrentTask = finishedReads();
                value.TotalTasks = totalTasks;
                value.Elapsed = stopwatch.Elapsed;
                progress?.Report(value);
            });
            return _queryBuffer.ExecuteAsync(request, operationId + "-" + task.Id, taskProgress, cancellationToken);
        }

        /// <summary>在 Excel UI 线程写入一个已缓冲的任务，并只推进写入进度条。</summary>
        /// <remarks>不填写读取百分比。串行时同一次 Send 先报告这条查询已结束，画出读取条，再阻塞写表。个数不到总数时不画成已读完。</remarks>
        private RefreshTaskResult WriteOnContext(
            SynchronizationContext context,
            ExcelInterop.Workbook workbook,
            RefreshTaskDefinition task,
            BufferedQueryResult buffered,
            IProgress<OperationProgress> progress,
            IReadOnlyList<RefreshTaskResult> completed,
            int totalTasks,
            Stopwatch stopwatch,
            string operationId,
            CancellationToken cancellationToken,
            bool reportReadBeforeWrite)
        {
            RefreshTaskResult result = null;
            Exception error = null;
            context.Send(state =>
            {
                try
                {
                    int writeTask = completed.Count + 1;
                    if (reportReadBeforeWrite)
                        ReportReadFinished(progress, task, writeTask, totalTasks, stopwatch);
                    IProgress<OperationProgress> writeProgress = new ForwardingProgress(value =>
                    {
                        value.TaskId = task.Id;
                        value.TargetSheetName = task.TargetSheetName;
                        value.CurrentTask = writeTask;
                        value.TotalTasks = totalTasks;
                        value.SucceededTasks = completed.Count(item => item.Succeeded);
                        value.FailedTasks = completed.Count(item => !item.Succeeded && !item.Skipped);
                        value.SkippedTasks = completed.Count(item => item.Skipped);
                        value.Elapsed = stopwatch.Elapsed;
                        value.WriteTaskCount = writeTask;
                        value.WritePercent = TrackPercent(writeTask, totalTasks, value.Percent);
                        progress?.Report(value);
                    });
                    long rows = _writer.WriteResult(
                        workbook,
                        task.TargetSheetName,
                        buffered,
                        SheetWriteOptions.Anchored(task.StartRow, task.StartColumnIndex, task.ClearExtraColumns),
                        writeProgress,
                        cancellationToken,
                        operationId + "-" + task.Id);
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
                _logger.Write(LogSeverity.Error, "写入 Excel 失败。", operationId + "-" + task.Id, error, SheetProperty(task));
                return Failed(task, error);
            }
            return result;
        }

        /// <summary>取出失败任务的目标表名，供文件日志关联。</summary>
        private static Dictionary<string, string> SheetProperty(RefreshTaskDefinition task)
        {
            return new Dictionary<string, string>
            {
                ["sheet"] = task.TargetSheetName ?? string.Empty
            };
        }

        /// <summary>组装批量刷新结束时的计数和当时使用的执行限制。</summary>
        private static Dictionary<string, string> BatchSummary(
            BatchExecutionMode mode,
            int maximumParallelism,
            int rowLimit,
            int blockSize,
            int timeoutSeconds,
            int totalTasks,
            IReadOnlyList<RefreshTaskResult> results,
            TimeSpan elapsed)
        {
            return new Dictionary<string, string>
            {
                ["mode"] = mode.ToString(),
                ["maxParallelism"] = maximumParallelism.ToString(),
                ["rowLimit"] = rowLimit.ToString(),
                ["blockSize"] = blockSize.ToString(),
                ["timeoutSeconds"] = timeoutSeconds.ToString(),
                ["totalTasks"] = totalTasks.ToString(),
                ["succeeded"] = results.Count(item => item.Succeeded).ToString(),
                ["failed"] = results.Count(item => !item.Succeeded && !item.Skipped).ToString(),
                ["skipped"] = results.Count(item => item.Skipped).ToString(),
                ["rowsWritten"] = results.Sum(item => item.RowsWritten).ToString(),
                ["elapsedMs"] = ((long)elapsed.TotalMilliseconds).ToString()
            };
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
                IsIndeterminate = stage != OperationStage.Completed,
                WriteTaskCount = totalTasks,
                WritePercent = stage == OperationStage.Completed ? 100 : (int?)null
            };
        }

        /// <summary>在界面线程上同步报告读取个数，返回后读取条已经画出。</summary>
        /// <param name="context">启动批量刷新的 Excel UI 同步上下文。</param>
        /// <param name="progress">进度报告。界面线程上会直接更新窗体。</param>
        /// <param name="task">关联的任务。整批读完时可以为空。</param>
        /// <param name="finished">已经结束的查询个数。</param>
        /// <param name="totalTasks">本批查询总数。</param>
        /// <param name="stopwatch">本批计时。</param>
        /// <remarks>必须在 <see cref="WriteOnContext"/> 的阻塞写入之前调用。Send 返回时读取条已经按这个个数画完。</remarks>
        private static void ReportReadFinishedOnUi(
            SynchronizationContext context,
            IProgress<OperationProgress> progress,
            RefreshTaskDefinition task,
            int finished,
            int totalTasks,
            Stopwatch stopwatch)
        {
            if (progress == null) return;
            context.Send(
                _ => ReportReadFinished(progress, task, finished, totalTasks, stopwatch),
                null);
        }

        /// <summary>某条查询已经结束时推进读取条。成功和失败都计数，取消不在这里报告。</summary>
        /// <remarks>只报告已结束个数和总数。读取没有任务内部的百分比，窗体按这一个比例画条。在界面线程上调用时会立刻画出。</remarks>
        private static void ReportReadFinished(
            IProgress<OperationProgress> progress,
            RefreshTaskDefinition task,
            int finished,
            int totalTasks,
            Stopwatch stopwatch)
        {
            bool allRead = totalTasks > 0 && finished >= totalTasks;
            progress?.Report(new OperationProgress
            {
                TaskId = task?.Id,
                TargetSheetName = task?.TargetSheetName,
                Stage = allRead ? OperationStage.Writing : OperationStage.Reading,
                Message = allRead ? "查询已全部读取，开始写入。" : "查询已结束。",
                CurrentTask = finished,
                TotalTasks = totalTasks,
                ReadTaskCount = finished,
                Elapsed = stopwatch.Elapsed,
                IsIndeterminate = true
            });
        }

        /// <summary>把当前任务序号和任务内部百分比合成整批的 0 到 100。</summary>
        private static int TrackPercent(int currentTask, int totalTasks, int? innerPercent)
        {
            if (totalTasks < 1) return 0;
            int completed = currentTask < 1 ? 0 : currentTask - 1;
            if (completed > totalTasks) completed = totalTasks;
            int inner = innerPercent ?? 0;
            if (inner < 0) inner = 0;
            if (inner > 100) inner = 100;
            long overall = (completed * 100L + inner) / totalTasks;
            if (overall < 0) return 0;
            if (overall > 100) return 100;
            return (int)overall;
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

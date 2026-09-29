using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.UI;

namespace DB2Sheet.Services
{
    /// <summary>使用模态进度窗体把 WinForms UI 与可取消异步任务连接起来。</summary>
    /// <remarks>
    /// 必须从具有同步上下文的 UI 线程调用。成功或取消后进度窗体自动关闭；发生错误时保留窗口并重新抛出异常。
    /// </remarks>
    public sealed class OperationRunner : IOperationRunner
    {
        private readonly ILogger _logger;

        /// <summary>创建操作运行器。</summary>
        /// <param name="logger">用于记录未处理任务异常的日志记录器。</param>
        public OperationRunner(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <inheritdoc/>
        public async Task<T> RunAsync<T>(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task<T>> operation)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));

            string operationId = Guid.NewGuid().ToString("N");
            using (CancellationTokenSource cancellation = new CancellationTokenSource())
            using (OperationProgressForm form = new OperationProgressForm(title, canCancel))
            {
                TaskCompletionSource<T> completion = new TaskCompletionSource<T>();

                form.CancelRequested += (sender, args) => cancellation.Cancel();
                Progress<OperationProgress> progress = new Progress<OperationProgress>(value =>
                {
                    value.OperationId = string.IsNullOrWhiteSpace(value.OperationId) ? operationId : value.OperationId;
                    form.UpdateProgress(value);
                });

                form.Shown += async (sender, args) =>
                {
                    try
                    {
                        T result = await operation(progress, cancellation.Token);
                        form.Complete(null, false);
                        completion.TrySetResult(result);
                    }
                    catch (OperationCanceledException)
                    {
                        form.Complete(null, true);
                        completion.TrySetCanceled();
                    }
                    catch (Exception exception)
                    {
                        form.Complete(exception, false);
                        _logger.Write(LogSeverity.Error, title + "失败。", operationId, exception);
                        completion.TrySetException(exception);
                    }
                };

                form.ShowDialog(owner);
                return await completion.Task;
            }
        }

        /// <inheritdoc/>
        public async Task RunAsync(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task> operation)
        {
            await RunAsync<object>(owner, title, canCancel, async (progress, cancellationToken) =>
            {
                await operation(progress, cancellationToken);
                return null;
            });
        }

        /// <inheritdoc/>
        public T Run<T>(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task<T>> operation)
        {
            return RunAsync(owner, title, canCancel, operation).GetAwaiter().GetResult();
        }

        /// <inheritdoc/>
        public void Run(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task> operation)
        {
            RunAsync(owner, title, canCancel, operation).GetAwaiter().GetResult();
        }
    }
}

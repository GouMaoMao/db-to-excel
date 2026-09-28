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
        public T Run<T>(
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
                form.CancelRequested += (sender, args) => cancellation.Cancel();
                Progress<OperationProgress> progress = new Progress<OperationProgress>(value =>
                {
                    value.OperationId = string.IsNullOrWhiteSpace(value.OperationId) ? operationId : value.OperationId;
                    form.UpdateProgress(value);
                });

                Task<T> task;
                try
                {
                    task = operation(progress, cancellation.Token);
                }
                catch (Exception exception)
                {
                    task = Task.FromException<T>(exception);
                }

                task.ContinueWith(completed =>
                {
                    Exception error = completed.IsFaulted ? completed.Exception?.GetBaseException() : null;
                    form.Complete(error, completed.IsCanceled);
                    if (error != null) _logger.Write(LogSeverity.Error, title + "失败。", operationId, error);
                }, CancellationToken.None, TaskContinuationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());

                form.ShowDialog(owner);
                return task.GetAwaiter().GetResult();
            }
        }

        /// <inheritdoc/>
        public void Run(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task> operation)
        {
            Run<object>(owner, title, canCancel, async (progress, cancellationToken) =>
            {
                await operation(progress, cancellationToken).ConfigureAwait(false);
                return null;
            });
        }
    }
}

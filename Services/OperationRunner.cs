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
    /// 进度报告在界面线程上直接更新窗体；后台线程才排队。写入用 Send 占住界面线程时，读取条必须在进入写入前已经画完。
    /// </remarks>
    public sealed class OperationRunner : IOperationRunner
    {
        private readonly ILogger _logger;
        private readonly ISettingsStore _settings;

        /// <summary>创建操作运行器。</summary>
        /// <param name="logger">用于记录未处理任务异常的日志记录器。文件日志不按级别过滤。</param>
        /// <param name="settings">读取进度日志最低显示级别。该值只影响进度窗口的日志列表。</param>
        public OperationRunner(ILogger logger, ISettingsStore settings)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
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
            LogSeverity minimumSeverity = _settings.Get(CoreSettings.LogLevel);
            using (CancellationTokenSource cancellation = new CancellationTokenSource())
            using (OperationProgressForm form = new OperationProgressForm(title, canCancel, minimumSeverity))
            {
                TaskCompletionSource<T> completion = new TaskCompletionSource<T>();

                form.CancelRequested += (sender, args) => cancellation.Cancel();
                UiProgress progress = new UiProgress(form, operationId);

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

        /// <summary>把进度快照送到进度窗。界面线程上立即更新，后台线程才排队。</summary>
        /// <remarks>
        /// <see cref="Progress{T}"/> 即使已经在界面线程上也会 Post。写入正在 Send 里阻塞时，排队的读取完成要等写入结束才画出来。
        /// </remarks>
        private sealed class UiProgress : IProgress<OperationProgress>
        {
            private readonly OperationProgressForm _form;
            private readonly string _operationId;

            /// <summary>创建与指定进度窗绑定的报告器。</summary>
            /// <param name="form">接收快照的进度窗。</param>
            /// <param name="operationId">本次操作标识。快照自己没带时补上。</param>
            public UiProgress(OperationProgressForm form, string operationId)
            {
                _form = form;
                _operationId = operationId;
            }

            /// <summary>更新进度窗。已在界面线程上则直接调用，否则排队。</summary>
            /// <param name="value">后台操作报告的状态快照。</param>
            public void Report(OperationProgress value)
            {
                if (value == null || _form.IsDisposed) return;
                value.OperationId = string.IsNullOrWhiteSpace(value.OperationId) ? _operationId : value.OperationId;
                if (_form.InvokeRequired)
                {
                    _form.BeginInvoke(new Action<OperationProgress>(_form.UpdateProgress), value);
                    return;
                }

                _form.UpdateProgress(value);
            }
        }
    }
}

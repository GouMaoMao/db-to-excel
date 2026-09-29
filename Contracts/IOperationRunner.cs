using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>在模态进度窗体中运行可取消的异步操作，并同步返回给 WinForms 调用方。</summary>
    /// <remarks>应从 UI 线程调用；实现负责显示进度、传递取消令牌和记录未处理异常。</remarks>
    public interface IOperationRunner
    {
        /// <summary>异步运行有返回值的操作。</summary>
        /// <typeparam name="T">操作结果类型。</typeparam>
        /// <param name="owner">进度窗体的所属窗口。</param>
        /// <param name="title">显示在进度窗体中的功能标题。</param>
        /// <param name="canCancel">是否向用户开放取消按钮。</param>
        /// <param name="operation">接收进度接收器和取消令牌的异步操作。</param>
        /// <returns>异步操作产生的结果任务。</returns>
        Task<T> RunAsync<T>(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task<T>> operation);

        /// <summary>异步运行无返回值的操作。</summary>
        /// <param name="owner">进度窗体的所属窗口。</param>
        /// <param name="title">显示在进度窗体中的功能标题。</param>
        /// <param name="canCancel">是否向用户开放取消按钮。</param>
        /// <param name="operation">接收进度接收器和取消令牌的异步操作。</param>
        /// <returns>异步操作任务。</returns>
        Task RunAsync(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task> operation);

        /// <summary>运行有返回值的异步操作。</summary>
        /// <typeparam name="T">操作结果类型。</typeparam>
        /// <param name="owner">进度窗体的所属窗口。</param>
        /// <param name="title">显示在进度窗体中的功能标题。</param>
        /// <param name="canCancel">是否向用户开放取消按钮。</param>
        /// <param name="operation">接收进度接收器和取消令牌的异步操作。</param>
        /// <returns>异步操作产生的结果。</returns>
        T Run<T>(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task<T>> operation);

        /// <summary>运行无返回值的异步操作。</summary>
        /// <param name="owner">进度窗体的所属窗口。</param>
        /// <param name="title">显示在进度窗体中的功能标题。</param>
        /// <param name="canCancel">是否向用户开放取消按钮。</param>
        /// <param name="operation">接收进度接收器和取消令牌的异步操作。</param>
        void Run(
            IWin32Window owner,
            string title,
            bool canCancel,
            Func<IProgress<OperationProgress>, CancellationToken, Task> operation);
    }
}

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>定义一种可注册的数据源类型及其校验、测试和查询能力。</summary>
    /// <remarks>实现应保持无状态或线程安全，因为同一实例可能被多个操作复用。</remarks>
    public interface IDataSourceProvider
    {
        /// <summary>获取稳定的内部提供程序标识，用于持久化和查找。</summary>
        string ProviderId { get; }
        /// <summary>获取显示给用户的数据源名称。</summary>
        string DisplayName { get; }
        /// <summary>获取该数据源支持的功能集合。</summary>
        ProviderCapabilities Capabilities { get; }
        /// <summary>获取构建连接配置所需的参数定义。</summary>
        IReadOnlyList<ParameterDefinition> ConnectionParameters { get; }

        /// <summary>检查连接配置并返回全部校验错误。</summary>
        /// <param name="profile">待校验的连接快照。</param>
        /// <returns>错误消息列表；空列表表示通过。</returns>
        IReadOnlyList<string> ValidateConnection(ConnectionProfileSnapshot profile);
        /// <summary>检查查询请求并返回全部校验错误。</summary>
        /// <param name="request">待校验的数据源请求。</param>
        /// <returns>错误消息列表；空列表表示通过。</returns>
        IReadOnlyList<string> ValidateRequest(DataSourceRequest request);
        /// <summary>异步测试连接配置。</summary>
        /// <param name="profile">待测试的连接快照。</param>
        /// <param name="progress">可选的进度接收器。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>连接测试结果。</returns>
        Task<ConnectionTestResult> TestConnectionAsync(
            ConnectionProfileSnapshot profile,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken);
        /// <summary>异步执行查询并返回流式结果。</summary>
        /// <param name="request">查询及限制条件。</param>
        /// <param name="progress">可选的进度接收器。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>由调用方负责释放的结果流。</returns>
        Task<IDataResultStream> ExecuteAsync(
            DataSourceRequest request,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken);
    }

    /// <summary>表示可按块异步读取的数据查询结果。</summary>
    /// <remarks>使用结束后必须调用 <see cref="IDisposable.Dispose"/>；<see cref="Cancel"/> 用于主动中止底层命令。</remarks>
    public interface IDataResultStream : IDisposable
    {
        /// <summary>获取结果集列结构。</summary>
        IReadOnlyList<ResultColumn> Columns { get; }
        /// <summary>获取目前累计读取的行数。</summary>
        long RowsRead { get; }
        /// <summary>获取结果是否已经全部读取完毕。</summary>
        bool IsCompleted { get; }
        /// <summary>获取结果是否因行数上限而被截断。</summary>
        bool IsTruncated { get; }
        /// <summary>读取下一批结果行。</summary>
        /// <param name="maximumRows">本批最多读取的行数，必须为正数。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>包含二维值数组、实际行数和完成状态的结果块。</returns>
        Task<ResultBlock> ReadBlockAsync(int maximumRows, CancellationToken cancellationToken);
        /// <summary>请求取消底层数据库命令。</summary>
        void Cancel();
    }
}

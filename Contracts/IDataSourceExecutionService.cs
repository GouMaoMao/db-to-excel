using System;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>
    /// 统一编排数据源查找、请求校验、执行和连接测试。
    /// </summary>
    /// <remarks>调用方无需直接选择具体数据库提供程序，但必须释放返回的结果流。</remarks>
    public interface IDataSourceExecutionService
    {
        /// <summary>验证并执行数据源请求，返回可分块读取的结果流。</summary>
        /// <param name="request">连接、查询和执行限制等输入。</param>
        /// <param name="operationId">用于关联进度与日志的操作标识。</param>
        /// <param name="progress">可选的进度接收器。</param>
        /// <param name="cancellationToken">用于取消连接、查询或读取的令牌。</param>
        /// <returns>由调用方负责释放的数据结果流。</returns>
        Task<IDataResultStream> ExecuteAsync(
            DataSourceRequest request,
            string operationId,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken);

        /// <summary>验证连接配置并测试是否能够访问数据源。</summary>
        /// <param name="profile">待测试的连接配置快照。</param>
        /// <param name="operationId">用于关联进度与日志的操作标识。</param>
        /// <param name="progress">可选的进度接收器。</param>
        /// <param name="cancellationToken">用于取消测试的令牌。</param>
        /// <returns>包含成功状态、提示消息和耗时的测试结果。</returns>
        Task<ConnectionTestResult> TestConnectionAsync(
            ConnectionProfileSnapshot profile,
            string operationId,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken);
    }
}

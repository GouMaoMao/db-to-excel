using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>定义关系型数据库提供程序的数据库目录浏览能力。</summary>
    /// <remarks>
    /// 实现负责创建、打开并释放元数据查询所需的连接、命令和读取器。方法可在后台线程完成，
    /// 调用方必须使用取消令牌停止不再需要的加载，并将返回结果切换到 UI 线程后再绑定控件。
    /// </remarks>
    public interface IDatabaseMetadataProvider
    {
        /// <summary>异步列举当前账号可见的数据库。</summary>
        /// <param name="profile">作为服务器地址和身份凭据来源的连接快照。</param>
        /// <param name="cancellationToken">用于取消连接和目录查询的令牌。</param>
        /// <returns>按名称排序的数据库元数据只读列表。</returns>
        /// <exception cref="System.OperationCanceledException">操作被取消。</exception>
        /// <exception cref="System.Data.Common.DbException">连接失败或账号无权读取数据库目录。</exception>
        Task<IReadOnlyList<DatabaseMetadata>> GetDatabasesAsync(
            ConnectionProfileSnapshot profile,
            CancellationToken cancellationToken);

        /// <summary>异步列举指定数据库内当前账号可见的表和视图。</summary>
        /// <param name="profile">作为服务器地址和身份凭据来源的连接快照。</param>
        /// <param name="databaseName">要读取的数据库名称。</param>
        /// <param name="cancellationToken">用于取消连接和目录查询的令牌。</param>
        /// <returns>按架构、类型和名称排序的数据库对象只读列表。</returns>
        /// <exception cref="System.OperationCanceledException">操作被取消。</exception>
        /// <exception cref="System.Data.Common.DbException">数据库不可访问或目录查询失败。</exception>
        Task<IReadOnlyList<DatabaseObjectMetadata>> GetDatabaseObjectsAsync(
            ConnectionProfileSnapshot profile,
            string databaseName,
            CancellationToken cancellationToken);

        /// <summary>从持久化连接派生使用指定数据库的临时快照。</summary>
        /// <param name="profile">原始连接快照。</param>
        /// <param name="databaseName">查询执行应使用的数据库名称。</param>
        /// <returns>参数已复制且数据库已替换的临时快照；不会修改仓储中的方案。</returns>
        ConnectionProfileSnapshot CreateDatabaseSnapshot(
            ConnectionProfileSnapshot profile,
            string databaseName);
    }
}

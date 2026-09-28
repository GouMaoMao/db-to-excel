using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>
    /// 定义关系型数据库查询提供程序需要实现的底层 ADO.NET 对象创建能力。
    /// </summary>
    /// <remarks>
    /// 调用方负责释放本接口创建的连接和命令。实现应只创建对象，不应在创建方法中隐式执行查询。
    /// </remarks>
    public interface IDatabaseQueryProvider : IDataSourceProvider
    {
        /// <summary>根据不可变连接快照创建尚未使用的数据库连接。</summary>
        /// <param name="profile">包含提供程序标识和连接参数的快照。</param>
        /// <returns>由调用方负责释放的数据库连接。</returns>
        DbConnection CreateConnection(ConnectionProfileSnapshot profile);

        /// <summary>为已创建的连接构造查询命令。</summary>
        /// <param name="connection">命令所属的数据库连接。</param>
        /// <param name="queryText">要执行的 SQL 文本。</param>
        /// <param name="timeoutSeconds">命令超时秒数。</param>
        /// <returns>由调用方负责释放的数据库命令。</returns>
        DbCommand CreateCommand(DbConnection connection, string queryText, int timeoutSeconds);

        /// <summary>在当前连接上启用数据库支持的只读会话约束。</summary>
        /// <param name="connection">已经打开的数据库连接。</param>
        /// <param name="cancellationToken">用于取消异步配置的令牌。</param>
        /// <returns>表示配置完成的任务。</returns>
        Task ConfigureReadOnlySessionAsync(DbConnection connection, CancellationToken cancellationToken);
    }
}

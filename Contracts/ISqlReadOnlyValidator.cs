using System.Collections.Generic;

namespace DB2Sheet.Contracts
{
    /// <summary>检查 SQL 是否符合插件的只读查询限制。</summary>
    /// <remarks>该校验用于降低误操作风险，不替代数据库账号自身的最小权限控制。</remarks>
    public interface ISqlReadOnlyValidator
    {
        /// <summary>检查 SQL 文本并返回全部违规原因。</summary>
        /// <param name="sql">待检查的 SQL 文本。</param>
        /// <param name="providerId">用于应用数据库方言规则的提供程序标识。</param>
        /// <returns>错误消息列表；空列表表示通过。</returns>
        IReadOnlyList<string> Validate(string sql, string providerId);
    }
}

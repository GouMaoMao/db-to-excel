using System.Collections.Generic;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>为 SQL 编辑器提供无数据库副作用的词法分析和文本格式化。</summary>
    public interface ISqlFormattingService
    {
        /// <summary>将 SQL 分解为可用于语法着色的非空白标记。</summary>
        /// <param name="sql">待分析 SQL。</param>
        /// <returns>按原始位置排序的只读标记列表。</returns>
        IReadOnlyList<SqlToken> Tokenize(string sql);

        /// <summary>按常见只读查询结构生成统一换行、大小写和缩进。</summary>
        /// <param name="sql">待格式化 SQL；字符串、注释及引号标识符内容保持不变。</param>
        /// <returns>格式化后的 SQL。</returns>
        string Format(string sql);
    }
}

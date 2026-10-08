namespace DB2Sheet.Contracts
{
    /// <summary>为 SQL 编辑器提供无数据库副作用的文本格式化。</summary>
    public interface ISqlFormattingService
    {
        /// <summary>按常见只读查询结构生成统一换行、关键字小写和按列对齐。</summary>
        /// <param name="sql">待格式化 SQL；字符串、注释及引号标识符内容保持不变。</param>
        /// <returns>格式化后的 SQL。子句关键字与第一项同一行，后续列表项及 AND/OR 对齐到该项起始列。</returns>
        string Format(string sql);
    }
}

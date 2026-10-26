namespace DB2Sheet.Contracts
{
    /// <summary>为 SQL 编辑器提供无数据库副作用的文本格式化。</summary>
    public interface ISqlFormattingService
    {
        /// <summary>按常见只读查询结构生成紧凑换行、关键字小写和两空格缩进。</summary>
        /// <param name="sql">待格式化 SQL；字符串、注释及引号标识符内容保持不变。</param>
        /// <returns>格式化后的 SQL。布尔条件分行，字段、CASE 和 BETWEEN 在行宽允许时保持紧凑。</returns>
        string Format(string sql);
    }
}

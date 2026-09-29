namespace DB2Sheet.Models
{
    /// <summary>区分 SQL 编辑器用于着色和格式化的词法标记类型。</summary>
    public enum SqlTokenKind
    {
        /// <summary>SQL 保留字或常用子句关键字。</summary>
        Keyword,
        /// <summary>普通名称。</summary>
        Word,
        /// <summary>单引号字符串。</summary>
        String,
        /// <summary>单行或块注释。</summary>
        Comment,
        /// <summary>整数或小数。</summary>
        Number,
        /// <summary>以 @、:、$ 或 ? 表示的参数。</summary>
        Parameter,
        /// <summary>由双引号、反引号或方括号包围的标识符。</summary>
        QuotedIdentifier,
        /// <summary>运算符、标点或括号。</summary>
        Symbol
    }

    /// <summary>描述 SQL 文本中一个保持原始位置和值的词法标记。</summary>
    public sealed class SqlToken
    {
        /// <summary>创建 SQL 词法标记。</summary>
        /// <param name="kind">标记类型。</param>
        /// <param name="start">在原始 SQL 中的起始字符索引。</param>
        /// <param name="length">标记字符数。</param>
        /// <param name="text">标记原始文本。</param>
        public SqlToken(SqlTokenKind kind, int start, int length, string text)
        {
            Kind = kind;
            Start = start;
            Length = length;
            Text = text ?? string.Empty;
        }

        /// <summary>获取标记类型。</summary>
        public SqlTokenKind Kind { get; }
        /// <summary>获取标记在原始 SQL 中的起始索引。</summary>
        public int Start { get; }
        /// <summary>获取标记字符数。</summary>
        public int Length { get; }
        /// <summary>获取标记原始文本。</summary>
        public string Text { get; }
    }
}

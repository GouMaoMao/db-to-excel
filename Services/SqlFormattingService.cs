using DB2Sheet.Contracts;
using PoorMansTSqlFormatterLib;
using PoorMansTSqlFormatterLib.Formatters;

namespace DB2Sheet.Services
{
    /// <summary>使用 Poor Man's T-SQL Formatter 为查询编辑器提供紧凑、稳定的 SQL 排版。</summary>
    /// <remarks>
    /// 服务不连接数据库，也不修改编辑器。成熟库负责词法分析和结构排版；这里仅固定应用需要的风格。
    /// 字段、CASE 和 BETWEEN 保持紧凑，布尔条件分行，并保留原有行尾注释的归属。
    /// 格式化器按 T-SQL 语法设计；其他方言采用其容错解析能力，不保证改写方言专有语法。
    /// </remarks>
    public sealed class SqlFormattingService : ISqlFormattingService
    {
        private readonly SqlFormattingManager _manager;

        /// <summary>创建使用两空格缩进和小写关键字的共享格式化服务。</summary>
        public SqlFormattingService()
        {
            TSqlStandardFormatter formatter = new TSqlStandardFormatter
            {
                IndentString = "  ",
                SpacesPerTab = 2,
                MaxLineWidth = 120,
                ExpandCommaLists = false,
                TrailingCommas = false,
                SpaceAfterExpandedComma = false,
                ExpandBooleanExpressions = true,
                ExpandCaseStatements = false,
                ExpandBetweenConditions = false,
                UppercaseKeywords = false,
                BreakJoinOnSections = false,
                HTMLColoring = false
            };
            _manager = new SqlFormattingManager(formatter);
        }

        /// <summary>格式化一段 SQL，并保留字符串、标识符和注释正文。</summary>
        /// <param name="sql">待格式化 SQL。空值和空白文本返回空字符串。</param>
        /// <returns>紧凑排版的 SQL；换行由格式化库使用当前平台换行符。</returns>
        public string Format(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql)) return string.Empty;
            return _manager.Format(sql);
        }
    }
}
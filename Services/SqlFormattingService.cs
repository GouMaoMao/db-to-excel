using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>以轻量词法分析实现常见查询语句的语法标记和结构化格式化。</summary>
    /// <remarks>服务不连接数据库且不修改 SQL 语义；它覆盖四类已支持数据库的常用查询语法，但不是完整方言解析器。</remarks>
    public sealed class SqlFormattingService : ISqlFormattingService
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "all", "and", "as", "asc", "between", "by", "case", "cast", "cross", "current", "desc", "distinct",
            "else", "end", "except", "exists", "false", "fetch", "first", "following", "for", "from", "full",
            "group", "having", "in", "inner", "intersect", "into", "is", "join", "last", "lateral", "left", "like",
            "limit", "not", "null", "nulls", "offset", "on", "or", "order", "outer", "over", "partition", "preceding",
            "recursive", "right", "rows", "select", "then", "true", "union", "using", "when", "where", "window", "with"
        };

        private static readonly HashSet<string> LineKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "WITH", "SELECT", "FROM", "WHERE", "HAVING", "UNION", "INTERSECT", "EXCEPT", "LIMIT", "OFFSET", "FETCH", "WINDOW"
        };

        private static readonly HashSet<string> JoinKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "JOIN", "LEFT", "RIGHT", "FULL", "INNER", "CROSS"
        };

        /// <inheritdoc/>
        public IReadOnlyList<SqlToken> Tokenize(string sql)
        {
            string value = sql ?? string.Empty;
            List<SqlToken> tokens = new List<SqlToken>();
            int index = 0;
            while (index < value.Length)
            {
                if (char.IsWhiteSpace(value[index]))
                {
                    index++;
                    continue;
                }

                int start = index;
                char current = value[index];
                char next = index + 1 < value.Length ? value[index + 1] : '\0';
                if (current == '-' && next == '-')
                {
                    index += 2;
                    while (index < value.Length && value[index] != '\r' && value[index] != '\n') index++;
                    Add(tokens, SqlTokenKind.Comment, value, start, index);
                    continue;
                }
                if (current == '/' && next == '*')
                {
                    index += 2;
                    while (index < value.Length && !(value[index - 1] == '*' && value[index] == '/')) index++;
                    if (index < value.Length) index++;
                    Add(tokens, SqlTokenKind.Comment, value, start, index);
                    continue;
                }
                if (current == '\'' || current == '"' || current == '`' || current == '[')
                {
                    char closing = current == '[' ? ']' : current;
                    bool isString = current == '\'';
                    index++;
                    while (index < value.Length)
                    {
                        if (value[index] == closing)
                        {
                            if (index + 1 < value.Length && value[index + 1] == closing)
                            {
                                index += 2;
                                continue;
                            }
                            index++;
                            break;
                        }
                        index++;
                    }
                    Add(tokens, isString ? SqlTokenKind.String : SqlTokenKind.QuotedIdentifier, value, start, index);
                    continue;
                }
                if (char.IsDigit(current))
                {
                    index++;
                    while (index < value.Length && (char.IsDigit(value[index]) || value[index] == '.')) index++;
                    Add(tokens, SqlTokenKind.Number, value, start, index);
                    continue;
                }
                if ((current == '@' || current == ':' || current == '$' || current == '?') &&
                    (next == '\0' || char.IsLetterOrDigit(next) || next == '_'))
                {
                    index++;
                    while (index < value.Length && (char.IsLetterOrDigit(value[index]) || value[index] == '_')) index++;
                    Add(tokens, SqlTokenKind.Parameter, value, start, index);
                    continue;
                }
                if (char.IsLetter(current) || current == '_')
                {
                    index++;
                    while (index < value.Length && (char.IsLetterOrDigit(value[index]) || value[index] == '_' || value[index] == '$')) index++;
                    string text = value.Substring(start, index - start);
                    tokens.Add(new SqlToken(Keywords.Contains(text) ? SqlTokenKind.Keyword : SqlTokenKind.Word, start, index - start, text));
                    continue;
                }

                index++;
                if (index < value.Length && IsPairOperator(current, value[index])) index++;
                Add(tokens, SqlTokenKind.Symbol, value, start, index);
            }
            return tokens.AsReadOnly();
        }

        /// <inheritdoc/>
        public string Format(string sql)
        {
            IReadOnlyList<SqlToken> tokens = Tokenize(sql);
            if (tokens.Count == 0) return string.Empty;

            StringBuilder result = new StringBuilder();
            int indent = 0;
            bool selectList = false;
            for (int index = 0; index < tokens.Count; index++)
            {
                SqlToken token = tokens[index];
                string upper = token.Kind == SqlTokenKind.Keyword ? token.Text.ToUpperInvariant() : token.Text;
                string next = index + 1 < tokens.Count && tokens[index + 1].Kind == SqlTokenKind.Keyword
                    ? tokens[index + 1].Text.ToUpperInvariant()
                    : string.Empty;

                if (upper == ")")
                {
                    indent = Math.Max(0, indent - 1);
                    TrimTrailingSpace(result);
                    result.Append(')');
                    continue;
                }
                if (upper == "(")
                {
                    TrimTrailingSpace(result);
                    result.Append('(');
                    indent++;
                    continue;
                }
                if (upper == ",")
                {
                    TrimTrailingSpace(result);
                    result.Append(',');
                    if (selectList) NewLine(result, indent);
                    else result.Append(' ');
                    continue;
                }
                if (upper == ";")
                {
                    TrimTrailingSpace(result);
                    result.Append(';');
                    if (index + 1 < tokens.Count) NewLine(result, 0);
                    selectList = false;
                    continue;
                }
                if (token.Kind == SqlTokenKind.Comment)
                {
                    if (result.Length > 0 && !EndsWithWhiteSpace(result)) NewLine(result, indent);
                    result.Append(token.Text.TrimEnd());
                    if (index + 1 < tokens.Count) NewLine(result, indent);
                    continue;
                }

                bool groupBy = upper == "GROUP" && next == "BY";
                bool orderBy = upper == "ORDER" && next == "BY";
                bool startsLine = LineKeywords.Contains(upper) || JoinKeywords.Contains(upper) || groupBy || orderBy;
                if (startsLine && result.Length > 0)
                {
                    NewLine(result, indent);
                }
                if (upper == "SELECT") selectList = true;
                else if (upper == "FROM") selectList = false;

                if (upper == "CASE")
                {
                    AppendToken(result, upper);
                    indent++;
                    NewLine(result, indent);
                    continue;
                }
                if (upper == "WHEN" || upper == "ELSE")
                {
                    NewLine(result, indent);
                }
                if (upper == "END")
                {
                    indent = Math.Max(0, indent - 1);
                    NewLine(result, indent);
                }

                AppendToken(result, upper);
                if ((groupBy || orderBy) && index + 1 < tokens.Count)
                {
                    AppendToken(result, tokens[++index].Text.ToUpperInvariant());
                }
                if (upper == "ON") NewLine(result, indent + 1);
            }
            return result.ToString().Trim();
        }

        private static void Add(ICollection<SqlToken> tokens, SqlTokenKind kind, string source, int start, int end)
        {
            tokens.Add(new SqlToken(kind, start, end - start, source.Substring(start, end - start)));
        }

        private static bool IsPairOperator(char first, char second)
        {
            return (first == '<' && (second == '=' || second == '>')) ||
                   (first == '>' && second == '=') ||
                   (first == '!' && second == '=') ||
                   (first == ':' && second == ':') ||
                   (first == '|' && second == '|');
        }

        private static void AppendToken(StringBuilder result, string text)
        {
            if (result.Length > 0 && !EndsWithWhiteSpace(result) && result[result.Length - 1] != '(' && text != ".")
                result.Append(' ');
            if (text == ".") TrimTrailingSpace(result);
            result.Append(text);
        }

        private static void NewLine(StringBuilder result, int indent)
        {
            while (result.Length > 0 && char.IsWhiteSpace(result[result.Length - 1])) result.Length--;
            if (result.Length > 0) result.AppendLine();
            result.Append(new string(' ', Math.Max(0, indent) * 4));
        }

        private static bool EndsWithWhiteSpace(StringBuilder value)
        {
            return value.Length > 0 && char.IsWhiteSpace(value[value.Length - 1]);
        }

        private static void TrimTrailingSpace(StringBuilder value)
        {
            while (value.Length > 0 && value[value.Length - 1] == ' ') value.Length--;
        }
    }
}

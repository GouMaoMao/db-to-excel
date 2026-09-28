using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DB2Sheet.Contracts;

namespace DB2Sheet.Services
{
    /// <summary>通过词法清理和关键字规则检查 SQL 是否为单条只读 SELECT 或 WITH 查询。</summary>
    /// <remarks>
    /// 校验会忽略注释、字符串和带引号标识符，以减少误判；它不是完整 SQL 解析器，也不能替代数据库只读账号和只读会话。
    /// </remarks>
    public sealed class SqlReadOnlyValidator : ISqlReadOnlyValidator
    {
        private static readonly HashSet<string> ForbiddenTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "insert", "update", "delete", "merge", "upsert", "replace",
            "create", "alter", "drop", "truncate", "rename",
            "grant", "revoke", "deny", "execute", "exec", "call", "do",
            "copy", "load", "backup", "restore", "vacuum", "reindex",
            "attach", "detach", "pragma", "lock", "unlock"
        };

        /// <inheritdoc/>
        public IReadOnlyList<string> Validate(string sql, string providerId)
        {
            List<string> errors = new List<string>();
            if (string.IsNullOrWhiteSpace(sql))
            {
                errors.Add("查询内容不能为空。");
                return errors;
            }

            string cleaned;
            try
            {
                cleaned = RemoveCommentsAndLiterals(sql);
            }
            catch (FormatException exception)
            {
                errors.Add(exception.Message);
                return errors;
            }

            string trimmed = cleaned.Trim();
            int semicolon = trimmed.IndexOf(';');
            if (semicolon >= 0 && semicolon != trimmed.Length - 1)
            {
                errors.Add("仅允许执行单条查询语句。");
            }

            trimmed = trimmed.TrimEnd(';').Trim();
            MatchCollection matches = Regex.Matches(trimmed, @"[A-Za-z_][A-Za-z0-9_$]*");
            List<string> tokens = matches.Cast<Match>().Select(match => match.Value).ToList();
            if (tokens.Count == 0 ||
                (!tokens[0].Equals("select", StringComparison.OrdinalIgnoreCase) &&
                 !tokens[0].Equals("with", StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add("仅允许 SELECT 或只读 WITH 查询。");
            }

            string forbidden = tokens.FirstOrDefault(token => ForbiddenTokens.Contains(token));
            if (forbidden != null)
            {
                errors.Add("查询包含不允许的关键字：" + forbidden.ToUpperInvariant());
            }

            if (tokens.Any(token => token.Equals("into", StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add("只读查询不允许 INTO 子句。");
            }

            if (ContainsSequence(tokens, "for", "update") ||
                ContainsSequence(tokens, "for", "share"))
            {
                errors.Add("只读查询不允许行锁定子句。");
            }

            if (string.Equals(providerId, "postgresql", StringComparison.OrdinalIgnoreCase) &&
                tokens.Any(token => token.Equals("returning", StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add("PostgreSQL 只读查询不允许 RETURNING 子句。");
            }

            return errors.Distinct().ToList().AsReadOnly();
        }

        private static bool ContainsSequence(IReadOnlyList<string> tokens, string first, string second)
        {
            for (int index = 0; index < tokens.Count - 1; index++)
            {
                if (tokens[index].Equals(first, StringComparison.OrdinalIgnoreCase) &&
                    tokens[index + 1].Equals(second, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static string RemoveCommentsAndLiterals(string sql)
        {
            StringBuilder result = new StringBuilder(sql.Length);
            for (int index = 0; index < sql.Length; index++)
            {
                char current = sql[index];
                char next = index + 1 < sql.Length ? sql[index + 1] : '\0';

                if (current == '-' && next == '-')
                {
                    result.Append("  ");
                    index += 2;
                    while (index < sql.Length && sql[index] != '\r' && sql[index] != '\n')
                    {
                        result.Append(' ');
                        index++;
                    }
                    if (index < sql.Length) result.Append(sql[index]);
                    continue;
                }

                if (current == '/' && next == '*')
                {
                    result.Append("  ");
                    index += 2;
                    bool closed = false;
                    while (index < sql.Length)
                    {
                        if (index + 1 < sql.Length && sql[index] == '*' && sql[index + 1] == '/')
                        {
                            result.Append("  ");
                            index++;
                            closed = true;
                            break;
                        }
                        result.Append(char.IsWhiteSpace(sql[index]) ? sql[index] : ' ');
                        index++;
                    }
                    if (!closed) throw new FormatException("SQL 块注释未闭合。");
                    continue;
                }

                if (current == '\'' || current == '"' || current == '`' || current == '[')
                {
                    char closing = current == '[' ? ']' : current;
                    result.Append(' ');
                    bool closed = false;
                    for (index++; index < sql.Length; index++)
                    {
                        if (sql[index] == closing)
                        {
                            if (closing != ']' && index + 1 < sql.Length && sql[index + 1] == closing)
                            {
                                result.Append("  ");
                                index++;
                                continue;
                            }
                            result.Append(' ');
                            closed = true;
                            break;
                        }
                        result.Append(char.IsWhiteSpace(sql[index]) ? sql[index] : ' ');
                    }
                    if (!closed) throw new FormatException("SQL 字符串或标识符未闭合。");
                    continue;
                }

                result.Append(current);
            }

            return result.ToString();
        }
    }
}

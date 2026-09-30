using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using DB2Sheet.Models;

namespace DB2Sheet.Excel
{
    /// <summary>解析 SQL 页首行里的目标表名和 <c>//</c> 后的特殊参数。</summary>
    /// <remarks>
    /// 表名不能包含 <c>/</c>，因此第一个 <c>//</c> 是稳定分隔。
    /// 当前会执行的键只有「起始单元格」和「清空多余列」；其他键原样保留，供以后扩展。
    /// </remarks>
    public sealed class SqlSheetHeader
    {
        /// <summary>起始单元格参数名。值是结果表左上角，含标题行。</summary>
        public const string StartCellKey = "起始单元格";

        /// <summary>清空多余列参数名。值为「是」或「否」。</summary>
        public const string ClearExtraColumnsKey = "清空多余列";

        private const int MaxRow = 1048576;
        private const int MaxColumn = 16384;

        private SqlSheetHeader(
            string sheetName,
            string optionsText,
            string startCell,
            int startRow,
            int startColumn,
            bool clearExtraColumns,
            string error,
            bool normalized)
        {
            SheetName = sheetName ?? string.Empty;
            OptionsText = optionsText ?? string.Empty;
            StartCell = string.IsNullOrWhiteSpace(startCell) ? "A1" : startCell;
            StartRow = startRow < 1 ? 1 : startRow;
            StartColumn = startColumn < 1 ? 1 : startColumn;
            ClearExtraColumns = clearExtraColumns;
            Error = error ?? string.Empty;
            Normalized = normalized;
        }

        /// <summary>获取目标表名。</summary>
        public string SheetName { get; }
        /// <summary>获取 <c>//</c> 后的特殊参数文本。没有参数时为空。</summary>
        public string OptionsText { get; }
        /// <summary>获取结果表左上角单元格。</summary>
        public string StartCell { get; }
        /// <summary>获取起始行号，从 1 开始。</summary>
        public int StartRow { get; }
        /// <summary>获取起始列号，从 1 开始。</summary>
        public int StartColumn { get; }
        /// <summary>获取是否清空结果没有覆盖到的原数据列。未写该参数时为 true。</summary>
        public bool ClearExtraColumns { get; }
        /// <summary>获取已知参数的解析错误。没有错误时为空。</summary>
        public string Error { get; }
        /// <summary>获取参数文本是否被规范，需要写回单元格。</summary>
        public bool Normalized { get; }

        /// <summary>获取写回首行的完整文本。</summary>
        public string HeaderText
        {
            get
            {
                return string.IsNullOrEmpty(OptionsText) ? SheetName : SheetName + "//" + OptionsText;
            }
        }

        /// <summary>解析首行单元格。</summary>
        /// <param name="cellText">第 1 行的原始文本。</param>
        /// <returns>表名、参数和已知键的执行值。表名为空时调用方应跳过该列。</returns>
        public static SqlSheetHeader Parse(string cellText)
        {
            string text = (cellText ?? string.Empty).Trim();
            int separator = text.IndexOf("//", StringComparison.Ordinal);
            if (separator < 0)
            {
                return Defaults(text, string.Empty, false);
            }

            string sheetName = text.Substring(0, separator).Trim();
            string optionsText = text.Substring(separator + 2).Trim();
            if (optionsText.Length == 0)
            {
                return Defaults(sheetName, string.Empty, false);
            }

            return ParseOptions(sheetName, optionsText);
        }

        /// <summary>用已有表名和新的参数文本重新解析。参数为空时首行只保留表名。</summary>
        /// <param name="sheetName">目标表名。</param>
        /// <param name="optionsText">用户编辑后的特殊参数文本。</param>
        /// <returns>可写回的首行解析结果。值为默认的已知键不写入；因此两个已知值都是默认且没有其他键时，首行只留表名。</returns>
        public static SqlSheetHeader ParseOptions(string sheetName, string optionsText)
        {
            string name = (sheetName ?? string.Empty).Trim();
            string raw = (optionsText ?? string.Empty).Trim();
            if (raw.Length == 0)
            {
                return Defaults(name, string.Empty, false);
            }

            string[] segments = raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> rebuilt = new List<string>();
            string startCell = "A1";
            int startRow = 1;
            int startColumn = 1;
            bool clearExtraColumns = true;
            string error = null;

            foreach (string segment in segments)
            {
                string item = segment.Trim();
                if (item.Length == 0) continue;
                int equals = item.IndexOf('=');
                if (equals <= 0)
                {
                    error = error ?? "特殊参数格式无效：" + item;
                    rebuilt.Add(item);
                    continue;
                }

                string key = item.Substring(0, equals).Trim();
                string value = item.Substring(equals + 1).Trim();
                if (key.Length == 0)
                {
                    error = error ?? "特殊参数格式无效：" + item;
                    rebuilt.Add(item);
                    continue;
                }

                if (string.Equals(key, StartCellKey, StringComparison.Ordinal))
                {
                    if (!TryParseCell(value, out int row, out int column, out string normalized))
                    {
                        error = error ?? "起始单元格无效：" + value;
                        rebuilt.Add(key + "=" + value);
                    }
                    else
                    {
                        startCell = normalized;
                        startRow = row;
                        startColumn = column;
                        rebuilt.Add(key + "=" + normalized);
                    }
                    continue;
                }

                if (string.Equals(key, ClearExtraColumnsKey, StringComparison.Ordinal))
                {
                    if (value == "是") clearExtraColumns = true;
                    else if (value == "否") clearExtraColumns = false;
                    else error = error ?? "清空多余列只能是「是」或「否」。";
                    rebuilt.Add(key + "=" + value);
                    continue;
                }

                rebuilt.Add(key + "=" + value);
            }

            if (!string.IsNullOrEmpty(error))
            {
                return new SqlSheetHeader(name, raw, "A1", 1, 1, true, error, false);
            }

            List<string> stored = new List<string>();
            foreach (string item in rebuilt)
            {
                if (string.Equals(item, StartCellKey + "=A1", StringComparison.Ordinal)) continue;
                if (string.Equals(item, ClearExtraColumnsKey + "=是", StringComparison.Ordinal)) continue;
                stored.Add(item);
            }

            string canonical = string.Join(";", stored);
            if (canonical.Length == 0)
            {
                return Defaults(name, string.Empty, raw.Length > 0);
            }

            return new SqlSheetHeader(
                name,
                canonical,
                startCell,
                startRow,
                startColumn,
                clearExtraColumns,
                string.Empty,
                !string.Equals(canonical, raw, StringComparison.Ordinal));
        }

        /// <summary>取出 SQL 命令开始前的注释，作为列表摘要。</summary>
        /// <param name="queryText">第 2 行起拼接后的 SQL，可含前置注释。</param>
        /// <returns>去掉注释记号后的文本。没有前置注释时为空。命令开始后的注释不计入。</returns>
        public static string ExtractLeadingComment(string queryText)
        {
            List<string> parts = new List<string>();
            SkipLeadingComments(queryText ?? string.Empty, parts);
            return string.Join(" ", parts);
        }

        /// <summary>判断文本去掉前置注释和空白后是否已经没有 SQL。</summary>
        /// <param name="queryText">已拼好的单元格文本。</param>
        /// <returns>只有注释或空白时为 true。出现命令记号后为 false。</returns>
        public static bool IsOnlyLeadingComment(string queryText)
        {
            string query = queryText ?? string.Empty;
            int index = SkipLeadingComments(query, null);
            while (index < query.Length && char.IsWhiteSpace(query[index])) index++;
            return index >= query.Length;
        }

        /// <summary>按分号拆开特殊参数，保留出现顺序。</summary>
        /// <param name="optionsText"><c>//</c> 后的参数文本。</param>
        /// <returns>每个键值一对。没有等号的片段以整段为键、值为空。</returns>
        public static IReadOnlyList<SqlSheetOption> SplitOptions(string optionsText)
        {
            string raw = (optionsText ?? string.Empty).Trim();
            List<SqlSheetOption> options = new List<SqlSheetOption>();
            if (raw.Length == 0) return options.AsReadOnly();

            string[] segments = raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string segment in segments)
            {
                string item = segment.Trim();
                if (item.Length == 0) continue;
                int equals = item.IndexOf('=');
                if (equals <= 0)
                {
                    options.Add(new SqlSheetOption(item, string.Empty));
                    continue;
                }

                string key = item.Substring(0, equals).Trim();
                string value = item.Substring(equals + 1).Trim();
                if (key.Length == 0)
                {
                    options.Add(new SqlSheetOption(item, string.Empty));
                    continue;
                }

                options.Add(new SqlSheetOption(key, value));
            }

            return options.AsReadOnly();
        }

        /// <summary>把列号转换成 Excel 列名。</summary>
        /// <param name="column">从 1 开始的列号。</param>
        /// <returns>例如 1 为 A，27 为 AA。</returns>
        public static string ColumnName(int column)
        {
            if (column < 1) return string.Empty;
            string name = string.Empty;
            while (column > 0)
            {
                column--;
                name = (char)('A' + column % 26) + name;
                column /= 26;
            }
            return name;
        }

        /// <summary>描述目标表重名。没有重名时返回空字符串。</summary>
        /// <param name="tasks">已解析的任务。比较表名时忽略大小写。</param>
        /// <returns>每个重名表一行说明，包含源列。没有重名时为空。</returns>
        public static string DescribeDuplicateTargets(IReadOnlyList<RefreshTaskDefinition> tasks)
        {
            if (tasks == null || tasks.Count == 0) return string.Empty;
            Dictionary<string, List<RefreshTaskDefinition>> groups =
                new Dictionary<string, List<RefreshTaskDefinition>>(StringComparer.OrdinalIgnoreCase);
            foreach (RefreshTaskDefinition task in tasks)
            {
                string key = task == null ? string.Empty : task.TargetSheetName ?? string.Empty;
                if (!groups.TryGetValue(key, out List<RefreshTaskDefinition> list))
                {
                    list = new List<RefreshTaskDefinition>();
                    groups.Add(key, list);
                }
                list.Add(task);
            }

            StringBuilder message = new StringBuilder();
            foreach (KeyValuePair<string, List<Models.RefreshTaskDefinition>> pair in groups)
            {
                if (pair.Value.Count < 2) continue;
                if (message.Length > 0) message.AppendLine();
                List<string> columns = new List<string>();
                foreach (RefreshTaskDefinition task in pair.Value)
                    columns.Add(ColumnName(task.SourceColumn));
                message.Append("目标表“");
                message.Append(pair.Value[0].TargetSheetName);
                message.Append("”重复，出现在列 ");
                message.Append(string.Join("、", columns));
                message.Append("。");
            }
            return message.ToString();
        }

        private static SqlSheetHeader Defaults(string sheetName, string optionsText, bool normalized)
        {
            return new SqlSheetHeader(sheetName, optionsText, "A1", 1, 1, true, string.Empty, normalized);
        }

        private static bool TryParseCell(string text, out int row, out int column, out string normalized)
        {
            row = 0;
            column = 0;
            normalized = null;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string value = text.Trim().Replace("$", string.Empty);
            int index = 0;
            while (index < value.Length && char.IsLetter(value[index])) index++;
            if (index == 0 || index >= value.Length) return false;
            string letters = value.Substring(0, index).ToUpperInvariant();
            string digits = value.Substring(index);
            if (!int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out row)) return false;
            if (row < 1 || row > MaxRow) return false;
            if (digits.Length > 1 && digits[0] == '0') return false;

            column = 0;
            foreach (char letter in letters)
            {
                if (letter < 'A' || letter > 'Z') return false;
                column = column * 26 + (letter - 'A' + 1);
                if (column > MaxColumn) return false;
            }

            if (column < 1) return false;
            normalized = letters + row.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static int SkipLeadingComments(string query, List<string> parts)
        {
            int index = 0;
            while (index < query.Length)
            {
                while (index < query.Length && char.IsWhiteSpace(query[index])) index++;
                if (index >= query.Length) break;

                if (StartsWith(query, index, "--"))
                {
                    index += 2;
                    int start = index;
                    while (index < query.Length && query[index] != '\n' && query[index] != '\r') index++;
                    AddComment(parts, query.Substring(start, index - start));
                    continue;
                }

                if (StartsWith(query, index, "/*"))
                {
                    index += 2;
                    int end = query.IndexOf("*/", index, StringComparison.Ordinal);
                    if (end < 0)
                    {
                        AddComment(parts, query.Substring(index));
                        return query.Length;
                    }

                    AddComment(parts, query.Substring(index, end - index));
                    index = end + 2;
                    continue;
                }

                break;
            }

            return index;
        }

        private static bool StartsWith(string text, int index, string token)
        {
            return index + token.Length <= text.Length
                && string.CompareOrdinal(text, index, token, 0, token.Length) == 0;
        }

        private static void AddComment(List<string> parts, string body)
        {
            if (parts == null) return;
            string trimmed = (body ?? string.Empty).Trim();
            if (trimmed.Length > 0) parts.Add(trimmed);
        }
    }

    /// <summary>SQL 页首行里的一项特殊参数。</summary>
    public sealed class SqlSheetOption
    {
        /// <summary>创建一项特殊参数。</summary>
        /// <param name="key">参数名。没有等号时为整段文本。</param>
        /// <param name="value">等号后的值。没有等号时为空。</param>
        public SqlSheetOption(string key, string value)
        {
            Key = key ?? string.Empty;
            Value = value ?? string.Empty;
        }

        /// <summary>获取参数名。</summary>
        public string Key { get; }

        /// <summary>获取参数值。</summary>
        public string Value { get; }
    }
}

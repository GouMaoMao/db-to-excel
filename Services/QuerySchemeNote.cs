using System;
using System.Globalization;
using System.Text;

namespace DB2Sheet.Services
{
    /// <summary>拼装并识别查询方案写在 SQL 开头的说明注释。</summary>
    /// <remarks>
    /// 悬停提示和保存进 SQL 的说明使用同一组标签：方案、连接、创建人、创建时间、修改人、修改时间。
    /// 说明只放在 SQL 最前面的一块块注释里。已有则整块替换，后面的注释和语句保持原样。
    /// 创建人和修改人由调用方决定是否覆盖；这里只负责显示，空名字显示为「未知」。
    /// </remarks>
    public static class QuerySchemeNote
    {
        private static readonly string[] Labels =
        {
            "方案：",
            "连接：",
            "创建人：",
            "创建时间：",
            "修改人：",
            "修改时间："
        };

        /// <summary>把 Windows 用户名和表格选项里的用户名拼成一个人名。</summary>
        /// <param name="windowsUser">Windows 登录名。为空时只用表格用户名。</param>
        /// <param name="spreadsheetUser">Excel 或 WPS 选项中的用户名。为空或与 Windows 名相同时不重复。</param>
        /// <returns>形如 <c>zhang / 张三</c>。两边都是空白时为空，由显示处写成「未知」。</returns>
        public static string FormatPerson(string windowsUser, string spreadsheetUser)
        {
            string windows = SingleLine(windowsUser);
            string spreadsheet = SingleLine(spreadsheetUser);
            if (windows.Length == 0) return spreadsheet;
            if (spreadsheet.Length == 0) return windows;
            if (string.Equals(windows, spreadsheet, StringComparison.OrdinalIgnoreCase)) return windows;
            return windows + " / " + spreadsheet;
        }

        /// <summary>按统一字段生成悬停用的六行文字。</summary>
        /// <param name="scheme">方案名称。</param>
        /// <param name="connection">连接方案名称。调用方在连接已删除时传入「连接已删除」。</param>
        /// <param name="createdBy">创建人。空表示没有记录。</param>
        /// <param name="createdUtc">创建时间（UTC）。默认值表示没有记录。</param>
        /// <param name="updatedBy">修改人。空表示没有记录。</param>
        /// <param name="updatedUtc">修改时间（UTC）。默认值表示没有记录。</param>
        /// <returns>六行标签文字，不含注释记号。</returns>
        public static string FormatLines(
            string scheme,
            string connection,
            string createdBy,
            DateTime createdUtc,
            string updatedBy,
            DateTime updatedUtc)
        {
            return string.Join(Environment.NewLine, BuildLines(scheme, connection, createdBy, createdUtc, updatedBy, updatedUtc));
        }

        /// <summary>在 SQL 开头写入或替换说明注释。</summary>
        /// <param name="sql">当前 SQL。可以已经带有一块说明。</param>
        /// <param name="scheme">方案名称。</param>
        /// <param name="connection">连接方案名称。</param>
        /// <param name="createdBy">创建人。空则显示「未知」，不表示本次要新建创建人。</param>
        /// <param name="createdUtc">创建时间（UTC）。</param>
        /// <param name="updatedBy">修改人。</param>
        /// <param name="updatedUtc">修改时间（UTC）。</param>
        /// <returns>说明位于最前面的 SQL。说明后面的原文保持不变。</returns>
        public static string Apply(
            string sql,
            string scheme,
            string connection,
            string createdBy,
            DateTime createdUtc,
            string updatedBy,
            DateTime updatedUtc)
        {
            string rest = WithoutLeadingNote(sql);
            string block = BuildBlock(scheme, connection, createdBy, createdUtc, updatedBy, updatedUtc);
            if (rest.Length == 0) return block;
            if (rest[0] != '\n' && rest[0] != '\r') return block + "\r\n" + rest;
            return block + rest;
        }

        /// <summary>判断去掉开头说明和空白后是否已经没有语句。</summary>
        /// <param name="sql">编辑器或选区中的 SQL。</param>
        /// <returns>只剩说明或空白时为 true。用户自己的注释仍算有内容。</returns>
        public static bool LacksStatement(string sql)
        {
            return string.IsNullOrWhiteSpace(WithoutLeadingNote(sql));
        }

        /// <summary>判断一块注释正文是不是查询方案说明。</summary>
        /// <param name="commentBody"><c>/*</c> 与 <c>*/</c> 之间的文本，不含记号本身。</param>
        /// <returns>正好是六行统一标签时为 true。摘要扫描用它跳过说明，不把它当成用户摘要。</returns>
        public static bool IsNoteBody(string commentBody)
        {
            string[] lines = SplitLines(commentBody);
            int count = 0;
            for (int index = 0; index < lines.Length; index++)
            {
                string line = lines[index].Trim();
                if (line.Length == 0) continue;
                if (count >= Labels.Length) return false;
                if (!line.StartsWith(Labels[count], StringComparison.Ordinal)) return false;
                count++;
            }

            return count == Labels.Length;
        }

        /// <summary>去掉位于 SQL 最前面的方案说明。不是说明时原样返回。</summary>
        /// <param name="sql">可能带说明的 SQL。</param>
        /// <returns>说明之后的文本，包含说明结束符后面的换行。前面的空白随说明一起去掉。</returns>
        private static string WithoutLeadingNote(string sql)
        {
            string value = sql ?? string.Empty;
            int index = 0;
            while (index < value.Length && char.IsWhiteSpace(value[index])) index++;
            if (index + 1 >= value.Length || value[index] != '/' || value[index + 1] != '*') return value;

            int bodyStart = index + 2;
            int end = value.IndexOf("*/", bodyStart, StringComparison.Ordinal);
            if (end < 0) return value;
            if (!IsNoteBody(value.Substring(bodyStart, end - bodyStart))) return value;
            return value.Substring(end + 2);
        }

        private static string BuildBlock(
            string scheme,
            string connection,
            string createdBy,
            DateTime createdUtc,
            string updatedBy,
            DateTime updatedUtc)
        {
            string[] lines = BuildLines(scheme, connection, createdBy, createdUtc, updatedBy, updatedUtc);
            StringBuilder block = new StringBuilder();
            block.Append("/*");
            block.Append("\r\n");
            for (int index = 0; index < lines.Length; index++)
            {
                block.Append(lines[index]);
                block.Append("\r\n");
            }

            block.Append("*/");
            return block.ToString();
        }

        private static string[] BuildLines(
            string scheme,
            string connection,
            string createdBy,
            DateTime createdUtc,
            string updatedBy,
            DateTime updatedUtc)
        {
            return new[]
            {
                Labels[0] + DisplayText(scheme, "未保存"),
                Labels[1] + DisplayText(connection, "未选择"),
                Labels[2] + DisplayPerson(createdBy),
                Labels[3] + DisplayTime(createdUtc),
                Labels[4] + DisplayPerson(updatedBy),
                Labels[5] + DisplayTime(updatedUtc)
            };
        }

        private static string DisplayPerson(string stored)
        {
            return DisplayText(stored, "未知");
        }

        private static string DisplayText(string value, string fallback)
        {
            string text = SingleLine(value);
            return text.Length == 0 ? fallback : text;
        }

        private static string DisplayTime(DateTime utc)
        {
            if (utc == default(DateTime)) return "未知";
            return utc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private static string SingleLine(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value.Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static string[] SplitLines(string value)
        {
            return (value ?? string.Empty).Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        }
    }
}

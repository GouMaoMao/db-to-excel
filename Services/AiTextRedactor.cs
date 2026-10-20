using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DB2Sheet.Services
{
    /// <summary>从将要离开本机的文字里去掉连接地址和密码样式片段。</summary>
    /// <remarks>替换是尽力而为，不能代替用户自己检查输入框。预览结果不应进入这里，调用方本来就不发送结果行。</remarks>
    public static class AiTextRedactor
    {
        private static readonly Regex JdbcUrl = new Regex(@"jdbc:[^\s'""]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex PasswordPair = new Regex(@"(?i)(password|pwd)\s*=\s*[^;\s]+", RegexOptions.Compiled);
        private static readonly Regex SecretPair = new Regex(
            @"(?i)(api[_-]?key|access[_-]?token|secret|authorization)\s*[:=]\s*\S+",
            RegexOptions.Compiled);
        private static readonly Regex Bearer = new Regex(@"(?i)bearer\s+\S+", RegexOptions.Compiled);

        /// <summary>返回去掉敏感片段后的文本，并说明去掉了哪几类。</summary>
        /// <param name="text">原始文字。为空时得到空字符串。</param>
        /// <param name="note">去掉的类别。什么都没去掉时为「无」。</param>
        /// <returns>可以发送和写入审计日志的文本。</returns>
        public static string Redact(string text, out string note)
        {
            List<string> notes = new List<string>();
            string current = text ?? string.Empty;
            current = Replace(current, JdbcUrl, "[已移除 JDBC 地址]", notes, "JDBC 地址");
            current = Replace(current, PasswordPair, "$1=***", notes, "密码");
            current = Replace(current, SecretPair, "$1=***", notes, "密钥或 Authorization");
            current = Replace(current, Bearer, "Bearer ***", notes, "Bearer");
            note = notes.Count == 0 ? "无" : string.Join("、", notes);
            return current;
        }

        /// <summary>拼出报错入口放进输入框的草稿。堆栈不进入草稿。</summary>
        /// <param name="sql">失败的那次 SQL。</param>
        /// <param name="errorMessage">异常消息，不是完整堆栈。</param>
        /// <returns>已经脱敏的草稿。错误消息超过 2000 字时截断。</returns>
        public static string BuildErrorDraft(string sql, string errorMessage)
        {
            string safeSql = Redact(sql, out string _);
            string safeError = Redact(errorMessage, out string _);
            if (safeError.Length > 2000)
                safeError = safeError.Substring(0, 2000) + "…";

            return AiPrompt.ErrorDraft(safeSql, safeError);
        }

        /// <summary>把提供程序标识说成发给模型的数据库类型。</summary>
        /// <param name="providerId">当前连接的提供程序标识。没有连接时为空。</param>
        /// <returns>类型说明。JDBC 会标明方言未知。</returns>
        public static string DescribeDialect(string providerId)
        {
            if (string.IsNullOrWhiteSpace(providerId))
                return "未选择连接，数据库类型未知";
            if (string.Equals(providerId, "jdbc", StringComparison.OrdinalIgnoreCase))
                return "jdbc，方言未知";
            return providerId;
        }

        private static string Replace(string text, Regex pattern, string replacement, List<string> notes, string label)
        {
            if (!pattern.IsMatch(text)) return text;
            if (!notes.Contains(label))
                notes.Add(label);
            return pattern.Replace(text, replacement);
        }
    }
}

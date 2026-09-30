using System;
using System.Collections.Generic;
using System.Globalization;
using DB2Sheet.Contracts;

namespace DB2Sheet.Providers
{
    /// <summary>把表注释没取到的原因写入文件日志。</summary>
    /// <remarks>
    /// 注释是附加信息。这里只记汇总，也不使用 Error，避免把目录缺口当成展开失败。
    /// </remarks>
    internal static class CommentDiagnostics
    {
        /// <summary>表或视图注释这一阶段的日志名称。</summary>
        public const string TableStage = "表注释";

        /// <summary>写入一条注释缺口日志。原因为空时不写。</summary>
        /// <param name="logger">文件日志。为空时直接返回。</param>
        /// <param name="providerId">数据源标识。</param>
        /// <param name="stage">表注释。</param>
        /// <param name="reason">给排查用的原因，不包含连接串、密码或 SQL。</param>
        /// <param name="count">这次涉及的对象或列数量。</param>
        /// <param name="failed">目录调用失败时为 true，记为 Warning；仅是空注释时为 false，记为 Information。</param>
        /// <param name="exception">目录调用抛出的异常。仅在 <paramref name="failed"/> 为 true 时写入。</param>
        public static void Write(
            ILogger logger,
            string providerId,
            string stage,
            string reason,
            int count,
            bool failed,
            Exception exception = null)
        {
            if (logger == null || string.IsNullOrWhiteSpace(reason)) return;
            Dictionary<string, string> properties = new Dictionary<string, string>
            {
                ["provider"] = providerId ?? string.Empty,
                ["stage"] = stage ?? string.Empty,
                ["reason"] = reason.Trim(),
                ["count"] = count.ToString(CultureInfo.InvariantCulture)
            };
            logger.Write(
                failed ? LogSeverity.Warning : LogSeverity.Information,
                "未能取得" + stage + "。",
                exception: failed ? exception : null,
                properties: properties);
        }

        /// <summary>去掉空白，并把没有实际含义的视图占位注释视为空。</summary>
        /// <param name="comment">目录返回的原始注释。</param>
        /// <returns>可展示的注释；没有时为空字符串。</returns>
        public static string Normalize(string comment)
        {
            if (string.IsNullOrWhiteSpace(comment)) return string.Empty;
            string trimmed = comment.Trim();
            if (string.Equals(trimmed, "VIEW", StringComparison.OrdinalIgnoreCase)) return string.Empty;
            return trimmed;
        }
    }
}

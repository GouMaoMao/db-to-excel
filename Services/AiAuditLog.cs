using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using DB2Sheet.Contracts;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>把将要离开本机的 AI 请求写成可审计的本地文件。</summary>
    /// <remarks>
    /// 文件是 <c>Logs\ai-audit-yyyyMMdd.log</c>。日常日志只留不含正文的一行索引。
    /// 呼叫模型前必须先写成功；写失败时调用方不得发送。删除对话不会删除已写下的记录。
    /// </remarks>
    public sealed class AiAuditLog
    {
        private const long MaximumFileBytes = 20L * 1024L * 1024L;
        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly ILogger _logger;

        /// <summary>创建审计日志。</summary>
        /// <param name="paths">日志目录来源。</param>
        /// <param name="logger">日常索引日志。正文不交给它。</param>
        public AiAuditLog(ApplicationPaths paths, ILogger logger)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>追加一条审计记录。磁盘失败时异常原样抛出。</summary>
        /// <param name="entry">要留下的内容。密钥不应出现在任何字段里。</param>
        /// <remarks>phase 为 pending 时不写日常索引。结果和本地动作各写一行索引。</remarks>
        public void Write(AiAuditEntry entry)
        {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            string text = Format(entry);
            lock (_syncRoot)
            {
                _paths.EnsureDirectories();
                string path = Path.Combine(_paths.LogsDirectory, "ai-audit-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
                RotateIfRequired(path);
                File.AppendAllText(path, text, Encoding.UTF8);
            }

            if (!string.Equals(entry.Phase, "pending", StringComparison.Ordinal))
                WriteIndex(entry);
        }

        /// <summary>计算请求体的 SHA256，十六进制小写。</summary>
        /// <param name="body">即将发送的 UTF-8 正文。</param>
        /// <returns>64 位十六进制。</returns>
        public static string Sha256(string body)
        {
            using (SHA256 algorithm = SHA256.Create())
            {
                byte[] hash = algorithm.ComputeHash(Encoding.UTF8.GetBytes(body ?? string.Empty));
                StringBuilder builder = new StringBuilder(hash.Length * 2);
                foreach (byte item in hash)
                    builder.Append(item.ToString("x2", CultureInfo.InvariantCulture));
                return builder.ToString();
            }
        }

        private void WriteIndex(AiAuditEntry entry)
        {
            _logger.Write(
                LogSeverity.Information,
                "AI 审计索引。",
                entry.OperationId,
                properties: new System.Collections.Generic.Dictionary<string, string>
                {
                    { "action", entry.Action ?? string.Empty },
                    { "conversationId", entry.ConversationId ?? string.Empty },
                    { "httpStatus", entry.HttpStatus.HasValue ? entry.HttpStatus.Value.ToString(CultureInfo.InvariantCulture) : string.Empty }
                });
        }

        private static void RotateIfRequired(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length < MaximumFileBytes)
                return;

            string archive = Path.Combine(
                Path.GetDirectoryName(path),
                Path.GetFileNameWithoutExtension(path) + "-" + DateTime.Now.ToString("HHmmssfff") + ".log");
            File.Move(path, archive);
        }

        private static string Format(AiAuditEntry entry)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("----- AI AUDIT BEGIN -----");
            Line(builder, "time", DateTimeOffset.Now.ToString("O", CultureInfo.InvariantCulture));
            Line(builder, "user", Environment.UserDomainName + "\\" + Environment.UserName);
            Line(builder, "action", entry.Action);
            Line(builder, "operationId", entry.OperationId);
            Line(builder, "phase", entry.Phase);
            Line(builder, "conversationId", entry.ConversationId);
            Line(builder, "conversationTitle", entry.ConversationTitle);
            Line(builder, "model", entry.ModelDisplayName);
            Line(builder, "protocol", entry.Protocol);
            Line(builder, "endpoint", entry.Endpoint);
            Line(builder, "skillName", entry.SkillName);
            Block(builder, "skillBody", entry.SkillBody);
            Block(builder, "system", entry.SystemText);
            Block(builder, "history", entry.HistoryText);
            Block(builder, "user", entry.UserText);
            Line(builder, "tabName", entry.TabName);
            Line(builder, "hasTab", entry.HasTab.HasValue ? (entry.HasTab.Value ? "true" : "false") : string.Empty);
            Block(builder, "tabSql", entry.TabSql);
            Line(builder, "dialect", entry.Dialect);
            Line(builder, "redaction", entry.Redaction);
            Block(builder, "context", entry.ContextText);
            Line(builder, "resultRows", entry.IncludesPreview == true ? "true" : "false");
            Line(builder, "connectionSecret", "false");
            Line(builder, "catalog", "false");
            Line(builder, "stackTrace", "false");
            Line(builder, "requestSha256", entry.RequestSha256);
            Line(builder, "httpStatus", entry.HttpStatus.HasValue ? entry.HttpStatus.Value.ToString(CultureInfo.InvariantCulture) : string.Empty);
            Line(builder, "elapsedMs", entry.ElapsedMs.HasValue ? entry.ElapsedMs.Value.ToString(CultureInfo.InvariantCulture) : string.Empty);
            Block(builder, "assistant", entry.AssistantText);
            Block(builder, "detail", entry.Detail);
            builder.AppendLine("----- AI AUDIT END -----");
            return builder.ToString();
        }

        private static void Line(StringBuilder builder, string name, string value)
        {
            builder.Append(name).Append('=').Append(value ?? string.Empty).AppendLine();
        }

        private static void Block(StringBuilder builder, string name, string value)
        {
            builder.Append(name).AppendLine(".begin");
            builder.AppendLine(value ?? string.Empty);
            builder.Append(name).AppendLine(".end");
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using DB2Sheet.Contracts;

namespace DB2Sheet.Infrastructure
{
    /// <summary>将结构化日志以 UTF-8 文本追加到按日期命名的本地文件。</summary>
    /// <remarks>
    /// 写入过程使用锁保证单进程线程安全；日志超过 5 MB 时归档。
    /// 日志失败会被吞掉，避免影响插件主流程。常见敏感键和值会被脱敏，但调用方仍不应主动记录密码。
    /// </remarks>
    public sealed class FileLogger : ILogger
    {
        private const long MaximumFileBytes = 5L * 1024L * 1024L;
        private static readonly string[] SensitiveKeys =
        {
            "password", "pwd", "token", "secret", "authorization", "connectionstring"
        };

        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;

        /// <summary>创建文件日志记录器。</summary>
        /// <param name="paths">日志目录来源。</param>
        public FileLogger(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        }

        /// <inheritdoc/>
        public void Write(
            LogSeverity severity,
            string message,
            string operationId = null,
            Exception exception = null,
            IReadOnlyDictionary<string, string> properties = null)
        {
            try
            {
                lock (_syncRoot)
                {
                    _paths.EnsureDirectories();
                    string path = GetCurrentLogPath();
                    RotateIfRequired(path);
                    File.AppendAllText(path, Format(severity, message, operationId, exception, properties), Encoding.UTF8);
                }
            }
            catch (Exception)
            {
                // Logging must never interrupt the add-in operation.
            }
        }

        private string GetCurrentLogPath()
        {
            return Path.Combine(_paths.LogsDirectory, "db2sheet-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
        }

        private static void RotateIfRequired(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length < MaximumFileBytes)
            {
                return;
            }

            string archive = Path.Combine(
                Path.GetDirectoryName(path),
                Path.GetFileNameWithoutExtension(path) + "-" + DateTime.Now.ToString("HHmmssfff") + ".log");
            File.Move(path, archive);
        }

        private static string Format(
            LogSeverity severity,
            string message,
            string operationId,
            Exception exception,
            IReadOnlyDictionary<string, string> properties)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(DateTimeOffset.Now.ToString("O", CultureInfo.InvariantCulture));
            builder.Append(" [").Append(severity).Append("]");
            if (!string.IsNullOrWhiteSpace(operationId)) builder.Append(" [operation:").Append(operationId).Append("]");
            builder.Append(' ').Append(Sanitize(message));

            if (properties != null)
            {
                foreach (KeyValuePair<string, string> pair in properties.OrderBy(item => item.Key))
                {
                    builder.Append(" | ").Append(pair.Key).Append('=');
                    builder.Append(IsSensitive(pair.Key) ? "***" : Sanitize(pair.Value));
                }
            }

            if (exception != null)
            {
                builder.Append(" | exception=").Append(exception.GetType().FullName);
                builder.Append(" | detail=").Append(Sanitize(exception.ToString()));
            }

            return builder.AppendLine().ToString();
        }

        private static bool IsSensitive(string key)
        {
            string normalized = key ?? string.Empty;
            return SensitiveKeys.Any(item => normalized.IndexOf(item, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            string sanitized = value.Replace("\r", " ").Replace("\n", " ");
            foreach (string marker in new[] { "Password=", "Pwd=", "Token=", "Authorization=" })
            {
                int start = sanitized.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                while (start >= 0)
                {
                    int valueStart = start + marker.Length;
                    int end = sanitized.IndexOf(';', valueStart);
                    if (end < 0) end = sanitized.Length;
                    sanitized = sanitized.Substring(0, valueStart) + "***" + sanitized.Substring(end);
                    start = sanitized.IndexOf(marker, valueStart + 3, StringComparison.OrdinalIgnoreCase);
                }
            }

            return sanitized;
        }
    }
}

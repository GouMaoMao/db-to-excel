using System;
using System.Collections.Generic;

namespace DB2Sheet.Contracts
{
    /// <summary>定义日志事件的严重程度。</summary>
    public enum LogSeverity
    {
        Debug,
        Information,
        Warning,
        Error
    }

    /// <summary>定义结构化应用日志写入接口。</summary>
    public interface ILogger
    {
        /// <summary>写入一条日志记录。</summary>
        /// <param name="severity">日志严重程度。</param>
        /// <param name="message">面向开发和排障的消息。</param>
        /// <param name="operationId">可选的操作关联标识。</param>
        /// <param name="exception">可选的异常详情。</param>
        /// <param name="properties">可选的结构化附加属性。</param>
        /// <remarks>敏感连接参数和密码不应放入消息或附加属性。</remarks>
        void Write(
            LogSeverity severity,
            string message,
            string operationId = null,
            Exception exception = null,
            IReadOnlyDictionary<string, string> properties = null);
    }
}

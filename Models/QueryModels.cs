using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace DB2Sheet.Models
{
    /// <summary>描述查询执行的业务目的，以便提供程序应用不同限制。</summary>
    public enum ExecutionPurpose
    {
        Preview,
        Export,
        ConnectionTest
    }

    /// <summary>封装一次数据源查询所需的连接、SQL 和执行限制。</summary>
    public sealed class DataSourceRequest
    {
        /// <summary>创建数据源请求。</summary>
        /// <param name="connection">不可变连接快照。</param>
        /// <param name="queryText">只读 SQL 文本。</param>
        /// <param name="purpose">执行目的。</param>
        /// <param name="rowLimit">允许返回的最大行数。</param>
        /// <param name="blockSize">每次读取的行数。</param>
        /// <param name="timeoutSeconds">数据库命令超时秒数。</param>
        public DataSourceRequest(
            ConnectionProfileSnapshot connection,
            string queryText,
            ExecutionPurpose purpose,
            int rowLimit,
            int blockSize,
            int timeoutSeconds = 300)
        {
            Connection = connection ?? throw new ArgumentNullException(nameof(connection));
            QueryText = queryText ?? string.Empty;
            Purpose = purpose;
            RowLimit = rowLimit;
            BlockSize = blockSize;
            TimeoutSeconds = timeoutSeconds;
        }

        /// <summary>获取连接快照。</summary>
        public ConnectionProfileSnapshot Connection { get; }
        /// <summary>获取 SQL 文本。</summary>
        public string QueryText { get; }
        /// <summary>获取执行目的。</summary>
        public ExecutionPurpose Purpose { get; }
        /// <summary>获取最大结果行数。</summary>
        public int RowLimit { get; }
        /// <summary>获取结果读取块大小。</summary>
        public int BlockSize { get; }
        /// <summary>获取数据库命令超时秒数。</summary>
        public int TimeoutSeconds { get; }
    }

    /// <summary>表示用户保存并可重复执行的 SQL 查询方案。</summary>
    [DataContract]
    public sealed class QueryProfile
    {
        /// <summary>创建具有新标识、空内容和当前 UTC 时间的查询方案。</summary>
        public QueryProfile()
        {
            Id = Guid.NewGuid().ToString("N");
            Name = string.Empty;
            ProviderId = string.Empty;
            ConnectionProfileId = string.Empty;
            QueryText = string.Empty;
            TargetSheetName = string.Empty;
            Description = string.Empty;
            CreatedUtc = DateTime.UtcNow;
            UpdatedUtc = DateTime.UtcNow;
            ProviderOptions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>获取或设置方案唯一标识。</summary>
        [DataMember(Order = 1)] public string Id { get; set; }
        /// <summary>获取或设置用户可见名称。</summary>
        [DataMember(Order = 2)] public string Name { get; set; }
        /// <summary>获取或设置数据源提供程序标识。</summary>
        [DataMember(Order = 3)] public string ProviderId { get; set; }
        /// <summary>获取或设置关联连接方案标识。</summary>
        [DataMember(Order = 4)] public string ConnectionProfileId { get; set; }
        /// <summary>获取或设置 SQL 文本。</summary>
        [DataMember(Order = 5)] public string QueryText { get; set; }
        /// <summary>获取或设置默认目标工作表名称。</summary>
        [DataMember(Order = 6)] public string TargetSheetName { get; set; }
        /// <summary>获取或设置方案说明。</summary>
        [DataMember(Order = 7)] public string Description { get; set; }
        /// <summary>获取或设置创建时间（UTC）。</summary>
        [DataMember(Order = 8)] public DateTime CreatedUtc { get; set; }
        /// <summary>获取或设置最后更新时间（UTC）。</summary>
        [DataMember(Order = 9)] public DateTime UpdatedUtc { get; set; }
        /// <summary>获取或设置提供程序专用选项。</summary>
        [DataMember(Order = 10)] public Dictionary<string, string> ProviderOptions { get; set; }
    }

    /// <summary>表示从 SQL Sheet 单元格解析出的一个批量刷新任务。</summary>
    public sealed class RefreshTaskDefinition
    {
        /// <summary>创建使用默认写入选项的刷新任务定义。</summary>
        /// <param name="id">任务标识；为空时自动生成。</param>
        /// <param name="targetSheetName">目标工作表名称。</param>
        /// <param name="queryText">要执行的 SQL。</param>
        /// <param name="sourceColumn">任务在 SQL Sheet 中的源列号。</param>
        public RefreshTaskDefinition(string id, string targetSheetName, string queryText, int sourceColumn)
            : this(id, targetSheetName, queryText, sourceColumn, string.Empty, "A1", 1, 1, true, null, string.Empty, false)
        {
        }

        /// <summary>创建带首行特殊参数的刷新任务定义。</summary>
        /// <param name="id">任务标识；为空时自动生成。</param>
        /// <param name="targetSheetName">目标工作表名称，不含 <c>//</c> 后的参数。</param>
        /// <param name="queryText">要执行的 SQL，包含前置注释。</param>
        /// <param name="sourceColumn">任务在 SQL Sheet 中的源列号。</param>
        /// <param name="optionsText"><c>//</c> 后面的特殊参数原文。没有参数时为空。</param>
        /// <param name="startCell">结果表左上角单元格，例如 A1。</param>
        /// <param name="startRow">起始单元格的行号，从 1 开始。</param>
        /// <param name="startColumn">起始单元格的列号，从 1 开始。</param>
        /// <param name="clearExtraColumns">是否清空结果没有覆盖到的原数据列。</param>
        /// <param name="optionsError">已知参数无法解析时的说明。没有错误时为空。</param>
        /// <param name="summary">SQL 命令前的注释摘要。没有注释时为空。</param>
        /// <param name="optionsNormalized">特殊参数在解析时被规范成可写回的文本时为 true。</param>
        /// <param name="sqlCommandRow">SQL 命令在 SQL 页中的起始行号，从 1 开始。整段都是注释时为第一处正文行。</param>
        public RefreshTaskDefinition(
            string id,
            string targetSheetName,
            string queryText,
            int sourceColumn,
            string optionsText,
            string startCell,
            int startRow,
            int startColumn,
            bool clearExtraColumns,
            string optionsError,
            string summary,
            bool optionsNormalized,
            int sqlCommandRow = 0)
        {
            Id = id ?? Guid.NewGuid().ToString("N");
            TargetSheetName = targetSheetName ?? string.Empty;
            QueryText = queryText ?? string.Empty;
            SourceColumn = sourceColumn;
            OptionsText = optionsText ?? string.Empty;
            StartCell = string.IsNullOrWhiteSpace(startCell) ? "A1" : startCell;
            StartRow = startRow < 1 ? 1 : startRow;
            StartColumnIndex = startColumn < 1 ? 1 : startColumn;
            ClearExtraColumns = clearExtraColumns;
            OptionsError = optionsError ?? string.Empty;
            Summary = summary ?? string.Empty;
            OptionsNormalized = optionsNormalized;
            SqlCommandRow = sqlCommandRow < 1 ? 0 : sqlCommandRow;
        }

        /// <summary>获取任务标识。</summary>
        public string Id { get; }
        /// <summary>获取目标工作表名称。</summary>
        public string TargetSheetName { get; }
        /// <summary>获取 SQL 文本。</summary>
        public string QueryText { get; }
        /// <summary>获取任务所在的源列号。</summary>
        public int SourceColumn { get; }
        /// <summary>获取首行 <c>//</c> 后面的特殊参数文本。没有参数时为空。</summary>
        public string OptionsText { get; }
        /// <summary>获取结果表左上角单元格。</summary>
        public string StartCell { get; }
        /// <summary>获取起始单元格行号，从 1 开始。</summary>
        public int StartRow { get; }
        /// <summary>获取起始单元格列号，从 1 开始。</summary>
        public int StartColumnIndex { get; }
        /// <summary>获取是否清空结果没有覆盖到的原数据列。</summary>
        public bool ClearExtraColumns { get; }
        /// <summary>获取已知特殊参数的解析错误。没有错误时为空。</summary>
        public string OptionsError { get; }
        /// <summary>获取 SQL 命令前的注释摘要。</summary>
        public string Summary { get; }
        /// <summary>获取特殊参数是否在解析时被规范，需要写回 SQL 页。</summary>
        public bool OptionsNormalized { get; }
        /// <summary>获取 SQL 命令在 SQL 页中的起始行号。从 1 开始；没有正文时为 0。</summary>
        public int SqlCommandRow { get; }
    }
}

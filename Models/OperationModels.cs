using System;
using DB2Sheet.Contracts;

namespace DB2Sheet.Models
{
    /// <summary>定义长时间操作当前所处的阶段。</summary>
    public enum OperationStage
    {
        Preparing,
        Connecting,
        Executing,
        Reading,
        Writing,
        Cancelling,
        Completed,
        Failed,
        Cancelled
    }

    /// <summary>表示后台操作向进度窗体发送的一次状态快照。</summary>
    /// <remarks>
    /// 该对象是可变传输模型；报告后调用方不应继续修改同一实例。
    /// <see cref="Severity"/> 的默认值是 <see cref="LogSeverity.Information"/>，未标注的进度行在默认设置下仍会写入进度日志列表。
    /// 枚举的数值顺序就是过滤顺序，进度窗口只追加不低于所选级别的行。
    /// </remarks>
    public sealed class OperationProgress
    {
        /// <summary>创建一条默认为信息级别的进度快照。</summary>
        public OperationProgress()
        {
            Severity = LogSeverity.Information;
        }

        /// <summary>获取或设置用于关联同一次操作的标识。</summary>
        public string OperationId { get; set; }
        /// <summary>获取或设置批处理中当前任务的标识。</summary>
        public string TaskId { get; set; }
        /// <summary>获取或设置当前执行阶段。</summary>
        public OperationStage Stage { get; set; }
        /// <summary>获取或设置面向用户的状态消息。</summary>
        public string Message { get; set; }
        /// <summary>获取或设置该条进度在日志列表中的显示级别。</summary>
        /// <remarks>只影响进度窗口的日志列表。文件日志不读取此值。</remarks>
        public LogSeverity Severity { get; set; }
        /// <summary>获取或设置当前连接方案名称。</summary>
        public string ConnectionName { get; set; }
        /// <summary>获取或设置当前工作簿名称。</summary>
        public string WorkbookName { get; set; }
        /// <summary>获取或设置目标工作表名称。</summary>
        public string TargetSheetName { get; set; }
        /// <summary>获取或设置批次任务总数。</summary>
        public int TotalTasks { get; set; }
        /// <summary>获取或设置当前任务序号。</summary>
        public int CurrentTask { get; set; }
        /// <summary>获取或设置已成功任务数。</summary>
        public int SucceededTasks { get; set; }
        /// <summary>获取或设置已失败任务数。</summary>
        public int FailedTasks { get; set; }
        /// <summary>获取或设置已跳过任务数。</summary>
        public int SkippedTasks { get; set; }
        /// <summary>获取或设置累计读取行数。</summary>
        public long RowsRead { get; set; }
        /// <summary>获取或设置累计写入行数。</summary>
        public long RowsWritten { get; set; }
        /// <summary>获取或设置 0 到 100 的完成百分比；未知时为 <see langword="null"/>。</summary>
        public int? Percent { get; set; }
        /// <summary>获取或设置是否显示无法估算完成比例的进度。</summary>
        public bool IsIndeterminate { get; set; }
        /// <summary>获取或设置当前结果是否因行数上限被截断。</summary>
        public bool IsTruncated { get; set; }
        /// <summary>获取或设置操作已耗用时间。</summary>
        public TimeSpan Elapsed { get; set; }
    }
}

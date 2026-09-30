using System;
using System.Collections.Generic;

namespace DB2Sheet.Models
{
    /// <summary>表示已从数据结果流完整读取到内存、可供 Excel 写入的查询结果。</summary>
    /// <remarks>数据按多个二维数组保存以避免构造单个超大数组，但整体仍占用内存。</remarks>
    public sealed class BufferedQueryResult
    {
        /// <summary>创建缓冲查询结果。</summary>
        /// <param name="columns">结果列定义。</param>
        /// <param name="blocks">按读取顺序排列的二维数据块。</param>
        /// <param name="rowCount">全部数据块的总行数。</param>
        /// <param name="isTruncated">是否因行数限制被截断。</param>
        public BufferedQueryResult(
            IReadOnlyList<ResultColumn> columns,
            IReadOnlyList<object[,]> blocks,
            long rowCount,
            bool isTruncated)
        {
            Columns = columns ?? throw new ArgumentNullException(nameof(columns));
            Blocks = blocks ?? throw new ArgumentNullException(nameof(blocks));
            RowCount = rowCount;
            IsTruncated = isTruncated;
        }

        /// <summary>获取结果列定义。</summary>
        public IReadOnlyList<ResultColumn> Columns { get; }
        /// <summary>获取按顺序排列的数据块。</summary>
        public IReadOnlyList<object[,]> Blocks { get; }
        /// <summary>获取总数据行数。</summary>
        public long RowCount { get; }
        /// <summary>获取结果是否被截断。</summary>
        public bool IsTruncated { get; }
    }

    /// <summary>表示单个批量刷新任务的最终执行结果。</summary>
    public sealed class RefreshTaskResult
    {
        /// <summary>创建任务结果。</summary>
        /// <param name="task">对应的任务定义。</param>
        /// <param name="succeeded">任务是否成功。</param>
        /// <param name="skipped">任务是否因批次策略而跳过。</param>
        /// <param name="rowsWritten">写入 Excel 的数据行数。</param>
        /// <param name="isTruncated">查询结果是否被截断。</param>
        /// <param name="message">状态或错误消息。</param>
        public RefreshTaskResult(
            RefreshTaskDefinition task,
            bool succeeded,
            bool skipped,
            long rowsWritten,
            bool isTruncated,
            string message)
        {
            Task = task ?? throw new ArgumentNullException(nameof(task));
            Succeeded = succeeded;
            Skipped = skipped;
            RowsWritten = rowsWritten;
            IsTruncated = isTruncated;
            Message = message ?? string.Empty;
        }

        /// <summary>获取对应的任务定义。</summary>
        public RefreshTaskDefinition Task { get; }
        /// <summary>获取任务是否成功。</summary>
        public bool Succeeded { get; }
        /// <summary>获取任务是否被跳过。</summary>
        public bool Skipped { get; }
        /// <summary>获取写入的数据行数。</summary>
        public long RowsWritten { get; }
        /// <summary>获取查询结果是否被截断。</summary>
        public bool IsTruncated { get; }
        /// <summary>获取状态或错误消息。</summary>
        public string Message { get; }
    }

    /// <summary>汇总一次批量刷新中所有任务的结果。</summary>
    public sealed class BatchRefreshResult
    {
        /// <summary>创建批量刷新结果。</summary>
        /// <param name="tasks">各任务执行结果。</param>
        public BatchRefreshResult(IReadOnlyList<RefreshTaskResult> tasks)
        {
            Tasks = tasks ?? throw new ArgumentNullException(nameof(tasks));
        }

        /// <summary>获取各任务执行结果。</summary>
        public IReadOnlyList<RefreshTaskResult> Tasks { get; }
    }

    /// <summary>描述一次查询结果写入工作表时的锚点和清空范围。</summary>
    /// <remarks>
    /// 交互式写入使用整表已用区域清空，并从 A1 开始。
    /// 批量刷新按起始单元格锚定，只清结果所占列；是否再清右侧原数据列由任务参数决定。
    /// </remarks>
    public sealed class SheetWriteOptions
    {
        private SheetWriteOptions(int anchorRow, int anchorColumn, bool clearEntireUsedRange, bool clearExtraColumns)
        {
            AnchorRow = anchorRow;
            AnchorColumn = anchorColumn;
            ClearEntireUsedRange = clearEntireUsedRange;
            ClearExtraColumns = clearExtraColumns;
        }

        /// <summary>创建从 A1 写入并清空整个已用区域的选项。</summary>
        /// <returns>交互式写入使用的选项。</returns>
        public static SheetWriteOptions EntireUsedRange()
        {
            return new SheetWriteOptions(1, 1, true, true);
        }

        /// <summary>创建按锚点写入的选项。</summary>
        /// <param name="anchorRow">结果表左上角行号，从 1 开始。</param>
        /// <param name="anchorColumn">结果表左上角列号，从 1 开始。</param>
        /// <param name="clearExtraColumns">是否清空锚点右侧、原数据区域内结果没有覆盖的列。</param>
        /// <returns>批量刷新使用的写入选项。</returns>
        public static SheetWriteOptions Anchored(int anchorRow, int anchorColumn, bool clearExtraColumns)
        {
            if (anchorRow < 1) throw new ArgumentOutOfRangeException(nameof(anchorRow));
            if (anchorColumn < 1) throw new ArgumentOutOfRangeException(nameof(anchorColumn));
            return new SheetWriteOptions(anchorRow, anchorColumn, false, clearExtraColumns);
        }

        /// <summary>获取结果表左上角行号。</summary>
        public int AnchorRow { get; }
        /// <summary>获取结果表左上角列号。</summary>
        public int AnchorColumn { get; }
        /// <summary>获取是否在写入前清空目标表整个已用区域。</summary>
        public bool ClearEntireUsedRange { get; }
        /// <summary>获取是否清空结果列右侧的原数据列。整表清空时忽略此项。</summary>
        public bool ClearExtraColumns { get; }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.Excel
{
    /// <summary>使用 Excel COM Interop 将缓冲查询结果批量写入指定工作表。</summary>
    /// <remarks>
    /// 交互式写入只清空目标工作表已用区域的值，并从 A1 开始。原先单元格上的数值格式、字体、对齐和填充保留；本次多出来的行不继承这些格式。
    /// 批量刷新按起始单元格锚定：结果所占列从锚点清掉值到原数据末行，右侧多余列仅在参数要求时清空值。
    /// 清空范围内的边框会去掉，避免结果变短后留下空网格。写完后只在标题和数据组成的区域重新画灰色细网格线。
    /// 标题只占一行字段编码，编码行加粗，底色表示成功或截断。这一步只改标题行的加粗和底色。成功和截断的底色来自参数设置，与批量刷新状态圆点相同。数据从标题的下一行开始。
    /// 每写完一块向进度接收器报告已写行数和块序号；总块数来自已缓冲结果，进度条按该比例填充。
    /// 批量刷新会在写入前激活目标表并滚到锚点，写入期间关闭 Excel 重画，结束时恢复。交互式写入不跳表，也不改重画。
    /// 文件日志只在写入完成或失败时各记一条汇总。
    /// 应在 Excel UI 线程调用，并避免在方法外继续使用本类释放的 COM Range 和 Worksheet 引用。
    /// </remarks>
    public sealed class ExcelResultWriter : IExcelResultWriter
    {
        private const int HeaderRowCount = 1;

        /// <summary>成功写入时标题行和状态圆点的默认浅绿色。黑字在该底色上可以辨认。</summary>
        public static readonly Color SuccessColor = Color.FromArgb(101, 219, 91);

        /// <summary>结果被截断时标题行和状态圆点的默认橙色。</summary>
        public static readonly Color TruncatedColor = Color.FromArgb(255, 184, 98);

        /// <summary>已勾选、尚未刷新时状态圆点的默认灰色。不用于标题行。</summary>
        public static readonly Color PendingColor = Color.FromArgb(162, 162, 162);

        /// <summary>失败或参数、重名有误时状态圆点的默认砖红色。不用于标题行。</summary>
        public static readonly Color ErrorColor = Color.FromArgb(217, 57, 57);

        private readonly ILogger _logger;
        private readonly ISettingsStore _settings;

        /// <summary>创建 Excel 结果写入器。</summary>
        /// <param name="logger">写入阶段的文件日志。不按进度日志级别过滤。</param>
        /// <param name="settings">提供成功色和截断色。每次写入标题行时读取当时的值。</param>
        public ExcelResultWriter(ILogger logger, ISettingsStore settings)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <inheritdoc/>
        public long WriteResult(
            ExcelInterop.Workbook workbook,
            string targetSheetName,
            BufferedQueryResult result,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken,
            string operationId)
        {
            return WriteResult(
                workbook,
                targetSheetName,
                result,
                SheetWriteOptions.EntireUsedRange(),
                progress,
                cancellationToken,
                operationId);
        }

        /// <inheritdoc/>
        public long WriteResult(
            ExcelInterop.Workbook workbook,
            string targetSheetName,
            BufferedQueryResult result,
            SheetWriteOptions options,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken,
            string operationId)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(targetSheetName)) throw new ArgumentException("目标 Sheet 名称不能为空。", nameof(targetSheetName));
            if (result == null) throw new ArgumentNullException(nameof(result));
            if (options == null) throw new ArgumentNullException(nameof(options));
            ValidateSheetName(targetSheetName);

            string workbookName = workbook.Name;
            ExcelInterop.Worksheet worksheet = GetOrCreateWorksheet(workbook, targetSheetName);
            Stopwatch writeWatch = Stopwatch.StartNew();
            long rowsWritten = 0;
            int totalBlocks = result.Blocks.Count;
            ExcelInterop.Application application = null;
            bool restoreScreenUpdating = false;
            bool previousScreenUpdating = true;
            try
            {
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Preparing,
                    WorkbookName = workbookName,
                    TargetSheetName = targetSheetName,
                    Message = string.Format(
                        "目标区域已确认，准备写入 {0:N0} 行、{1:N0} 列。",
                        result.RowCount,
                        result.Columns.Count),
                    RowsRead = result.RowCount,
                    IsIndeterminate = true
                });
                if (options.RevealTarget)
                {
                    RevealTarget(worksheet, options.AnchorRow, options.AnchorColumn);
                    try
                    {
                        application = worksheet.Application;
                        previousScreenUpdating = application.ScreenUpdating;
                        restoreScreenUpdating = true;
                        application.ScreenUpdating = false;
                    }
                    catch (COMException)
                    {
                        // 关不掉重画时仍继续写入。能改的话，finally 会写回原值。
                    }
                    catch (InvalidComObjectException)
                    {
                    }
                }

                cancellationToken.ThrowIfCancellationRequested();
                int columnCount = result.Columns.Count;
                ClearBeforeWrite(worksheet, options, columnCount);
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Writing,
                    WorkbookName = workbookName,
                    TargetSheetName = targetSheetName,
                    Message = "目标区域已清理，正在写入字段标题…",
                    RowsRead = result.RowCount,
                    Percent = 0,
                    IsIndeterminate = false
                });

                if (columnCount == 0)
                {
                    ReportWriteCompleted(progress, workbookName, targetSheetName, result, 0, writeWatch.Elapsed);
                    LogWriteCompleted(operationId, workbookName, targetSheetName, result, rowsWritten, totalBlocks, writeWatch.Elapsed);
                    return 0;
                }
                WriteHeaders(
                    worksheet,
                    result,
                    options.AnchorRow,
                    options.AnchorColumn,
                    result.IsTruncated
                        ? _settings.Get(CoreSettings.ResultTruncatedColor)
                        : _settings.Get(CoreSettings.ResultSuccessColor));

                int targetRow = options.AnchorRow + HeaderRowCount;
                int writtenBlocks = 0;
                foreach (object[,] block in result.Blocks)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int rowCount = block.GetLength(0);
                    if (rowCount == 0) continue;
                    WriteBlock(worksheet, targetRow, options.AnchorColumn, columnCount, block, rowCount);
                    targetRow += rowCount;
                    rowsWritten += rowCount;
                    writtenBlocks++;
                    progress?.Report(new OperationProgress
                    {
                        Stage = OperationStage.Writing,
                        TargetSheetName = targetSheetName,
                        Message = string.Format(
                            "正在写入 {0:N0} 行（{1}/{2}）",
                            rowsWritten,
                            writtenBlocks,
                            totalBlocks),
                        RowsRead = result.RowCount,
                        RowsWritten = rowsWritten,
                        Percent = totalBlocks == 0 ? 0 : (int)(writtenBlocks * 100L / totalBlocks),
                        IsTruncated = result.IsTruncated,
                        IsIndeterminate = false
                    });
                }

                ApplyTableBorder(worksheet, options.AnchorRow, options.AnchorColumn, targetRow - 1, columnCount);
                ReportWriteCompleted(progress, workbookName, targetSheetName, result, rowsWritten, writeWatch.Elapsed);
                LogWriteCompleted(operationId, workbookName, targetSheetName, result, rowsWritten, totalBlocks, writeWatch.Elapsed);
                return rowsWritten;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                _logger.Write(
                    LogSeverity.Error,
                    "工作表写入失败。",
                    operationId,
                    exception,
                    WriteProperties(workbookName, targetSheetName, result, rowsWritten, totalBlocks, writeWatch.Elapsed, includeRate: false));
                throw;
            }
            finally
            {
                RestoreScreenUpdating(application, restoreScreenUpdating, previousScreenUpdating);
                Marshal.FinalReleaseComObject(worksheet);
            }
        }

        /// <summary>向进度窗口报告工作表数据和边框均已写完。</summary>
        /// <param name="progress">进度接收器，可为空。</param>
        /// <param name="workbookName">目标工作簿名。</param>
        /// <param name="targetSheetName">目标工作表名。</param>
        /// <param name="result">本次缓冲结果。</param>
        /// <param name="rowsWritten">已经写入的数据行数。</param>
        /// <param name="elapsed">写入阶段耗时。</param>
        private static void ReportWriteCompleted(
            IProgress<OperationProgress> progress,
            string workbookName,
            string targetSheetName,
            BufferedQueryResult result,
            long rowsWritten,
            TimeSpan elapsed)
        {
            progress?.Report(new OperationProgress
            {
                Stage = OperationStage.Writing,
                WorkbookName = workbookName,
                TargetSheetName = targetSheetName,
                Message = string.Format(
                    "数据与边框写入完成：{0:N0} 行、{1:N0} 列，耗时 {2:N1} 秒{3}",
                    rowsWritten,
                    result.Columns.Count,
                    elapsed.TotalSeconds,
                    result.IsTruncated ? "，结果已截断。" : "。"),
                RowsRead = result.RowCount,
                RowsWritten = rowsWritten,
                Percent = 100,
                IsTruncated = result.IsTruncated,
                Severity = result.IsTruncated ? LogSeverity.Warning : LogSeverity.Information,
                Elapsed = elapsed,
                IsIndeterminate = false
            });
        }

        /// <summary>激活目标表，并把锚点滚到当前窗口视口的左上角。</summary>
        /// <remarks>不选择单元格，也不调用 Goto。隐藏表或没有窗口时吞掉 COM 异常，调用方继续写入。</remarks>
        private static void RevealTarget(ExcelInterop.Worksheet worksheet, int anchorRow, int anchorColumn)
        {
            ExcelInterop.Window window = null;
            try
            {
                worksheet.Activate();
                window = worksheet.Application.ActiveWindow;
                if (window == null) return;
                window.ScrollRow = anchorRow;
                window.ScrollColumn = anchorColumn;
            }
            catch (COMException)
            {
                // 隐藏表或没有窗口时继续写入。
            }
            catch (InvalidComObjectException)
            {
            }
            finally
            {
                if (window != null) Marshal.FinalReleaseComObject(window);
            }
        }

        /// <summary>把 Excel 重画恢复成进入写入前的值。</summary>
        /// <remarks>不释放 Application。恢复失败时吞掉，避免盖住原来的写入异常。</remarks>
        private static void RestoreScreenUpdating(ExcelInterop.Application application, bool restore, bool previous)
        {
            if (!restore || application == null) return;
            try
            {
                application.ScreenUpdating = previous;
            }
            catch (COMException)
            {
                // 恢复失败不能盖住原来的写入异常。
            }
            catch (InvalidComObjectException)
            {
            }
        }

        /// <summary>把一次成功的工作表写入记入文件日志。</summary>
        private void LogWriteCompleted(
            string operationId,
            string workbookName,
            string targetSheetName,
            BufferedQueryResult result,
            long rowsWritten,
            int blockCount,
            TimeSpan elapsed)
        {
            _logger.Write(
                LogSeverity.Information,
                "工作表写入完成。",
                operationId,
                properties: WriteProperties(workbookName, targetSheetName, result, rowsWritten, blockCount, elapsed, includeRate: true));
        }

        /// <summary>组装写入汇总的结构化属性。耗时不大于 0 时不写行每秒。</summary>
        private static Dictionary<string, string> WriteProperties(
            string workbookName,
            string targetSheetName,
            BufferedQueryResult result,
            long rowsWritten,
            int blockCount,
            TimeSpan elapsed,
            bool includeRate)
        {
            long elapsedMs = (long)elapsed.TotalMilliseconds;
            Dictionary<string, string> properties = new Dictionary<string, string>
            {
                ["workbook"] = workbookName ?? string.Empty,
                ["sheet"] = targetSheetName ?? string.Empty,
                ["rowsWritten"] = rowsWritten.ToString(),
                ["blockCount"] = blockCount.ToString(),
                ["columnCount"] = (result.Columns == null ? 0 : result.Columns.Count).ToString(),
                ["truncated"] = result.IsTruncated ? "true" : "false",
                ["elapsedMs"] = elapsedMs.ToString()
            };
            if (includeRate && elapsedMs > 0)
            {
                properties["rowsPerSecond"] = (rowsWritten * 1000L / elapsedMs).ToString();
            }

            return properties;
        }

        private static ExcelInterop.Worksheet GetOrCreateWorksheet(ExcelInterop.Workbook workbook, string name)
        {
            foreach (object item in workbook.Worksheets)
            {
                ExcelInterop.Worksheet worksheet = item as ExcelInterop.Worksheet;
                if (worksheet == null) continue;
                if (string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase)) return worksheet;
                Marshal.FinalReleaseComObject(worksheet);
            }

            ExcelInterop.Worksheet created = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]) as ExcelInterop.Worksheet;
            created.Name = name;
            return created;
        }

        /// <summary>按选项清空写入区域的值，并去掉该区域边框。整表模式处理已用区域；锚定模式只处理锚点向下的结果列，并按参数决定是否处理右侧原数据列。</summary>
        /// <remarks>数值格式、字体、对齐和填充留在原单元格上。边框先整块去掉，调用方写完后再只给新的标题和数据区域画回灰色网格。</remarks>
        private static void ClearBeforeWrite(ExcelInterop.Worksheet worksheet, SheetWriteOptions options, int columnCount)
        {
            ExcelInterop.Range usedRange = null;
            try
            {
                usedRange = worksheet.UsedRange;
                if (options.ClearEntireUsedRange)
                {
                    ClearValuesAndBorders(usedRange);
                    return;
                }

                int usedRow = usedRange.Row;
                int usedColumn = usedRange.Column;
                ExcelInterop.Range usedRows = null;
                ExcelInterop.Range usedColumns = null;
                int usedLastRow = 0;
                int usedLastColumn = 0;
                try
                {
                    usedRows = usedRange.Rows;
                    usedColumns = usedRange.Columns;
                    usedLastRow = usedRow + usedRows.Count - 1;
                    usedLastColumn = usedColumn + usedColumns.Count - 1;
                }
                finally
                {
                    Release(ref usedColumns);
                    Release(ref usedRows);
                }
                if (usedLastRow < options.AnchorRow) return;

                if (columnCount > 0)
                {
                    ClearRectangle(
                        worksheet,
                        options.AnchorRow,
                        options.AnchorColumn,
                        usedLastRow,
                        options.AnchorColumn + columnCount - 1);
                }

                int extraStart = options.AnchorColumn + Math.Max(columnCount, 0);
                if (options.ClearExtraColumns && usedLastColumn >= extraStart)
                {
                    ClearRectangle(worksheet, options.AnchorRow, extraStart, usedLastRow, usedLastColumn);
                }
            }
            finally
            {
                Release(ref usedRange);
            }
        }

        /// <summary>清空一个矩形的值并去掉边框。不触及矩形以外的单元格。</summary>
        /// <remarks>保留矩形内原有的数值格式、字体、对齐和填充。</remarks>
        private static void ClearRectangle(ExcelInterop.Worksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            if (endRow < startRow || endColumn < startColumn) return;
            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            try
            {
                start = worksheet.Cells[startRow, startColumn] as ExcelInterop.Range;
                end = worksheet.Cells[endRow, endColumn] as ExcelInterop.Range;
                range = worksheet.Range[start, end];
                ClearValuesAndBorders(range);
            }
            finally
            {
                Release(ref range);
                Release(ref end);
                Release(ref start);
            }
        }

        /// <summary>只清除区域内的值，并去掉边框。不改数值格式、字体、对齐和填充。</summary>
        /// <param name="range">要处理的区域。调用方负责释放。</param>
        private static void ClearValuesAndBorders(ExcelInterop.Range range)
        {
            ExcelInterop.Borders borders = null;
            try
            {
                range.ClearContents();
                borders = range.Borders;
                borders.LineStyle = ExcelInterop.XlLineStyle.xlLineStyleNone;
            }
            finally
            {
                Release(ref borders);
            }
        }

        /// <summary>写入一行字段编码标题。编码加粗，底色表示成功或截断。数据从下一行开始。</summary>
        /// <param name="headerColor">标题行底色。来自当时的成功色或截断色设置。</param>
        /// <remarks>只改标题行的加粗和底色。该行原有的数值格式、字体名称和对齐保留。</remarks>
        private static void WriteHeaders(
            ExcelInterop.Worksheet worksheet,
            BufferedQueryResult result,
            int anchorRow,
            int anchorColumn,
            Color headerColor)
        {
            int columnCount = result.Columns.Count;
            object[,] headers = new object[HeaderRowCount, columnCount];
            for (int column = 0; column < columnCount; column++)
            {
                ResultColumn item = result.Columns[column];
                headers[0, column] = item == null ? string.Empty : item.Name;
            }

            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            ExcelInterop.Font font = null;
            ExcelInterop.Interior interior = null;
            try
            {
                start = worksheet.Cells[anchorRow, anchorColumn] as ExcelInterop.Range;
                end = worksheet.Cells[anchorRow + HeaderRowCount - 1, anchorColumn + columnCount - 1] as ExcelInterop.Range;
                range = worksheet.Range[start, end];
                range.Value2 = headers;
                interior = range.Interior;
                interior.Color = ColorTranslator.ToOle(headerColor);
                font = range.Font;
                font.Bold = true;
            }
            finally
            {
                Release(ref font);
                Release(ref interior);
                Release(ref range);
                Release(ref end);
                Release(ref start);
            }
        }

        /// <summary>给标题和数据区域加上灰色细网格线。没有数据时只框住标题行。</summary>
        private static void ApplyTableBorder(ExcelInterop.Worksheet worksheet, int anchorRow, int anchorColumn, int lastRow, int columnCount)
        {
            if (lastRow < anchorRow || columnCount <= 0) return;
            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            ExcelInterop.Borders borders = null;
            try
            {
                start = worksheet.Cells[anchorRow, anchorColumn] as ExcelInterop.Range;
                end = worksheet.Cells[lastRow, anchorColumn + columnCount - 1] as ExcelInterop.Range;
                range = worksheet.Range[start, end];
                borders = range.Borders;
                borders.LineStyle = ExcelInterop.XlLineStyle.xlContinuous;
                borders.Weight = ExcelInterop.XlBorderWeight.xlThin;
                borders.Color = ColorTranslator.ToOle(Color.FromArgb(128, 128, 128));
            }
            finally
            {
                Release(ref borders);
                Release(ref range);
                Release(ref end);
                Release(ref start);
            }
        }

        private static void WriteBlock(ExcelInterop.Worksheet worksheet, int startRow, int startColumn, int columnCount, object[,] values, int rowCount)
        {
            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            try
            {
                start = worksheet.Cells[startRow, startColumn] as ExcelInterop.Range;
                end = worksheet.Cells[startRow + rowCount - 1, startColumn + columnCount - 1] as ExcelInterop.Range;
                range = worksheet.Range[start, end];
                range.Value2 = values;
            }
            finally
            {
                Release(ref range);
                Release(ref end);
                Release(ref start);
            }
        }

        private static void ValidateSheetName(string name)
        {
            if (name.Length > 31 || name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            {
                throw new ArgumentException("目标 Sheet 名称无效：" + name, nameof(name));
            }
        }

        private static void Release<T>(ref T value) where T : class
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
            value = null;
        }
    }
}

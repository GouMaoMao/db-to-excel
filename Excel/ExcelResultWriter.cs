using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.Excel
{
    /// <summary>使用 Excel COM Interop 将缓冲查询结果批量写入指定工作表。</summary>
    /// <remarks>
    /// 写入前会清空目标工作表的现有内容；目标工作表不存在时会创建。
    /// 每写完一块向进度接收器报告已写行数和块序号；总块数来自已缓冲结果，进度条按该比例填充。
    /// 文件日志只在写入完成或失败时各记一条汇总。
    /// 应在 Excel UI 线程调用，并避免在方法外继续使用本类释放的 COM Range 和 Worksheet 引用。
    /// </remarks>
    public sealed class ExcelResultWriter : IExcelResultWriter
    {
        private readonly ILogger _logger;

        /// <summary>创建 Excel 结果写入器。</summary>
        /// <param name="logger">写入阶段的文件日志。不按进度日志级别过滤。</param>
        public ExcelResultWriter(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
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
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(targetSheetName)) throw new ArgumentException("目标 Sheet 名称不能为空。", nameof(targetSheetName));
            if (result == null) throw new ArgumentNullException(nameof(result));
            ValidateSheetName(targetSheetName);

            string workbookName = workbook.Name;
            ExcelInterop.Worksheet worksheet = GetOrCreateWorksheet(workbook, targetSheetName);
            ExcelInterop.Range usedRange = null;
            Stopwatch writeWatch = Stopwatch.StartNew();
            long rowsWritten = 0;
            int totalBlocks = result.Blocks.Count;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                usedRange = worksheet.UsedRange;
                usedRange.ClearContents();
                Release(ref usedRange);

                int columnCount = result.Columns.Count;
                if (columnCount == 0)
                {
                    LogWriteCompleted(operationId, workbookName, targetSheetName, result, rowsWritten, totalBlocks, writeWatch.Elapsed);
                    return 0;
                }
                WriteHeaders(worksheet, result);

                int targetRow = 2;
                int writtenBlocks = 0;
                foreach (object[,] block in result.Blocks)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int rowCount = block.GetLength(0);
                    if (rowCount == 0) continue;
                    WriteBlock(worksheet, targetRow, columnCount, block, rowCount);
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
                Release(ref usedRange);
                Marshal.FinalReleaseComObject(worksheet);
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

        private static void WriteHeaders(ExcelInterop.Worksheet worksheet, BufferedQueryResult result)
        {
            int columnCount = result.Columns.Count;
            object[,] headers = new object[1, columnCount];
            for (int column = 0; column < columnCount; column++)
            {
                headers[0, column] = FormatHeaderText(result.Columns[column]);
            }

            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            try
            {
                start = worksheet.Cells[1, 1] as ExcelInterop.Range;
                end = worksheet.Cells[1, columnCount] as ExcelInterop.Range;
                range = worksheet.Range[start, end];
                range.Value2 = headers;
                range.Font.Bold = true;
            }
            finally
            {
                Release(ref range);
                Release(ref end);
                Release(ref start);
            }
        }

        private static void WriteBlock(ExcelInterop.Worksheet worksheet, int startRow, int columnCount, object[,] values, int rowCount)
        {
            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            try
            {
                start = worksheet.Cells[startRow, 1] as ExcelInterop.Range;
                end = worksheet.Cells[startRow + rowCount - 1, columnCount] as ExcelInterop.Range;
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

        private static string FormatHeaderText(ResultColumn column)
        {
            if (column == null) return string.Empty;
            Type dataType = column.GetDataType();
            string typeName = dataType == null ? string.Empty : dataType.Name;
            if (string.IsNullOrWhiteSpace(typeName)) return column.Name;
            return string.Format("{0}{1}{2}", column.Name, Environment.NewLine, typeName);
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

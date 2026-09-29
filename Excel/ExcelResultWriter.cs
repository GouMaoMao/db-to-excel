using System;
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
    /// 应在 Excel UI 线程调用，并避免在方法外继续使用本类释放的 COM Range 和 Worksheet 引用。
    /// </remarks>
    public sealed class ExcelResultWriter : IExcelResultWriter
    {
        /// <inheritdoc/>
        public long WriteResult(
            ExcelInterop.Workbook workbook,
            string targetSheetName,
            BufferedQueryResult result,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(targetSheetName)) throw new ArgumentException("目标 Sheet 名称不能为空。", nameof(targetSheetName));
            if (result == null) throw new ArgumentNullException(nameof(result));
            ValidateSheetName(targetSheetName);

            ExcelInterop.Worksheet worksheet = GetOrCreateWorksheet(workbook, targetSheetName);
            ExcelInterop.Range usedRange = null;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                usedRange = worksheet.UsedRange;
                usedRange.ClearContents();
                Release(ref usedRange);

                int columnCount = result.Columns.Count;
                if (columnCount == 0) return 0;
                WriteHeaders(worksheet, result);

                long rowsWritten = 0;
                int targetRow = 2;
                foreach (object[,] block in result.Blocks)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int rowCount = block.GetLength(0);
                    if (rowCount == 0) continue;
                    WriteBlock(worksheet, targetRow, columnCount, block, rowCount);
                    targetRow += rowCount;
                    rowsWritten += rowCount;
                    progress?.Report(new OperationProgress
                    {
                        Stage = OperationStage.Writing,
                        TargetSheetName = targetSheetName,
                        Message = "正在写入工作表…",
                        RowsRead = result.RowCount,
                        RowsWritten = rowsWritten,
                        IsTruncated = result.IsTruncated,
                        IsIndeterminate = true
                    });
                }

                return rowsWritten;
            }
            finally
            {
                Release(ref usedRange);
                Marshal.FinalReleaseComObject(worksheet);
            }
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

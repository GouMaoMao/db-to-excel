using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.Excel
{
    /// <summary>从工作簿中名为 SQL 的工作表读取单元格，并委托解析器生成刷新任务。</summary>
    /// <remarks>该类型直接访问 Excel COM 对象，应在 Excel UI 线程调用。</remarks>
    public sealed class SqlSheetTaskReader : ISqlSheetTaskReader
    {
        /// <summary>SQL 任务配置工作表的固定名称。</summary>
        public const string SqlSheetName = "SQL";

        private readonly SqlSheetTaskParser _parser = new SqlSheetTaskParser();

        /// <inheritdoc/>
        public bool Exists(ExcelInterop.Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            ExcelInterop.Worksheet worksheet = FindWorksheet(workbook, SqlSheetName);
            if (worksheet == null) return false;
            Marshal.FinalReleaseComObject(worksheet);
            return true;
        }

        /// <inheritdoc/>
        public void WriteColumnHeader(ExcelInterop.Workbook workbook, int sourceColumn, string headerText)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (sourceColumn < 1) throw new ArgumentOutOfRangeException(nameof(sourceColumn));

            ExcelInterop.Worksheet worksheet = FindWorksheet(workbook, SqlSheetName);
            if (worksheet == null)
            {
                throw new InvalidOperationException("当前工作簿中不存在名为“SQL”的工作表。");
            }

            ExcelInterop.Range cell = null;
            try
            {
                cell = worksheet.Cells[1, sourceColumn] as ExcelInterop.Range;
                if (cell == null) throw new InvalidOperationException("无法写入 SQL 页第 1 行。");
                cell.Value2 = headerText ?? string.Empty;
            }
            finally
            {
                if (cell != null) Marshal.FinalReleaseComObject(cell);
                Marshal.FinalReleaseComObject(worksheet);
            }
        }

        /// <inheritdoc/>
        public void ActivateSqlCell(ExcelInterop.Workbook workbook, int sourceColumn, int row)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (sourceColumn < 1) throw new ArgumentOutOfRangeException(nameof(sourceColumn));
            if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));

            ExcelInterop.Worksheet worksheet = FindWorksheet(workbook, SqlSheetName);
            if (worksheet == null)
            {
                throw new InvalidOperationException("当前工作簿中不存在名为“SQL”的工作表。");
            }

            ExcelInterop.Range cell = null;
            try
            {
                worksheet.Activate();
                cell = worksheet.Cells[row, sourceColumn] as ExcelInterop.Range;
                if (cell == null) throw new InvalidOperationException("无法定位 SQL 页单元格。");
                cell.Select();
            }
            finally
            {
                if (cell != null) Marshal.FinalReleaseComObject(cell);
                Marshal.FinalReleaseComObject(worksheet);
            }
        }

        /// <inheritdoc/>
        public bool TryActivateWorksheet(ExcelInterop.Workbook workbook, string sheetName)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            ExcelInterop.Worksheet worksheet = FindWorksheet(workbook, sheetName.Trim());
            if (worksheet == null) return false;
            try
            {
                worksheet.Activate();
                return true;
            }
            finally
            {
                Marshal.FinalReleaseComObject(worksheet);
            }
        }

        /// <inheritdoc/>
        public void WriteTaskColumn(ExcelInterop.Workbook workbook, string targetSheetName, string queryText)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(targetSheetName))
                throw new ArgumentException("目标 Sheet 名称不能为空。", nameof(targetSheetName));

            string target = targetSheetName.Trim();
            string[] lines = SplitSqlLines(queryText);
            ExcelInterop.Worksheet worksheet = GetOrCreateSqlSheet(workbook);
            try
            {
                int column;
                int clearThroughRow;
                LocateTaskColumn(worksheet, target, out column, out clearThroughRow);
                if (clearThroughRow >= 2)
                    ClearColumn(worksheet, column, 2, clearThroughRow);
                WriteTaskCells(worksheet, column, target, lines);
            }
            finally
            {
                Marshal.FinalReleaseComObject(worksheet);
            }
        }

        /// <inheritdoc/>
        public IReadOnlyList<RefreshTaskDefinition> ReadTasks(ExcelInterop.Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            ExcelInterop.Worksheet worksheet = FindWorksheet(workbook, SqlSheetName);
            if (worksheet == null)
            {
                throw new InvalidOperationException("当前工作簿中不存在名为“SQL”的工作表。");
            }

            ExcelInterop.Range usedRange = null;
            try
            {
                usedRange = worksheet.UsedRange;
                int firstColumn = usedRange.Column;
                int lastColumn = firstColumn + usedRange.Columns.Count - 1;
                int firstRow = usedRange.Row;
                int lastRow = Math.Max(2, firstRow + usedRange.Rows.Count - 1);
                object[,] cells = new object[lastRow, lastColumn - firstColumn + 1];

                for (int column = firstColumn; column <= lastColumn; column++)
                {
                    for (int row = 1; row <= lastRow; row++)
                    {
                        cells[row - 1, column - firstColumn] = ReadCellValue(worksheet, row, column);
                    }
                }

                return _parser.Parse(cells, firstColumn);
            }
            finally
            {
                if (usedRange != null) Marshal.FinalReleaseComObject(usedRange);
                Marshal.FinalReleaseComObject(worksheet);
            }
        }

        /// <summary>找到与目标表名相同的任务列；没有时用已用区域右侧的下一列。空表使用 A 列。</summary>
        /// <param name="worksheet">SQL 工作表。</param>
        /// <param name="targetSheetName">要比对的目标表名。</param>
        /// <param name="column">要写入的列号，从 1 开始。</param>
        /// <param name="clearThroughRow">更新已有列时需要清到的行号。新列不需要清理时为 1。</param>
        private static void LocateTaskColumn(
            ExcelInterop.Worksheet worksheet,
            string targetSheetName,
            out int column,
            out int clearThroughRow)
        {
            ExcelInterop.Range usedRange = null;
            try
            {
                usedRange = worksheet.UsedRange;
                int firstColumn = usedRange.Column;
                int lastColumn = firstColumn + usedRange.Columns.Count - 1;
                int lastRow = usedRange.Row + usedRange.Rows.Count - 1;
                int matched = 0;
                int lastOccupied = 0;
                for (int candidate = firstColumn; candidate <= lastColumn; candidate++)
                {
                    bool occupied = false;
                    for (int row = 1; row <= lastRow; row++)
                    {
                        string text = Convert.ToString(ReadCellValue(worksheet, row, candidate)) ?? string.Empty;
                        if (row == 1)
                        {
                            string sheetName = SqlSheetHeader.Parse(text).SheetName;
                            if (matched == 0
                                && !string.IsNullOrWhiteSpace(sheetName)
                                && string.Equals(sheetName, targetSheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                matched = candidate;
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(text)) occupied = true;
                    }

                    if (occupied) lastOccupied = candidate;
                }

                if (matched > 0)
                {
                    column = matched;
                    clearThroughRow = lastRow;
                    return;
                }

                column = lastOccupied == 0 ? 1 : lastOccupied + 1;
                clearThroughRow = 1;
            }
            finally
            {
                if (usedRange != null) Marshal.FinalReleaseComObject(usedRange);
            }
        }

        /// <summary>取得已有的 SQL 工作表；没有时在工作簿末尾新建。</summary>
        /// <param name="workbook">当前工作簿。</param>
        /// <returns>调用方负责释放的工作表引用。</returns>
        private static ExcelInterop.Worksheet GetOrCreateSqlSheet(ExcelInterop.Workbook workbook)
        {
            ExcelInterop.Worksheet existing = FindWorksheet(workbook, SqlSheetName);
            if (existing != null) return existing;

            ExcelInterop.Worksheet last = workbook.Worksheets[workbook.Worksheets.Count] as ExcelInterop.Worksheet;
            try
            {
                ExcelInterop.Worksheet created = workbook.Worksheets.Add(After: last) as ExcelInterop.Worksheet;
                if (created == null) throw new InvalidOperationException("无法创建 SQL 工作表。");
                try
                {
                    created.Name = SqlSheetName;
                    return created;
                }
                catch
                {
                    Marshal.FinalReleaseComObject(created);
                    throw;
                }
            }
            finally
            {
                if (last != null) Marshal.FinalReleaseComObject(last);
            }
        }

        /// <summary>清掉一列里的旧 SQL。只清值，保留单元格格式。</summary>
        /// <param name="worksheet">SQL 工作表。</param>
        /// <param name="column">列号，从 1 开始。</param>
        /// <param name="firstRow">起始行，从 1 开始。</param>
        /// <param name="lastRow">结束行，含本行。</param>
        private static void ClearColumn(ExcelInterop.Worksheet worksheet, int column, int firstRow, int lastRow)
        {
            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            try
            {
                start = worksheet.Cells[firstRow, column] as ExcelInterop.Range;
                end = worksheet.Cells[lastRow, column] as ExcelInterop.Range;
                range = worksheet.Range[start, end];
                range.ClearContents();
            }
            finally
            {
                if (range != null) Marshal.FinalReleaseComObject(range);
                if (end != null) Marshal.FinalReleaseComObject(end);
                if (start != null) Marshal.FinalReleaseComObject(start);
            }
        }

        /// <summary>把目标表名和 SQL 行写入同一列。第 2 行留空，SQL 从第 3 行开始。</summary>
        /// <param name="worksheet">SQL 工作表。</param>
        /// <param name="column">列号，从 1 开始。</param>
        /// <param name="targetSheetName">第 1 行文本。</param>
        /// <param name="lines">按换行拆开的 SQL。可含空行。</param>
        private static void WriteTaskCells(
            ExcelInterop.Worksheet worksheet,
            int column,
            string targetSheetName,
            string[] lines)
        {
            int rowCount = 2 + lines.Length;
            object[,] values = new object[rowCount, 1];
            values[0, 0] = targetSheetName;
            values[1, 0] = string.Empty;
            for (int index = 0; index < lines.Length; index++)
                values[index + 2, 0] = ExcelText(lines[index]);

            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            ExcelInterop.Range range = null;
            try
            {
                start = worksheet.Cells[1, column] as ExcelInterop.Range;
                end = worksheet.Cells[rowCount, column] as ExcelInterop.Range;
                range = worksheet.Range[start, end];
                range.Value2 = values;
            }
            finally
            {
                if (range != null) Marshal.FinalReleaseComObject(range);
                if (end != null) Marshal.FinalReleaseComObject(end);
                if (start != null) Marshal.FinalReleaseComObject(start);
            }
        }

        /// <summary>按换行拆开 SQL，保留空行。Windows、Unix 和旧 Mac 换行都分开。</summary>
        /// <param name="queryText">编辑器中的 SQL。</param>
        /// <returns>每个元素对应一个单元格。</returns>
        private static string[] SplitSqlLines(string queryText)
        {
            string text = queryText ?? string.Empty;
            return text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        }

        /// <summary>把会让 Excel 当成公式的注释行改成文本。</summary>
        /// <param name="line">一行 SQL。</param>
        /// <returns>以 <c>--</c> 开头时前面加一个空格。行首已有其他字符时原样返回。</returns>
        /// <remarks>加空格后仍是合法 SQL 注释。批量刷新会先跳过空白再识别 <c>--</c>。</remarks>
        private static string ExcelText(string line)
        {
            if (line != null && line.StartsWith("--", StringComparison.Ordinal))
                return " " + line;
            return line ?? string.Empty;
        }

        private static ExcelInterop.Worksheet FindWorksheet(ExcelInterop.Workbook workbook, string name)
        {
            foreach (object item in workbook.Worksheets)
            {
                ExcelInterop.Worksheet worksheet = item as ExcelInterop.Worksheet;
                if (worksheet == null) continue;
                if (string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
                Marshal.FinalReleaseComObject(worksheet);
            }
            return null;
        }

        private static object ReadCellValue(ExcelInterop.Worksheet worksheet, int row, int column)
        {
            ExcelInterop.Range cell = null;
            try
            {
                cell = worksheet.Cells[row, column] as ExcelInterop.Range;
                return cell?.Value2;
            }
            finally
            {
                if (cell != null) Marshal.FinalReleaseComObject(cell);
            }
        }
    }
}

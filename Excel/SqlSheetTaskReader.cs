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

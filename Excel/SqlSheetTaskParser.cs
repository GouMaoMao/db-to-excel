using System;
using System.Collections.Generic;
using System.Text;
using DB2Sheet.Models;

namespace DB2Sheet.Excel
{
    /// <summary>将 SQL Sheet 的二维单元格值解析为批量刷新任务。</summary>
    /// <remarks>
    /// 每一列代表一个任务：首行是目标表名，可在 <c>//</c> 后写特殊参数；后续非空单元格按行拼接为 SQL。
    /// 命令开始前的注释只进入摘要，仍保留在 SQL 文本中。任务记下命令开始的那一行，供列表跳转。
    /// </remarks>
    public sealed class SqlSheetTaskParser
    {
        /// <summary>解析二维单元格数组。</summary>
        /// <param name="cells">以“行、列”索引的单元格值；可使用任意数组下界。</param>
        /// <param name="firstColumn">数组第一列在 Excel 中的实际列号。</param>
        /// <returns>跳过空名称或空 SQL 后得到的只读任务列表。</returns>
        public IReadOnlyList<RefreshTaskDefinition> Parse(object[,] cells, int firstColumn)
        {
            if (cells == null) throw new ArgumentNullException(nameof(cells));
            if (firstColumn < 1) throw new ArgumentOutOfRangeException(nameof(firstColumn));

            int firstRowIndex = cells.GetLowerBound(0);
            int lastRowIndex = cells.GetUpperBound(0);
            int firstColumnIndex = cells.GetLowerBound(1);
            int lastColumnIndex = cells.GetUpperBound(1);
            List<RefreshTaskDefinition> tasks = new List<RefreshTaskDefinition>();

            for (int columnIndex = firstColumnIndex; columnIndex <= lastColumnIndex; columnIndex++)
            {
                string headerText = Convert.ToString(cells[firstRowIndex, columnIndex]) ?? string.Empty;
                SqlSheetHeader header = SqlSheetHeader.Parse(headerText);
                if (string.IsNullOrWhiteSpace(header.SheetName)) continue;

                StringBuilder query = new StringBuilder();
                int firstBodyRow = 0;
                int commandRow = 0;
                for (int rowIndex = firstRowIndex + 1; rowIndex <= lastRowIndex; rowIndex++)
                {
                    string fragment = Convert.ToString(cells[rowIndex, columnIndex]) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(fragment)) continue;
                    int excelRow = rowIndex - firstRowIndex + 1;
                    if (firstBodyRow == 0) firstBodyRow = excelRow;
                    if (query.Length > 0) query.AppendLine();
                    query.Append(fragment);
                    if (commandRow == 0 && !SqlSheetHeader.IsOnlyLeadingComment(query.ToString()))
                        commandRow = excelRow;
                }

                if (query.Length == 0) continue;
                if (commandRow == 0) commandRow = firstBodyRow;

                int sourceColumn = firstColumn + columnIndex - firstColumnIndex;
                string queryText = query.ToString();
                tasks.Add(new RefreshTaskDefinition(
                    "column-" + sourceColumn,
                    header.SheetName,
                    queryText,
                    sourceColumn,
                    header.OptionsText,
                    header.StartCell,
                    header.StartRow,
                    header.StartColumn,
                    header.ClearExtraColumns,
                    header.Error,
                    SqlSheetHeader.ExtractLeadingComment(queryText),
                    header.Normalized,
                    commandRow));
            }

            return tasks.AsReadOnly();
        }
    }
}

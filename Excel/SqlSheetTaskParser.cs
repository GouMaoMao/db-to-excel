using System;
using System.Collections.Generic;
using System.Text;
using DB2Sheet.Models;

namespace DB2Sheet.Excel
{
    /// <summary>将 SQL Sheet 的二维单元格值解析为批量刷新任务。</summary>
    /// <remarks>每一列代表一个任务：首行是目标 Sheet 名，后续非空单元格按行拼接为 SQL。</remarks>
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
                string targetSheetName = Convert.ToString(cells[firstRowIndex, columnIndex]) ?? string.Empty;
                targetSheetName = targetSheetName.Trim();
                if (string.IsNullOrWhiteSpace(targetSheetName)) continue;

                StringBuilder query = new StringBuilder();
                for (int rowIndex = firstRowIndex + 1; rowIndex <= lastRowIndex; rowIndex++)
                {
                    string fragment = Convert.ToString(cells[rowIndex, columnIndex]) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(fragment)) continue;
                    if (query.Length > 0) query.AppendLine();
                    query.Append(fragment);
                }

                if (query.Length == 0) continue;

                int sourceColumn = firstColumn + columnIndex - firstColumnIndex;
                tasks.Add(new RefreshTaskDefinition(
                    "column-" + sourceColumn,
                    targetSheetName,
                    query.ToString(),
                    sourceColumn));
            }

            return tasks.AsReadOnly();
        }
    }
}

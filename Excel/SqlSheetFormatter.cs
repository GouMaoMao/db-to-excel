using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;
using DB2Sheet.UI;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.Excel
{
    /// <summary>按当前调色板美化 SQL 工作表的字体和颜色，不改单元格文字，也不画边框。</summary>
    /// <remarks>
    /// 应在 Excel UI 线程调用。整表美化期间关闭重画，结束、取消或出错时恢复进入前的开关。
    /// 先清空整张表的格式，再给有表名的列上色，然后冻结首行并设置标签颜色。
    /// 第 1 行只画表名和参数。第 2 行起按 AvalonEdit 的 T-SQL 定义上色。公式格整格用公式色，不套语法色。
    /// </remarks>
    internal static class SqlSheetFormatter
    {
        private static readonly Color ParameterColor = Color.FromArgb(128, 128, 128);

        /// <summary>统计第 1 行有表名的列数。</summary>
        /// <param name="worksheet">SQL 工作表。调用方保持引用。</param>
        /// <returns>可美化的列数。没有表名时为 0。</returns>
        public static int CountColumns(ExcelInterop.Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            return ListColumns(worksheet).Count;
        }

        /// <summary>美化每一个有表名的列，并按列报告进度。</summary>
        /// <param name="worksheet">SQL 工作表。调用方在方法返回前保持引用。</param>
        /// <param name="palette">表名、公式和标签颜色。为 null 时抛出。</param>
        /// <param name="progress">进度接收器。可为空。</param>
        /// <param name="cancellationToken">清空前和每列开始前检查。当前列做完才停。</param>
        /// <returns>已美化的列数。没有可美化的列时为 0，且不改工作表。</returns>
        /// <remarks>
        /// 进度只用单条：阶段为执行，不填读取个数和写入百分比。每列开始时该列百分比为 0，做完为 100。
        /// 每列前后让出一次，进度窗才能重画。取消抛出 <see cref="OperationCanceledException"/>，已完成的列保留。
        /// 全部列做完后冻结首行并设置标签颜色。冻结失败时颜色仍然保留。不把 Excel 抢到进度窗前面。
        /// </remarks>
        public static async Task<int> FormatSheet(
            ExcelInterop.Worksheet worksheet,
            SqlSheetPalette palette,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (palette == null) throw new ArgumentNullException(nameof(palette));

            List<StyleColumn> columns = ListColumns(worksheet);
            if (columns.Count == 0) return 0;

            ExcelInterop.Application application = null;
            bool restoreScreenUpdating = false;
            bool previousScreenUpdating = true;
            try
            {
                Report(progress, 1, columns.Count, worksheet.Name, 0,
                    string.Format("检测到 {0:N0} 列，正在清空 SQL 页原有格式…", columns.Count));
                try
                {
                    application = worksheet.Application;
                    previousScreenUpdating = application.ScreenUpdating;
                    restoreScreenUpdating = true;
                    application.ScreenUpdating = false;
                }
                catch (COMException)
                {
                }
                catch (InvalidComObjectException)
                {
                }

                cancellationToken.ThrowIfCancellationRequested();
                ClearSheetFormats(worksheet);
                Report(progress, 1, columns.Count, worksheet.Name, 0, "原有格式已清空，开始逐列美化。");
                for (int index = 0; index < columns.Count; index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    StyleColumn column = columns[index];
                    Report(progress, index + 1, columns.Count, column.Name, 0, "正在美化。");
                    await Task.Yield();
                    cancellationToken.ThrowIfCancellationRequested();
                    ApplyColumn(worksheet, column.Column, palette);
                    Report(progress, index + 1, columns.Count, column.Name, 100, "已美化。");
                    await Task.Yield();
                }

                Report(progress, columns.Count, columns.Count, worksheet.Name, 100,
                    "所有列已美化，正在冻结首行并设置工作表标签颜色…");
                TryFreezeHeader(worksheet);
                SetTabColor(worksheet, palette.Tab);
                Report(progress, columns.Count, columns.Count, worksheet.Name, 100,
                    string.Format("SQL 页美化完成，共处理 {0:N0} 列。", columns.Count));
                return columns.Count;
            }
            finally
            {
                RestoreScreenUpdating(application, restoreScreenUpdating, previousScreenUpdating);
            }
        }

        /// <summary>只重套第 1 行的表名颜色。不改 SQL 正文，也不改工作表标签。</summary>
        /// <param name="cell">第 1 行单元格。调用方保持并释放。</param>
        /// <param name="headerText">刚写入的首行文本，用来给 <c>//</c> 后的参数着灰色。</param>
        /// <param name="palette">当前颜色。</param>
        public static void StyleHeader(ExcelInterop.Range cell, string headerText, SqlSheetPalette palette)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            if (palette == null) throw new ArgumentNullException(nameof(palette));
            PaintHeader(cell, headerText ?? string.Empty, palette);
        }

        private static void ApplyColumn(ExcelInterop.Worksheet worksheet, int column, SqlSheetPalette palette)
        {
            int contentLast = LastContentRow(worksheet, column);
            if (contentLast < 1) contentLast = 1;

            string[] texts = ReadTexts(worksheet, column, contentLast, out bool[] formulas);
            PaintBase(worksheet, column, 1, contentLast);
            ExcelInterop.Range header = null;
            try
            {
                header = CellAt(worksheet, 1, column);
                PaintHeader(header, texts.Length > 0 ? texts[0] : string.Empty, palette);
                PaintBody(worksheet, column, texts, formulas, palette);
            }
            finally
            {
                Release(ref header);
            }
        }

        private static void PaintHeader(ExcelInterop.Range cell, string headerText, SqlSheetPalette palette)
        {
            SetFont(cell, palette.Header, bold: true);
            ClearFill(cell);
            int separator = headerText.IndexOf("//", StringComparison.Ordinal);
            if (separator < 0 || separator >= headerText.Length) return;
            ColorCharacters(cell, separator, headerText.Length - separator, ParameterColor, bold: false);
        }

        private static void PaintBody(
            ExcelInterop.Worksheet worksheet,
            int column,
            string[] texts,
            bool[] formulas,
            SqlSheetPalette palette)
        {
            if (texts.Length <= 1) return;
            BodyStyle[] styles = ClassifyBody(texts, formulas, palette);
            int runStart = -1;
            int runEnd = -1;
            BodyStyle run = default(BodyStyle);
            for (int index = 0; index < styles.Length; index++)
            {
                int row = index + 2;
                BodyStyle style = styles[index];
                if (style.Uniform)
                {
                    if (runStart >= 0 && SameUniform(run, style))
                    {
                        runEnd = row;
                        continue;
                    }

                    FlushUniform(worksheet, column, ref runStart, ref runEnd, ref run);
                    runStart = row;
                    runEnd = row;
                    run = style;
                    continue;
                }

                FlushUniform(worksheet, column, ref runStart, ref runEnd, ref run);
                if (style.Pieces == null) continue;
                ExcelInterop.Range cell = null;
                try
                {
                    cell = CellAt(worksheet, row, column);
                    foreach (TextPiece piece in style.Pieces)
                        ColorCharacters(cell, piece.Start, piece.Length, piece.Color, piece.Bold);
                }
                finally
                {
                    Release(ref cell);
                }
            }

            FlushUniform(worksheet, column, ref runStart, ref runEnd, ref run);
        }

        private static void FlushUniform(
            ExcelInterop.Worksheet worksheet,
            int column,
            ref int runStart,
            ref int runEnd,
            ref BodyStyle run)
        {
            if (runStart < 0) return;
            ExcelInterop.Range range = null;
            try
            {
                range = ColumnRange(worksheet, column, runStart, runEnd);
                SetFont(range, run.Color, run.Bold);
                if (run.Fill) SetFill(range, SqlSheetPalette.Lighten(run.Color));
            }
            finally
            {
                Release(ref range);
            }

            runStart = -1;
            runEnd = -1;
        }

        private static bool SameUniform(BodyStyle left, BodyStyle right)
        {
            return left.Fill == right.Fill && left.Bold == right.Bold && left.Color.ToArgb() == right.Color.ToArgb();
        }

        private static BodyStyle[] ClassifyBody(string[] texts, bool[] formulas, SqlSheetPalette palette)
        {
            int bodyCount = texts.Length - 1;
            string[] lines = new string[bodyCount];
            int[] starts = new int[bodyCount];
            StringBuilder joined = new StringBuilder();
            for (int index = 0; index < bodyCount; index++)
            {
                starts[index] = joined.Length;
                string text = texts[index + 1] ?? string.Empty;
                lines[index] = text;
                joined.Append(text);
                if (index + 1 < bodyCount) joined.Append('\n');
            }

            IReadOnlyList<SqlHighlightSpan> spans = SqlSyntax.Highlight(joined.ToString());
            BodyStyle[] styles = new BodyStyle[bodyCount];
            for (int index = 0; index < bodyCount; index++)
            {
                if (formulas[index + 1])
                {
                    styles[index] = BodyStyle.UniformStyle(palette.Formula, fill: true, bold: false);
                    continue;
                }

                styles[index] = StyleOfLine(lines[index], starts[index], spans);
            }

            return styles;
        }

        private static BodyStyle StyleOfLine(string text, int lineStart, IReadOnlyList<SqlHighlightSpan> spans)
        {
            int lineEnd = lineStart + text.Length;
            List<TextPiece> pieces = null;
            Color color = Color.Black;
            bool bold = false;
            bool saw = false;
            bool mixed = false;
            bool[] covered = text.Length == 0 ? null : new bool[text.Length];
            foreach (SqlHighlightSpan span in spans)
            {
                int spanEnd = span.Start + span.Length;
                if (spanEnd <= lineStart || span.Start >= lineEnd) continue;
                int localStart = Math.Max(span.Start, lineStart) - lineStart;
                int localEnd = Math.Min(spanEnd, lineEnd) - lineStart;
                if (localEnd <= localStart) continue;
                if (!saw)
                {
                    color = span.Color;
                    bold = span.Bold;
                    saw = true;
                }
                else if (span.Bold != bold || span.Color.ToArgb() != color.ToArgb())
                {
                    mixed = true;
                }

                if (pieces == null) pieces = new List<TextPiece>();
                pieces.Add(new TextPiece(localStart, localEnd - localStart, span.Color, span.Bold));
                for (int index = localStart; index < localEnd; index++)
                    covered[index] = true;
            }

            if (!saw) return default(BodyStyle);
            if (!mixed && IsWhitespaceUncovered(text, covered))
                return BodyStyle.UniformStyle(color, fill: false, bold);
            return BodyStyle.Mixed(pieces);
        }

        private static bool IsWhitespaceUncovered(string text, bool[] covered)
        {
            for (int index = 0; index < text.Length; index++)
            {
                if (!covered[index] && !char.IsWhiteSpace(text[index])) return false;
            }

            return true;
        }

        private static string[] ReadTexts(
            ExcelInterop.Worksheet worksheet,
            int column,
            int lastRow,
            out bool[] formulas)
        {
            formulas = new bool[Math.Max(0, lastRow)];
            ExcelInterop.Range range = null;
            try
            {
                range = ColumnRange(worksheet, column, 1, lastRow);
                object[,] values = ToGrid(range.Value2, lastRow, 1);
                object[,] formulaGrid = ToGrid(range.Formula, lastRow, 1);
                string[] texts = new string[lastRow];
                for (int index = 0; index < lastRow; index++)
                {
                    object formula = CellOf(formulaGrid, index);
                    if (index > 0 && IsFormula(formula))
                    {
                        formulas[index] = true;
                        texts[index] = formula as string ?? string.Empty;
                        continue;
                    }

                    texts[index] = Convert.ToString(CellOf(values, index)) ?? string.Empty;
                }

                return texts;
            }
            finally
            {
                Release(ref range);
            }
        }

        private static bool IsFormula(object formula)
        {
            string text = formula as string;
            if (string.IsNullOrEmpty(text)) return false;
            int index = 0;
            while (index < text.Length && char.IsWhiteSpace(text[index])) index++;
            if (index >= text.Length) return false;
            if (text[index] == '=') return true;
            return index + 1 < text.Length && text[index] == '{' && text[index + 1] == '=';
        }

        private static List<StyleColumn> ListColumns(ExcelInterop.Worksheet worksheet)
        {
            List<StyleColumn> columns = new List<StyleColumn>();
            ExcelInterop.Range used = null;
            ExcelInterop.Range headers = null;
            try
            {
                used = worksheet.UsedRange;
                int firstColumn = used.Column;
                int lastColumn = firstColumn + used.Columns.Count - 1;
                if (used.Row > 1) return columns;
                headers = ColumnRange(worksheet, firstColumn, 1, 1);
                if (lastColumn != firstColumn)
                {
                    Release(ref headers);
                    ExcelInterop.Range start = null;
                    ExcelInterop.Range end = null;
                    try
                    {
                        start = CellAt(worksheet, 1, firstColumn);
                        end = CellAt(worksheet, 1, lastColumn);
                        headers = worksheet.Range[start, end];
                    }
                    finally
                    {
                        Release(ref end);
                        Release(ref start);
                    }
                }

                object[,] grid = ToGrid(headers.Value2, 1, lastColumn - firstColumn + 1);
                for (int offset = 0; offset <= lastColumn - firstColumn; offset++)
                {
                    string text = Convert.ToString(CellOf(grid, 0, offset)) ?? string.Empty;
                    string name = SqlSheetHeader.Parse(text).SheetName;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    columns.Add(new StyleColumn(firstColumn + offset, name));
                }

                return columns;
            }
            finally
            {
                Release(ref headers);
                Release(ref used);
            }
        }

        /// <summary>清空整张工作表的单元格格式。不改文字。</summary>
        /// <param name="worksheet">SQL 工作表。</param>
        /// <remarks>整表清空会把列宽恢复成默认。已用区域里的列宽先记下，清空后再写回去。</remarks>
        private static void ClearSheetFormats(ExcelInterop.Worksheet worksheet)
        {
            ExcelInterop.Range used = null;
            ExcelInterop.Range cells = null;
            try
            {
                used = worksheet.UsedRange;
                int firstColumn = used.Column;
                int count = used.Columns.Count;
                double[] widths = new double[count];
                for (int index = 0; index < count; index++)
                    widths[index] = ColumnWidth(worksheet, firstColumn + index);

                cells = worksheet.Cells;
                cells.ClearFormats();

                for (int index = 0; index < count; index++)
                    SetColumnWidth(worksheet, firstColumn + index, widths[index]);
            }
            finally
            {
                Release(ref cells);
                Release(ref used);
            }
        }

        /// <summary>读取一列的列宽。</summary>
        /// <param name="worksheet">工作表。</param>
        /// <param name="column">列号，从 1 开始。</param>
        /// <returns>Excel 列宽。</returns>
        private static double ColumnWidth(ExcelInterop.Worksheet worksheet, int column)
        {
            ExcelInterop.Range range = null;
            try
            {
                range = worksheet.Columns[column] as ExcelInterop.Range;
                return (double)range.ColumnWidth;
            }
            finally
            {
                Release(ref range);
            }
        }

        /// <summary>写回一列的列宽。</summary>
        /// <param name="worksheet">工作表。</param>
        /// <param name="column">列号，从 1 开始。</param>
        /// <param name="width">清空格式前的列宽。</param>
        private static void SetColumnWidth(ExcelInterop.Worksheet worksheet, int column, double width)
        {
            ExcelInterop.Range range = null;
            try
            {
                range = worksheet.Columns[column] as ExcelInterop.Range;
                range.ColumnWidth = width;
            }
            finally
            {
                Release(ref range);
            }
        }

        /// <summary>冻结第 1 行。失败时保持原冻结，不让这次美化失败。</summary>
        /// <param name="worksheet">SQL 工作表。方法内会激活它，但不把 Excel 抢到进度窗前面。</param>
        /// <remarks>重画仍由调用方关闭。只改窗口冻结，不选择单元格。</remarks>
        private static void TryFreezeHeader(ExcelInterop.Worksheet worksheet)
        {
            ExcelInterop.Window window = null;
            try
            {
                worksheet.Activate();
                window = worksheet.Application.ActiveWindow;
                if (window == null) return;
                window.FreezePanes = false;
                window.SplitColumn = 0;
                window.SplitRow = 1;
                window.FreezePanes = true;
            }
            catch (COMException)
            {
                // 冻结只是附加效果。窗口状态读不到时，已经画上的颜色仍然保留。
            }
            catch (InvalidComObjectException)
            {
            }
            finally
            {
                Release(ref window);
            }
        }

        /// <summary>该列最后一个有内容的行。从列底向上找，空列回到第 1 行。</summary>
        /// <remarks>只认值和公式，不把仅有格式的空单元格算进去。</remarks>
        private static int LastContentRow(ExcelInterop.Worksheet worksheet, int column)
        {
            ExcelInterop.Range bottom = null;
            ExcelInterop.Range last = null;
            try
            {
                bottom = worksheet.Cells[worksheet.Rows.Count, column] as ExcelInterop.Range;
                last = bottom.End[ExcelInterop.XlDirection.xlUp] as ExcelInterop.Range;
                return last.Row < 1 ? 1 : last.Row;
            }
            finally
            {
                Release(ref last);
                Release(ref bottom);
            }
        }

        private static void PaintBase(ExcelInterop.Worksheet worksheet, int column, int firstRow, int lastRow)
        {
            ExcelInterop.Range range = null;
            try
            {
                range = ColumnRange(worksheet, column, firstRow, lastRow);
                SetFont(range, Color.Black, bold: false);
                ClearFill(range);
            }
            finally
            {
                Release(ref range);
            }
        }

        private static void Report(
            IProgress<OperationProgress> progress,
            int current,
            int total,
            string name,
            int percent,
            string message)
        {
            progress?.Report(new OperationProgress
            {
                Stage = OperationStage.Executing,
                TargetSheetName = name,
                Message = message,
                TotalTasks = total,
                CurrentTask = current,
                Percent = percent,
                IsIndeterminate = false
            });
        }

        private static void SetTabColor(ExcelInterop.Worksheet worksheet, Color color)
        {
            ExcelInterop.Tab tab = null;
            try
            {
                tab = worksheet.Tab;
                tab.Color = ColorTranslator.ToOle(color);
            }
            finally
            {
                Release(ref tab);
            }
        }

        private static void SetFont(ExcelInterop.Range range, Color color, bool bold)
        {
            ExcelInterop.Font font = null;
            try
            {
                font = range.Font;
                font.Name = AppPresentation.SqlSheetFontName;
                font.Size = AppPresentation.SqlSheetFontSize;
                font.Color = ColorTranslator.ToOle(color);
                font.Bold = bold;
            }
            finally
            {
                Release(ref font);
            }
        }

        private static void ColorCharacters(ExcelInterop.Range cell, int start, int length, Color color, bool bold)
        {
            if (length <= 0 || start < 0) return;
            ExcelInterop.Characters characters = null;
            ExcelInterop.Font font = null;
            try
            {
                characters = cell.get_Characters(start + 1, length);
                font = characters.Font;
                font.Name = AppPresentation.SqlSheetFontName;
                font.Size = AppPresentation.SqlSheetFontSize;
                font.Color = ColorTranslator.ToOle(color);
                font.Bold = bold;
            }
            finally
            {
                Release(ref font);
                Release(ref characters);
            }
        }

        private static void ClearFill(ExcelInterop.Range range)
        {
            ExcelInterop.Interior interior = null;
            try
            {
                interior = range.Interior;
                interior.ColorIndex = ExcelInterop.XlColorIndex.xlColorIndexNone;
            }
            finally
            {
                Release(ref interior);
            }
        }

        private static void SetFill(ExcelInterop.Range range, Color color)
        {
            ExcelInterop.Interior interior = null;
            try
            {
                interior = range.Interior;
                interior.Color = ColorTranslator.ToOle(color);
            }
            finally
            {
                Release(ref interior);
            }
        }

        private static void RestoreScreenUpdating(ExcelInterop.Application application, bool restore, bool previous)
        {
            if (!restore || application == null) return;
            try
            {
                application.ScreenUpdating = previous;
            }
            catch (COMException)
            {
            }
            catch (InvalidComObjectException)
            {
            }
        }

        private static ExcelInterop.Range CellAt(ExcelInterop.Worksheet worksheet, int row, int column)
        {
            return worksheet.Cells[row, column] as ExcelInterop.Range;
        }

        private static ExcelInterop.Range ColumnRange(ExcelInterop.Worksheet worksheet, int column, int firstRow, int lastRow)
        {
            ExcelInterop.Range start = null;
            ExcelInterop.Range end = null;
            try
            {
                start = CellAt(worksheet, firstRow, column);
                end = CellAt(worksheet, lastRow, column);
                return worksheet.Range[start, end];
            }
            finally
            {
                Release(ref end);
                Release(ref start);
            }
        }

        private static object[,] ToGrid(object raw, int rows, int columns)
        {
            object[,] grid = raw as object[,];
            if (grid != null) return grid;
            object[,] single = new object[rows, columns];
            if (rows == 1 && columns == 1) single[0, 0] = raw;
            return single;
        }

        private static object CellOf(object[,] grid, int rowOffset, int columnOffset)
        {
            return grid[grid.GetLowerBound(0) + rowOffset, grid.GetLowerBound(1) + columnOffset];
        }

        private static object CellOf(object[,] grid, int rowOffset)
        {
            return CellOf(grid, rowOffset, 0);
        }

        private static void Release<T>(ref T value) where T : class
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
            value = null;
        }

        /// <summary>一列任务在 SQL 页中的列号和目标表名。</summary>
        private struct StyleColumn
        {
            public StyleColumn(int column, string name)
            {
                Column = column;
                Name = name ?? string.Empty;
            }

            public int Column { get; }
            public string Name { get; }
        }

        /// <summary>一行 SQL 的着法。整行同色时用区域一次设置；混色时按片段设置。</summary>
        private struct BodyStyle
        {
            public bool Uniform { get; private set; }
            public bool Fill { get; private set; }
            public bool Bold { get; private set; }
            public Color Color { get; private set; }
            public List<TextPiece> Pieces { get; private set; }

            public static BodyStyle UniformStyle(Color color, bool fill, bool bold)
            {
                return new BodyStyle { Uniform = true, Fill = fill, Bold = bold, Color = color };
            }

            public static BodyStyle Mixed(List<TextPiece> pieces)
            {
                return new BodyStyle { Pieces = pieces };
            }
        }

        /// <summary>同一单元格里的一段着色。起点从 0 开始。</summary>
        private struct TextPiece
        {
            public TextPiece(int start, int length, Color color, bool bold)
            {
                Start = start;
                Length = length;
                Color = color;
                Bold = bold;
            }

            public int Start { get; }
            public int Length { get; }
            public Color Color { get; }
            public bool Bold { get; }
        }
    }
}

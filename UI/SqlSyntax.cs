using System;
using System.Collections.Generic;
using System.Drawing;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Highlighting;

namespace DB2Sheet.UI
{
    /// <summary>提供 AvalonEdit 自带的 T-SQL 语法定义，以及按该定义拆出的着色区间。</summary>
    /// <remarks>
    /// 查询编辑器和 SQL 页正文共用这一份定义，不再维护第二套关键字表。
    /// 定义由 AvalonEdit 在首次读取时从程序集资源加载。找不到时抛出，避免两边静默变成纯黑文本。
    /// <see cref="Highlight"/> 只读文本，可在 Excel UI 线程调用；它不修改文档，也不接触 Excel。
    /// </remarks>
    internal static class SqlSyntax
    {
        /// <summary>AvalonEdit 注册的 T-SQL 定义名称。</summary>
        public const string DefinitionName = "TSQL";

        /// <summary>获取已注册的 T-SQL 语法定义。</summary>
        /// <returns>供编辑器直接挂上、或供 <see cref="Highlight"/> 使用的定义。</returns>
        /// <exception cref="InvalidOperationException">当前 AvalonEdit 程序集里没有这份定义。</exception>
        public static IHighlightingDefinition Definition
        {
            get
            {
                IHighlightingDefinition definition = HighlightingManager.Instance.GetDefinition(DefinitionName);
                if (definition == null)
                    throw new InvalidOperationException("找不到 AvalonEdit 的 T-SQL 语法定义。");
                return definition;
            }
        }

        /// <summary>按 T-SQL 定义标出需要改颜色或加粗的区间。</summary>
        /// <param name="sql">待着色文本。换行用换行符。为空时不产生区间。</param>
        /// <returns>按出现顺序排列的区间。前景色为空的片段不包含在内。</returns>
        /// <remarks>从第一行依次高亮，块注释才能跨行延续。调用方负责把区间写到编辑器或单元格。</remarks>
        public static IReadOnlyList<SqlHighlightSpan> Highlight(string sql)
        {
            TextDocument document = new TextDocument(sql ?? string.Empty);
            DocumentHighlighter highlighter = new DocumentHighlighter(document, Definition);
            try
            {
                List<SqlHighlightSpan> spans = new List<SqlHighlightSpan>();
                for (int lineNumber = 1; lineNumber <= document.LineCount; lineNumber++)
                {
                    HighlightedLine line = highlighter.HighlightLine(lineNumber);
                    foreach (HighlightedSection section in line.Sections)
                        TryAdd(spans, section);
                }

                return spans;
            }
            finally
            {
                DisposeHighlighter(highlighter);
            }
        }

        /// <summary>把有前景色的一节收成区间。黑色且不加粗的节留给底色，不单独写出。</summary>
        /// <param name="spans">正在收集的区间。</param>
        /// <param name="section">高亮器给出的一节。颜色为空时忽略。</param>
        private static void TryAdd(List<SqlHighlightSpan> spans, HighlightedSection section)
        {
            if (section == null || section.Length <= 0 || section.Color == null || section.Color.Foreground == null)
                return;

            System.Windows.Media.SolidColorBrush brush =
                section.Color.Foreground.GetBrush(null) as System.Windows.Media.SolidColorBrush;
            if (brush == null) return;

            System.Windows.Media.Color media = brush.Color;
            bool bold = section.Color.FontWeight.HasValue &&
                        section.Color.FontWeight.Value.ToOpenTypeWeight() >= System.Windows.FontWeights.Bold.ToOpenTypeWeight();
            if (!bold && media.R == 0 && media.G == 0 && media.B == 0) return;

            spans.Add(new SqlHighlightSpan(
                section.Offset,
                section.Length,
                Color.FromArgb(media.R, media.G, media.B),
                bold));
        }

        /// <summary>释放高亮器持有的文档订阅。没有释放方法时保持原样。</summary>
        /// <param name="highlighter">本次高亮使用的实例。</param>
        private static void DisposeHighlighter(DocumentHighlighter highlighter)
        {
            IDisposable disposable = highlighter as IDisposable;
            if (disposable != null) disposable.Dispose();
        }
    }

    /// <summary>T-SQL 定义标出的一段原文。起点从 0 开始，长度含该段全部字符。</summary>
    internal struct SqlHighlightSpan
    {
        /// <summary>创建一段着色区间。</summary>
        /// <param name="start">在源文本中的起点，从 0 开始。</param>
        /// <param name="length">字符数。调用方保证大于 0。</param>
        /// <param name="color">字体颜色，不含透明度。</param>
        /// <param name="bold">该段是否加粗。</param>
        public SqlHighlightSpan(int start, int length, Color color, bool bold)
        {
            Start = start;
            Length = length;
            Color = color;
            Bold = bold;
        }

        /// <summary>获取起点，从 0 开始。</summary>
        public int Start { get; }

        /// <summary>获取字符数。</summary>
        public int Length { get; }

        /// <summary>获取字体颜色。</summary>
        public Color Color { get; }

        /// <summary>获取该段是否加粗。</summary>
        public bool Bold { get; }
    }
}

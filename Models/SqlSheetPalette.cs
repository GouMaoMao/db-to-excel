using System.Drawing;

namespace DB2Sheet.Models
{
    /// <summary>SQL 页美化使用的字体颜色和工作表标签颜色。</summary>
    /// <remarks>
    /// 六项分别保存在参数设置里，这里只是一次美化要用的一组。公式单元格的浅底由公式色淡化得到，不单独保存。参数灰色也不在这六项里。
    /// </remarks>
    public sealed class SqlSheetPalette
    {
        /// <summary>创建一组颜色。</summary>
        /// <param name="header">目标表名颜色。</param>
        /// <param name="comment">注释颜色。</param>
        /// <param name="formula">公式字体颜色。</param>
        /// <param name="keyword">关键字颜色。</param>
        /// <param name="literal">单引号文本颜色。</param>
        /// <param name="tab">SQL 工作表底部标签的颜色。</param>
        public SqlSheetPalette(Color header, Color comment, Color formula, Color keyword, Color literal, Color tab)
        {
            Header = header;
            Comment = comment;
            Formula = formula;
            Keyword = keyword;
            Literal = literal;
            Tab = tab;
        }

        /// <summary>获取默认颜色。目标表名和公式为深红，注释为结果表成功绿，关键字为深蓝，引号文本为褐色，标签为暗蓝。</summary>
        public static SqlSheetPalette Default { get; } = new SqlSheetPalette(
            Color.FromArgb(153, 0, 0),
            Color.FromArgb(46, 160, 67),
            Color.FromArgb(153, 0, 0),
            Color.FromArgb(0, 51, 153),
            Color.FromArgb(139, 69, 19),
            Color.FromArgb(44, 74, 110));

        /// <summary>获取目标表名颜色。</summary>
        public Color Header { get; }

        /// <summary>获取注释颜色。</summary>
        public Color Comment { get; }

        /// <summary>获取公式字体颜色。</summary>
        public Color Formula { get; }

        /// <summary>获取关键字颜色。</summary>
        public Color Keyword { get; }

        /// <summary>获取单引号文本颜色。</summary>
        public Color Literal { get; }

        /// <summary>获取工作表标签颜色。</summary>
        public Color Tab { get; }

        /// <summary>把公式色淡向白色，用作公式单元格的底色。</summary>
        /// <param name="formula">公式字体颜色。</param>
        /// <returns>约 85% 白色和 15% 公式色的混合。</returns>
        public static Color Lighten(Color formula)
        {
            return Color.FromArgb(
                formula.R + ((255 - formula.R) * 85) / 100,
                formula.G + ((255 - formula.G) * 85) / 100,
                formula.B + ((255 - formula.B) * 85) / 100);
        }
    }
}

using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>批量刷新正常结束后的说明窗。计数和明细是白底正文，标题行底色图例用窗体底色。</summary>
    /// <remarks>
    /// 明细为空时白底只包住计数，窗口更矮。图例固定在底部，两行紧凑，文字用较小的灰色斜体，不随明细滚动。
    /// 色块与目标表标题行相同。图例字体由本窗持有，关闭时释放。小窗不记宽高，模态显示，不收集输入。
    /// </remarks>
    internal sealed class RefreshSummaryForm : AppForm
    {
        private const int SwatchRowHeight = 20;

        /// <summary>图例说明用的较小灰色斜体。窗体关闭时释放。</summary>
        private readonly Font _legendFont;

        /// <summary>创建完成说明窗。</summary>
        /// <param name="summary">成功、失败、截断计数。</param>
        /// <param name="details">失败和截断明细。空白时不显示明细区。</param>
        /// <param name="successColor">结果完整时标题行和成功图例的颜色。</param>
        /// <param name="truncatedColor">结果被截断时标题行和截断图例的颜色。</param>
        private RefreshSummaryForm(string summary, string details, Color successColor, Color truncatedColor)
        {
            bool hasDetails = !string.IsNullOrWhiteSpace(details);
            InitializeWindow(
                AppPresentation.WindowTitle("批量刷新"),
                560,
                hasDetails ? 400 : 236,
                new Size(480, hasDetails ? 340 : 210),
                FormStartPosition.CenterParent);
            FormBorderStyle = hasDetails ? FormBorderStyle.Sizable : FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            _legendFont = new Font(
                AppPresentation.DefaultFontName,
                8.25f,
                FontStyle.Italic,
                GraphicsUnit.Point);

            Button okButton = new Button
            {
                Text = "确定",
                AutoSize = true,
                DialogResult = DialogResult.OK
            };
            FlowLayoutPanel actions = CreateActionBar(okButton);

            TableLayoutPanel root = CreateRootTable(
                hasDetails
                    ? new RowStyle(SizeType.Percent, 100)
                    : new RowStyle(SizeType.Absolute, 44),
                new RowStyle(SizeType.AutoSize),
                new RowStyle(SizeType.AutoSize));
            root.Controls.Add(CreateBody(summary, details, hasDetails), 0, 0);
            root.Controls.Add(CreateLegend(_legendFont, successColor, truncatedColor), 0, 1);
            root.Controls.Add(actions, 0, 2);
            Controls.Add(root);

            AcceptButton = okButton;
            CancelButton = okButton;
            Shown += (sender, args) => okButton.Focus();
            FormClosed += (sender, args) => _legendFont.Dispose();
        }

        /// <summary>弹出批量刷新完成说明。</summary>
        /// <param name="owner">所属窗口。</param>
        /// <param name="summary">成功、失败、截断计数。</param>
        /// <param name="details">失败和截断明细。空白时不显示明细区。</param>
        /// <param name="successColor">结果完整时的标题行底色。与状态圆点相同。</param>
        /// <param name="truncatedColor">结果被截断时的标题行底色。与状态圆点相同。</param>
        /// <remarks>在调用方的 UI 线程上以模态方式显示。不因文本为空而拒绝关闭。</remarks>
        public static void Show(
            IWin32Window owner, string summary, string details, Color successColor, Color truncatedColor)
        {
            using (RefreshSummaryForm dialog = new RefreshSummaryForm(summary, details, successColor, truncatedColor))
            {
                dialog.ShowDialog(owner);
            }
        }

        /// <summary>白底正文。没有明细时只包住计数那一行；有明细时文本框去掉边框，并进同一块白底。</summary>
        /// <param name="summary">成功、失败、截断计数。</param>
        /// <param name="details">失败和截断明细。</param>
        /// <param name="hasDetails">是否显示明细。</param>
        /// <returns>带细边框的白底区域。</returns>
        private static Control CreateBody(string summary, string details, bool hasDetails)
        {
            Label summaryLabel = new Label
            {
                Text = summary ?? string.Empty,
                AutoSize = true,
                BackColor = Color.White,
                Dock = DockStyle.Top,
                Margin = new Padding(0, 0, 0, hasDetails ? 6 : 0)
            };
            Panel surface = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10, 8, 10, 8),
                Margin = new Padding(0)
            };
            if (hasDetails)
            {
                surface.Controls.Add(new TextBox
                {
                    Dock = DockStyle.Fill,
                    Multiline = true,
                    ReadOnly = true,
                    HideSelection = true,
                    ScrollBars = ScrollBars.Vertical,
                    WordWrap = true,
                    TabStop = false,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White,
                    Text = details,
                    Margin = new Padding(0)
                });
            }

            surface.Controls.Add(summaryLabel);

            Control frame = LightBorder.Wrap(surface);
            frame.Margin = new Padding(0);
            return frame;
        }

        /// <summary>底部图例用窗体底色，两行色块各占一行字高，不随明细滚动。说明文字用较小的灰色斜体。</summary>
        /// <param name="legendFont">图例文字字体。由窗体持有，关闭时释放。</param>
        /// <param name="successColor">结果完整时的色块。</param>
        /// <param name="truncatedColor">结果被截断时的色块。</param>
        /// <returns>按内容撑开的图例。</returns>
        private static Control CreateLegend(Font legendFont, Color successColor, Color truncatedColor)
        {
            TableLayoutPanel legend = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Margin = new Padding(0, 8, 0, 0),
                Padding = new Padding(0)
            };
            legend.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 20));
            legend.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            legend.RowStyles.Add(new RowStyle(SizeType.Absolute, SwatchRowHeight));
            legend.RowStyles.Add(new RowStyle(SizeType.Absolute, SwatchRowHeight));
            legend.RowStyles.Add(new RowStyle(SizeType.Absolute, SwatchRowHeight));

            Label caption = new Label
            {
                Text = "目标表标题行颜色含义：",
                AutoSize = false,
                Font = legendFont,
                ForeColor = SystemColors.GrayText,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0)
            };
            legend.SetColumnSpan(caption, 2);
            legend.Controls.Add(caption, 0, 0);
            AddSwatchRow(legend, legendFont, 1, successColor, "本次写入成功，结果完整");
            AddSwatchRow(legend, legendFont, 2, truncatedColor, "写入成功，但结果达到「最大导出行数」后被截断");
            return legend;
        }

        /// <summary>图例中的一行：色块在左，含义在右。</summary>
        /// <param name="legend">图例表格。</param>
        /// <param name="legendFont">图例文字字体。由窗体持有，关闭时释放。</param>
        /// <param name="row">行号。</param>
        /// <param name="color">与标题行相同的底色。</param>
        /// <param name="meaning">该底色的含义。</param>
        private static void AddSwatchRow(TableLayoutPanel legend, Font legendFont, int row, Color color, string meaning)
        {
            Panel swatch = new Panel
            {
                BackColor = color,
                Size = new Size(12, 12),
                MinimumSize = new Size(12, 12),
                MaximumSize = new Size(12, 12),
                Margin = new Padding(0),
                Anchor = AnchorStyles.Left
            };
            Label text = new Label
            {
                Text = meaning,
                AutoSize = false,
                Font = legendFont,
                ForeColor = SystemColors.GrayText,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(6, 0, 0, 0)
            };
            legend.Controls.Add(swatch, 0, row);
            legend.Controls.Add(text, 1, row);
        }
    }
}

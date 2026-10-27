using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>列表末项「新增」。点它开始填写一条新记录，本身不是已保存的数据。</summary>
    internal sealed class AddItem
    {
        /// <summary>列表里显示的文字。</summary>
        /// <returns>固定为「新增」。</returns>
        public override string ToString()
        {
            return "新增";
        }
    }

    /// <summary>带「新增」末项的自绘列表，统一处理右键状态、分隔线、末项和行高。</summary>
    /// <remarks>
    /// 非「新增」行的测量和绘制交给 <see cref="EntryMeasure"/> 与 <see cref="EntryDraw"/>，
    /// 让模型设置和 Skill 设置保留各自的图标和缩进。
    /// </remarks>
    internal sealed class AddItemListBox : ListBox
    {
        /// <summary>普通条目的默认行高。</summary>
        public const int DefaultRowHeight = 28;

        /// <summary>「新增」上方分隔带的高度。</summary>
        public const int DefaultSeparatorBand = 8;

        private readonly int _rowHeight;
        private readonly int _separatorBand;
        private readonly Image _addIcon;
        private bool _restoringSelection;

        /// <summary>创建自绘列表并挂好测量、绘制和右键状态处理。</summary>
        /// <param name="addIcon">「新增」左侧的图标。没有时只画文字。</param>
        public AddItemListBox(Image addIcon)
        {
            _rowHeight = DefaultRowHeight;
            _separatorBand = DefaultSeparatorBand;
            _addIcon = addIcon;
            Dock = DockStyle.Fill;
            IntegralHeight = false;
            BorderStyle = BorderStyle.FixedSingle;
            DrawMode = DrawMode.OwnerDrawVariable;
            ItemHeight = _rowHeight;
            MeasureItem += ListMeasureItem;
            DrawItem += ListDrawItem;
        }

        /// <summary>非「新增」行的测量事件。调用方在这里给子型号等特殊行指定高度。</summary>
        public event MeasureItemEventHandler EntryMeasure;

        /// <summary>非「新增」行的绘制事件。调用方在这里画图标、缩进和选中底色。</summary>
        public event DrawItemEventHandler EntryDraw;

        /// <summary>当前这次按下是否为右键。在窗口消息里、选中变化之前写好。</summary>
        public bool RightButton { get; private set; }

        /// <summary>判断一项是否是可直接编辑的普通条目。子型号等只展示行返回 false。</summary>
        /// <remarks>由两个设置窗设置。为 null 时只有「新增」参与退选。</remarks>
        public Func<object, bool> SelectableItem { get; set; }

        /// <summary>为 true 时程序化改选中不更新退选锚点。</summary>
        /// <remarks>窗体批量重建列表或把选中改回配置时设置，避免覆盖用户上次停留的条目。</remarks>
        public bool SuppressSelectionTracking { get; set; }

        /// <summary>右键「新增」时退回的条目序号。没有可退回项时为 -1。</summary>
        public int AnchorIndex { get; set; } = -1;

        /// <summary>先处理「新增」的右键和锚点，再交给窗体处理普通选中。</summary>
        /// <param name="e">事件参数。</param>
        /// <remarks>
        /// 右键落在「新增」上时把选中退回锚点，并且不触发窗体的选中变化，避免把右键当成开始新建。
        /// 左键选中「新增」时记下锚点并正常触发，窗体再清空编辑器。
        /// </remarks>
        protected override void OnSelectedIndexChanged(EventArgs e)
        {
            if (_restoringSelection) return;

            object selectedItem = SelectedItem;
            if (selectedItem is AddItem)
            {
                if (RightButton)
                {
                    _restoringSelection = true;
                    SelectedIndex = RestorableIndex();
                    _restoringSelection = false;
                    return;
                }

                if (!SuppressSelectionTracking)
                    AnchorIndex = SelectedIndex;
            }
            else if (!SuppressSelectionTracking && selectedItem != null
                && SelectableItem != null && SelectableItem(selectedItem))
            {
                AnchorIndex = SelectedIndex;
            }

            base.OnSelectedIndexChanged(e);
        }

        /// <summary>右键没有点在普通条目上时退回的行。</summary>
        /// <returns>可退回的序号；没有时返回 -1。</returns>
        private int RestorableIndex()
        {
            if (AnchorIndex < 0 || AnchorIndex >= Items.Count) return -1;
            object item = Items[AnchorIndex];
            if (item is AddItem) return AnchorIndex;
            return SelectableItem != null && SelectableItem(item) ? AnchorIndex : -1;
        }

        /// <summary>在一行里画出图标和文字。选中时文字用系统高亮色。</summary>
        /// <param name="graphics">列表项的画布。</param>
        /// <param name="font">列表字体。</param>
        /// <param name="bounds">图标和文字所在的区域。分隔线不在这里面。</param>
        /// <param name="text">条目文字。</param>
        /// <param name="icon">左侧图标。没有时只画文字。</param>
        /// <param name="selected">为 true 时文字用系统高亮色。</param>
        /// <param name="textLeft">文字距条目左边的像素。</param>
        /// <param name="iconLeft">图标距条目左边的像素。</param>
        public static void DrawEntry(
            Graphics graphics, Font font, Rectangle bounds, string text, Image icon, bool selected, int textLeft, int iconLeft)
        {
            if (icon != null)
                graphics.DrawImage(icon, bounds.Left + iconLeft, bounds.Top + (bounds.Height - icon.Height) / 2, icon.Width, icon.Height);
            Color color = selected ? SystemColors.HighlightText : SystemColors.ControlText;
            TextRenderer.DrawText(
                graphics,
                text,
                font,
                new Rectangle(bounds.Left + textLeft, bounds.Top, System.Math.Max(0, bounds.Width - textLeft - 4), bounds.Height),
                color,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.Left);
        }

        /// <summary>先记下鼠标键，再交给列表改变选中。</summary>
        /// <param name="m">列表窗口消息。</param>
        protected override void WndProc(ref Message m)
        {
            const int leftDown = 0x0201;
            const int rightDown = 0x0204;
            const int rightUp = 0x0205;
            if (m.Msg == rightDown)
                RightButton = true;
            else if (m.Msg == leftDown)
                RightButton = false;

            base.WndProc(ref m);

            if (m.Msg == rightUp)
                RightButton = false;
        }

        private void ListMeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= Items.Count)
            {
                e.ItemHeight = _rowHeight;
                return;
            }

            if (Items[e.Index] is AddItem)
            {
                e.ItemHeight = _rowHeight + (e.Index > 0 ? _separatorBand : 0);
                return;
            }

            MeasureItemEventHandler handler = EntryMeasure;
            if (handler != null)
                handler(this, e);
            else
                e.ItemHeight = _rowHeight;
        }

        private void ListDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= Items.Count) return;
            if (Items[e.Index] is AddItem)
            {
                DrawAddItem(e);
                return;
            }

            DrawItemEventHandler handler = EntryDraw;
            if (handler != null)
            {
                handler(this, e);
                return;
            }

            e.DrawBackground();
            DrawEntry(e.Graphics, e.Font, e.Bounds, Items[e.Index].ToString(), null, (e.State & DrawItemState.Selected) != 0, 24, 4);
            e.DrawFocusRectangle();
        }

        /// <summary>选中时只高亮文字行。上面已有条目时，分隔线仍用窗口底色。</summary>
        private void DrawAddItem(DrawItemEventArgs e)
        {
            bool selected = (e.State & DrawItemState.Selected) != 0;
            Rectangle content = e.Bounds;
            if (e.Index > 0)
            {
                e.Graphics.FillRectangle(SystemBrushes.Window, new Rectangle(e.Bounds.Left, e.Bounds.Top, e.Bounds.Width, _separatorBand));
                int y = e.Bounds.Top + (_separatorBand / 2);
                e.Graphics.DrawLine(SystemPens.ControlDark, e.Bounds.Left, y, e.Bounds.Right - 1, y);
                content = new Rectangle(e.Bounds.Left, e.Bounds.Top + _separatorBand, e.Bounds.Width, System.Math.Max(0, e.Bounds.Height - _separatorBand));
            }

            e.Graphics.FillRectangle(selected ? SystemBrushes.Highlight : SystemBrushes.Window, content);
            DrawEntry(e.Graphics, e.Font, content, Items[e.Index].ToString(), _addIcon, selected, 24, 4);
            if ((e.State & DrawItemState.Focus) != 0)
                ControlPaint.DrawFocusRectangle(e.Graphics, content);
        }
    }
}

using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>随父容器可用宽度折行并自动撑高的说明标签。</summary>
    /// <remarks>
    /// 放在单列根表格里时按表格内容宽度折行；放在多列表格里时按所在列的宽度折行。
    /// 只限制最大宽度，不改父容器的边距和列宽。宽度在父容器布局后重新计算，不需要调用方额外挂钩子。
    /// </remarks>
    internal sealed class WrapLabel : Label
    {
        private Control _observedParent;

        /// <summary>创建自动折行的说明标签。</summary>
        /// <remarks>高度由文本在最大宽度内折行后决定；调用方只设置字体、颜色、边距和文字。</remarks>
        public WrapLabel()
        {
            AutoSize = true;
        }

        /// <summary>按父容器当前可用宽度重新计算最大宽度。</summary>
        /// <param name="e">父容器变化参数。</param>
        protected override void OnParentChanged(EventArgs e)
        {
            DetachParent();
            base.OnParentChanged(e);
            AttachParent();
            FitWidth();
        }

        /// <summary>文字变化后重新计算折行宽度，避免旧宽度裁掉新内容。</summary>
        /// <param name="e">事件参数。</param>
        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            FitWidth();
        }

        /// <summary>字体变化后重新计算折行宽度，粗体说明也能完整显示。</summary>
        /// <param name="e">事件参数。</param>
        protected override void OnFontChanged(EventArgs e)
        {
            base.OnFontChanged(e);
            FitWidth();
        }

        /// <summary>边距变化后重新计算折行宽度。</summary>
        /// <param name="e">事件参数。</param>
        protected override void OnMarginChanged(EventArgs e)
        {
            base.OnMarginChanged(e);
            FitWidth();
        }

        /// <summary>句柄创建后补算一次宽度，覆盖父容器尚未完成布局的情况。</summary>
        /// <param name="e">事件参数。</param>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            FitWidth();
        }

        /// <summary>释放对父容器事件的订阅。</summary>
        /// <param name="disposing">为 true 时释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                DetachParent();
            base.Dispose(disposing);
        }

        private void AttachParent()
        {
            _observedParent = Parent;
            if (_observedParent == null) return;
            _observedParent.Layout += ParentLayout;
            _observedParent.SizeChanged += ParentSizeChanged;
        }

        private void DetachParent()
        {
            if (_observedParent == null) return;
            _observedParent.Layout -= ParentLayout;
            _observedParent.SizeChanged -= ParentSizeChanged;
            _observedParent = null;
        }

        private void ParentLayout(object sender, LayoutEventArgs e)
        {
            FitWidth();
        }

        private void ParentSizeChanged(object sender, EventArgs e)
        {
            FitWidth();
        }

        private void FitWidth()
        {
            if (IsDisposed || Parent == null) return;
            int width = AvailableWidth();
            if (width <= 0 || MaximumSize.Width == width) return;
            MaximumSize = new Size(width, MaximumSize.Height);
        }

        private int AvailableWidth()
        {
            TableLayoutPanel table = Parent as TableLayoutPanel;
            if (table != null)
            {
                int[] widths = table.GetColumnWidths();
                int column = table.GetColumn(this);
                if (column >= 0 && column < widths.Length)
                {
                    int span = Math.Max(1, table.GetColumnSpan(this));
                    int total = 0;
                    for (int index = column; index < widths.Length && index < column + span; index++)
                        total += widths[index];
                    total -= Margin.Horizontal;
                    if (total > 0) return total;
                }
            }

            return Parent.DisplayRectangle.Width - Margin.Horizontal;
        }
    }
}

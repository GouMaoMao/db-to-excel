using System;
using System.Drawing;
using System.Windows.Forms;
using DB2Sheet.Excel;

namespace DB2Sheet.UI
{
    /// <summary>在写入查询结果前收集目标工作表名称，以及是否把 SQL 写入 SQL 页。</summary>
    /// <remarks>只返回用户的选择，不修改工作簿。不记住宽高。</remarks>
    internal sealed class ExportSheetDialog : AppForm
    {
        private readonly TextBox _sheetName;
        private readonly CheckBox _writeSql;

        /// <summary>创建写入确认窗。</summary>
        /// <param name="initialSheetName">目标表名输入框的初始文字，通常是当前查询方案名。</param>
        private ExportSheetDialog(string initialSheetName)
        {
            Text = "写入 Sheet";
            Width = 480;
            Height = 230;
            MinimumSize = new Size(420, 210);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            Label prompt = new Label
            {
                Text = "请输入目标 Sheet 名称",
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8)
            };
            _sheetName = new TextBox
            {
                Dock = DockStyle.Top,
                Text = initialSheetName ?? string.Empty
            };
            _writeSql = new CheckBox
            {
                Text = "同时把SQL代码保存在Excel中",
                AutoSize = true,
                Margin = new Padding(0, 12, 0, 0),
                Checked = false
            };

            Button okButton = new Button { Text = "确定", AutoSize = true };
            Button cancelButton = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };
            okButton.Click += OkButtonClick;
            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft
            };
            actions.Controls.Add(cancelButton);
            actions.Controls.Add(okButton);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                RowCount = 4,
                ColumnCount = 1
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(prompt, 0, 0);
            layout.Controls.Add(_sheetName, 0, 1);
            layout.Controls.Add(_writeSql, 0, 2);
            layout.Controls.Add(actions, 0, 3);
            Controls.Add(layout);

            AcceptButton = okButton;
            CancelButton = cancelButton;
            Shown += (sender, args) =>
            {
                _sheetName.SelectionStart = 0;
                _sheetName.SelectionLength = _sheetName.TextLength;
                _sheetName.Focus();
            };
        }

        /// <summary>获取用户确认的目标工作表名称。</summary>
        public string SheetName => _sheetName.Text.Trim();

        /// <summary>获取是否把当前 SQL 写入 SQL 工作表。</summary>
        public bool WriteSql => _writeSql.Checked;

        /// <summary>显示写入确认窗。</summary>
        /// <param name="owner">所属窗口。</param>
        /// <param name="initialSheetName">目标表名的初始文字。</param>
        /// <param name="sheetName">用户确认的表名。取消时保持初始文字。</param>
        /// <param name="writeSql">用户是否勾选写入 SQL 页。取消时为 false。</param>
        /// <returns>用户确认且表名可作为工作表名时为 true。</returns>
        public static bool TryShow(
            IWin32Window owner,
            string initialSheetName,
            out string sheetName,
            out bool writeSql)
        {
            sheetName = initialSheetName ?? string.Empty;
            writeSql = false;
            using (ExportSheetDialog dialog = new ExportSheetDialog(initialSheetName))
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK)
                    return false;

                sheetName = dialog.SheetName;
                writeSql = dialog.WriteSql;
                return true;
            }
        }

        /// <summary>表名能作为结果工作表时才关闭。SQL 页本身的名称留给批量刷新。</summary>
        private void OkButtonClick(object sender, EventArgs e)
        {
            string error = ValidateSheetName(SheetName);
            if (error != null)
            {
                MessageBox.Show(this, error, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _sheetName.Focus();
                _sheetName.SelectAll();
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>检查目标表名是否能直接用作 Excel 工作表名。</summary>
        /// <param name="name">已去掉首尾空白的名称。</param>
        /// <returns>可以写入时为空。否则是留在窗体上显示的说明。</returns>
        private static string ValidateSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "请输入目标 Sheet 名称。";
            if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
                return "目标 Sheet 名称不能包含下列字符：: \\ / ? * [ ]";
            if (name.Length > 31)
                return "目标 Sheet 名称不能超过 31 个字符。";
            if (string.Equals(name, SqlSheetTaskReader.SqlSheetName, StringComparison.OrdinalIgnoreCase))
                return "目标 Sheet 不能命名为 SQL，这张表要留给批量刷新。";
            return null;
        }
    }
}

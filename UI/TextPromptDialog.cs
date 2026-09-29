using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>提供单行文本输入的通用弹窗，用于收集方案名或目标 Sheet 名称。</summary>
    /// <remarks>该窗体仅返回文本，不直接执行业务保存或写入；调用方负责后续校验和副作用处理。</remarks>
    internal sealed class TextPromptDialog : AppForm
    {
        private readonly TextBox _valueTextBox;

        /// <summary>创建文本输入弹窗。</summary>
        /// <param name="title">窗口标题。</param>
        /// <param name="label">输入框标签。</param>
        /// <param name="value">初始值。</param>
        private TextPromptDialog(string title, string label, string value)
        {
            Text = title;
            Width = 480;
            Height = 180;
            MinimumSize = new Size(420, 170);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            Label promptLabel = new Label
            {
                Text = label,
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8)
            };

            _valueTextBox = new TextBox { Dock = DockStyle.Top, Text = value ?? string.Empty };

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
                RowCount = 3,
                ColumnCount = 1
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(promptLabel, 0, 0);
            layout.Controls.Add(_valueTextBox, 0, 1);
            layout.Controls.Add(actions, 0, 2);
            Controls.Add(layout);

            AcceptButton = okButton;
            CancelButton = cancelButton;
            Shown += (sender, args) =>
            {
                _valueTextBox.SelectionStart = 0;
                _valueTextBox.SelectionLength = _valueTextBox.TextLength;
                _valueTextBox.Focus();
            };
        }

        /// <summary>获取用户确认后的输入值。</summary>
        public string Value => _valueTextBox.Text.Trim();

        /// <summary>显示文本输入弹窗并输出用户确认后的非空值。</summary>
        /// <param name="owner">所属窗口。</param>
        /// <param name="title">窗口标题。</param>
        /// <param name="label">输入框标签。</param>
        /// <param name="initialValue">输入框初始值。</param>
        /// <param name="value">用户确认的输入值；取消时保持原值。</param>
        /// <returns>用户确认且输入非空时为 <see langword="true"/>。</returns>
        public static bool TryShow(
            IWin32Window owner,
            string title,
            string label,
            string initialValue,
            out string value)
        {
            value = initialValue ?? string.Empty;
            using (TextPromptDialog dialog = new TextPromptDialog(title, label, initialValue))
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK)
                    return false;

                value = dialog.Value;
                return !string.IsNullOrWhiteSpace(value);
            }
        }

        private void OkButtonClick(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Value))
            {
                MessageBox.Show(this, "请输入内容。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _valueTextBox.Focus();
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}

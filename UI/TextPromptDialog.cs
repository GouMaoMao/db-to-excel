using System;
using System.Drawing;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>收集单行输入，或只读展示多行说明。</summary>
    /// <remarks>
    /// 输入弹窗只返回文本和可选的附加勾选，不直接执行业务保存或写入。
    /// 只读说明不收集输入。调用方负责业务校验，校验失败时对话窗留在原地。
    /// </remarks>
    internal sealed class TextPromptDialog : AppForm
    {
        private readonly TextBox _valueTextBox;
        private readonly CheckBox _checkBox;
        private readonly Func<string, string> _validate;

        /// <summary>创建文本输入弹窗。</summary>
        /// <param name="title">完整窗口标题。</param>
        /// <param name="label">输入框标签。</param>
        /// <param name="value">初始值。</param>
        /// <param name="checkText">附加勾选的文字。为空时不显示勾选框。</param>
        /// <param name="validate">确认前的校验。返回空表示通过，否则返回要提示的说明。</param>
        private TextPromptDialog(string title, string label, string value, string checkText, Func<string, string> validate)
        {
            bool hasCheck = !string.IsNullOrEmpty(checkText);
            InitializeWindow(
                title,
                480,
                hasCheck ? 230 : 180,
                new Size(420, hasCheck ? 210 : 170),
                FormStartPosition.CenterParent);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            _validate = validate;

            Label promptLabel = new WrapLabel
            {
                Text = label,
                Margin = new Padding(0, 0, 0, 8)
            };
            _valueTextBox = new TextBox { Dock = DockStyle.Top, Text = value ?? string.Empty };
            _checkBox = hasCheck
                ? new CheckBox
                {
                    Text = checkText,
                    AutoSize = true,
                    Margin = new Padding(0, 12, 0, 0),
                    Checked = false
                }
                : null;

            Button okButton = new Button { Text = "确定", AutoSize = true };
            Button cancelButton = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };
            okButton.Click += OkButtonClick;
            FlowLayoutPanel actions = CreateActionBar(cancelButton, okButton);

            TableLayoutPanel root = hasCheck
                ? CreateRootTable(
                    new RowStyle(SizeType.AutoSize),
                    new RowStyle(SizeType.AutoSize),
                    new RowStyle(SizeType.AutoSize),
                    new RowStyle(SizeType.Percent, 100),
                    new RowStyle(SizeType.AutoSize))
                : CreateRootTable(
                    new RowStyle(SizeType.AutoSize),
                    new RowStyle(SizeType.AutoSize),
                    new RowStyle(SizeType.Percent, 100),
                    new RowStyle(SizeType.AutoSize));
            root.Controls.Add(promptLabel, 0, 0);
            root.Controls.Add(_valueTextBox, 0, 1);
            if (hasCheck)
            {
                root.Controls.Add(_checkBox, 0, 2);
                root.Controls.Add(actions, 0, 4);
            }
            else
            {
                root.Controls.Add(actions, 0, 3);
            }

            Controls.Add(root);
            AcceptButton = okButton;
            CancelButton = cancelButton;
            Shown += (sender, args) =>
            {
                _valueTextBox.SelectionStart = 0;
                _valueTextBox.SelectionLength = _valueTextBox.TextLength;
                _valueTextBox.Focus();
            };
        }

        /// <summary>创建只读多行说明窗。不收集输入。</summary>
        /// <remarks>焦点停在「确定」上。文本框失焦后不保留蓝色选区，避免一打开整段被选中。</remarks>
        /// <param name="title">完整窗口标题。</param>
        /// <param name="text">要展示的多行文本。</param>
        private TextPromptDialog(string title, string text)
        {
            InitializeWindow(title, 560, 360, new Size(420, 240), FormStartPosition.CenterParent);
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;

            _valueTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                HideSelection = true,
                ScrollBars = ScrollBars.Vertical,
                Text = text ?? string.Empty,
                WordWrap = true
            };
            Button okButton = new Button { Text = "确定", AutoSize = true, DialogResult = DialogResult.OK };
            FlowLayoutPanel actions = CreateActionBar(okButton);

            TableLayoutPanel root = CreateRootTable(
                new RowStyle(SizeType.Percent, 100),
                new RowStyle(SizeType.AutoSize));
            root.Controls.Add(_valueTextBox, 0, 0);
            root.Controls.Add(actions, 0, 1);
            Controls.Add(root);

            AcceptButton = okButton;
            CancelButton = okButton;
            Shown += (sender, args) => okButton.Focus();
        }

        /// <summary>获取用户确认后的输入值。</summary>
        public string Value => _valueTextBox.Text.Trim();

        /// <summary>获取用户是否勾选了附加项。没有勾选框时恒为 false。</summary>
        public bool Checked => _checkBox != null && _checkBox.Checked;

        /// <summary>显示只读多行说明，用户看完后关闭。</summary>
        /// <param name="owner">所属窗口。</param>
        /// <param name="title">窗口标题。</param>
        /// <param name="text">要展示的多行文本。</param>
        /// <remarks>不收集输入，也不因文本为空而拒绝关闭。</remarks>
        public static void ShowText(IWin32Window owner, string title, string text)
        {
            using (TextPromptDialog dialog = new TextPromptDialog(title, text))
            {
                dialog.ShowDialog(owner);
            }
        }

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
            return TryShow(owner, title, label, initialValue, null, null, out value, out bool ignored);
        }

        /// <summary>显示带可选附加勾选和业务校验的文本输入弹窗。</summary>
        /// <param name="owner">所属窗口。</param>
        /// <param name="title">窗口标题。</param>
        /// <param name="label">输入框标签。</param>
        /// <param name="initialValue">输入框初始值。</param>
        /// <param name="checkText">附加勾选的文字。为空时不显示勾选框。</param>
        /// <param name="validate">确认前的校验。返回空表示通过，否则返回要提示的说明。</param>
        /// <param name="value">用户确认的输入值；取消时保持原值。</param>
        /// <param name="check">用户是否勾选附加项。取消时为 false。</param>
        /// <returns>用户确认且校验通过时为 <see langword="true"/>。</returns>
        public static bool TryShow(
            IWin32Window owner,
            string title,
            string label,
            string initialValue,
            string checkText,
            Func<string, string> validate,
            out string value,
            out bool check)
        {
            value = initialValue ?? string.Empty;
            check = false;
            using (TextPromptDialog dialog = new TextPromptDialog(title, label, initialValue, checkText, validate))
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK)
                    return false;

                value = dialog.Value;
                check = dialog.Checked;
                return !string.IsNullOrWhiteSpace(value);
            }
        }

        private void OkButtonClick(object sender, EventArgs e)
        {
            string error = _validate != null
                ? _validate(Value)
                : (string.IsNullOrWhiteSpace(Value) ? "请输入内容。" : null);
            if (error != null)
            {
                MessageBox.Show(this, error, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _valueTextBox.Focus();
                _valueTextBox.SelectAll();
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}

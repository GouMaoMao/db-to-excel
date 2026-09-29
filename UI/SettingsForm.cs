using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.UI
{
    /// <summary>根据设置注册表动态生成分类编辑界面，并批量保存用户设置。</summary>
    /// <remarks>保存成功时设置 <see cref="Form.DialogResult"/> 为 OK；恢复默认值需要用户确认。</remarks>
    public sealed class SettingsForm : AppForm
    {
        private readonly ISettingsRegistry _registry;
        private readonly ISettingsStore _store;
        private readonly Dictionary<string, SettingEditor> _editors =
            new Dictionary<string, SettingEditor>(StringComparer.OrdinalIgnoreCase);
        private readonly Label _status;

        /// <summary>创建设置窗体。</summary>
        /// <param name="registry">提供设置类型、分类、默认值和校验规则。</param>
        /// <param name="store">读取及持久化设置值。</param>
        public SettingsForm(ISettingsRegistry registry, ISettingsStore store)
        {
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
            _store = store ?? throw new ArgumentNullException(nameof(store));

            Text = AppPresentation.WindowTitle("设置");
            Width = 660;
            Height = 600;
            MinimumSize = new Size(520, 420);
            StartPosition = FormStartPosition.CenterParent;

            TabControl categories = new TabControl { Dock = DockStyle.Fill };
            _status = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Padding = new Padding(0, 6, 0, 6)
            };
            Button resetButton = new Button { Text = "恢复默认值", AutoSize = true };
            Button saveButton = new Button { Text = "保存", AutoSize = true };
            Button cancelButton = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };

            BuildCategories(categories);

            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft
            };
            actions.Controls.Add(cancelButton);
            actions.Controls.Add(saveButton);
            actions.Controls.Add(resetButton);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 3
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(categories, 0, 0);
            root.Controls.Add(_status, 0, 1);
            root.Controls.Add(actions, 0, 2);
            Controls.Add(root);

            resetButton.Click += ResetButtonClick;
            saveButton.Click += SaveButtonClick;
            AcceptButton = saveButton;
            CancelButton = cancelButton;
        }

        private void BuildCategories(TabControl categories)
        {
            foreach (IGrouping<string, SettingDefinition> category in _registry.GetAll()
                .Where(definition => !definition.Key.StartsWith("session.", StringComparison.OrdinalIgnoreCase))
                .GroupBy(definition => string.IsNullOrWhiteSpace(definition.Category) ? "常规" : definition.Category))
            {
                TabPage page = new TabPage(category.Key) { Padding = new Padding(8) };
                TableLayoutPanel table = new TableLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    ColumnCount = 2
                };
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 210));
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

                foreach (SettingDefinition definition in category)
                {
                    int row = table.RowCount++;
                    table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    Label label = new Label
                    {
                        Text = definition.DisplayName + (definition.RequiresRestart ? "（需重启）" : string.Empty),
                        AutoSize = true,
                        Anchor = AnchorStyles.Left,
                        Margin = new Padding(3, 9, 8, 3)
                    };
                    SettingEditor editor = CreateEditor(definition);
                    editor.SetValue(_store.Get(definition));
                    _editors.Add(definition.Key, editor);
                    table.Controls.Add(label, 0, row);
                    table.Controls.Add(editor.Control, 1, row);
                }

                page.Controls.Add(table);
                categories.TabPages.Add(page);
            }
        }

        private static SettingEditor CreateEditor(SettingDefinition definition)
        {
            if (definition.ValueType == typeof(bool))
            {
                CheckBox checkBox = new CheckBox { AutoSize = true, Anchor = AnchorStyles.Left };
                return new SettingEditor(
                    definition,
                    checkBox,
                    value => checkBox.Checked = value is bool state && state,
                    () => checkBox.Checked);
            }

            if (definition.ValueType.IsEnum)
            {
                ComboBox choice = new ComboBox
                {
                    Dock = DockStyle.Fill,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                choice.Items.AddRange(Enum.GetValues(definition.ValueType).Cast<object>().ToArray());
                return new SettingEditor(
                    definition,
                    choice,
                    value => choice.SelectedItem = value,
                    () => choice.SelectedItem);
            }

            TextBox textBox = new TextBox { Dock = DockStyle.Fill };
            return new SettingEditor(
                definition,
                textBox,
                value => textBox.Text = definition.Serialize(value),
                () =>
                {
                    if (!definition.TryParse(textBox.Text.Trim(), out object parsed))
                    {
                        throw new FormatException("“" + definition.DisplayName + "”的值无效。");
                    }
                    return parsed;
                });
        }

        private void ResetButtonClick(object sender, EventArgs e)
        {
            DialogResult confirmation = MessageBox.Show(
                this,
                "将界面中的全部设置恢复为默认值？保存前不会写入文件。",
                "恢复默认值",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            if (confirmation != DialogResult.Yes) return;

            foreach (SettingEditor editor in _editors.Values)
            {
                editor.SetValue(editor.Definition.DefaultValue);
            }
            _status.ForeColor = SystemColors.GrayText;
            _status.Text = "已恢复默认值，请点击“保存”应用。";
        }

        private void SaveButtonClick(object sender, EventArgs e)
        {
            try
            {
                Dictionary<string, object> values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (KeyValuePair<string, SettingEditor> pair in _editors)
                {
                    object value = pair.Value.GetValue();
                    if (!pair.Value.Definition.IsValid(value))
                    {
                        throw new FormatException("“" + pair.Value.Definition.DisplayName + "”的值不符合约束。");
                    }
                    values.Add(pair.Key, value);
                }

                _store.SetMany(values);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception) when (exception is FormatException || exception is ArgumentException)
            {
                _status.ForeColor = Color.DarkRed;
                _status.Text = exception.Message;
                MessageBox.Show(this, exception.Message, "设置", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>把设置定义与其编辑控件、取值和赋值委托关联起来。</summary>
        private sealed class SettingEditor
        {
            private readonly Action<object> _setValue;
            private readonly Func<object> _getValue;

            public SettingEditor(
                SettingDefinition definition,
                Control control,
                Action<object> setValue,
                Func<object> getValue)
            {
                Definition = definition;
                Control = control;
                _setValue = setValue;
                _getValue = getValue;
            }

            public SettingDefinition Definition { get; }
            public Control Control { get; }
            public void SetValue(object value) => _setValue(value);
            public object GetValue() => _getValue();
        }
    }
}

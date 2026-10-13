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
    /// <remarks>
    /// 保存成功时设置 <see cref="Form.DialogResult"/> 为 OK；恢复默认值需要用户确认。框下说明来自 <see cref="AppPresentation.SettingNote"/>，随窗体宽度折行。
    /// 分类页内容高出可见区域时出现纵向滚动条，滚轮在该页上时滚动这一页。
    /// </remarks>
    public sealed class SettingsForm : AppForm
    {
        private readonly ISettingsRegistry _registry;
        private readonly ISettingsStore _store;
        private readonly Dictionary<string, SettingEditor> _editors =
            new Dictionary<string, SettingEditor>(StringComparer.OrdinalIgnoreCase);
        private readonly Label _status;
        private readonly List<Label> _notes = new List<Label>();
        private readonly Dictionary<TabPage, Panel> _categoryScrolls = new Dictionary<TabPage, Panel>();
        private readonly WheelRedirector _wheelRedirector;
        private TabControl _categories;

        /// <summary>创建设置窗体。</summary>
        /// <param name="registry">提供设置类型、分类、默认值和校验规则。</param>
        /// <param name="store">读取及持久化设置值。</param>
        /// <param name="initialCategory">打开后选中的分类。为空或找不到时停在第一页。</param>
        public SettingsForm(ISettingsRegistry registry, ISettingsStore store, string initialCategory = null)
        {
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
            _store = store ?? throw new ArgumentNullException(nameof(store));

            Text = AppPresentation.WindowTitle("设置");
            Width = 660;
            Height = 600;
            MinimumSize = new Size(520, 420);
            StartPosition = FormStartPosition.Manual;
            FormSizeMemory.Attach(this, _store, "Settings");

            _categories = new TabControl { Dock = DockStyle.Fill };
            _wheelRedirector = new WheelRedirector(this);
            Application.AddMessageFilter(_wheelRedirector);
            Disposed += (sender, args) => Application.RemoveMessageFilter(_wheelRedirector);
            _status = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Padding = new Padding(0, 6, 0, 6)
            };
            Button resetButton = new Button { Text = "恢复默认值", AutoSize = true };
            Button saveButton = new Button { Text = "保存", AutoSize = true };
            Button cancelButton = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };

            BuildCategories(_categories, initialCategory);

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
            root.Controls.Add(_categories, 0, 0);
            root.Controls.Add(_status, 0, 1);
            root.Controls.Add(actions, 0, 2);
            Controls.Add(root);

            resetButton.Click += ResetButtonClick;
            saveButton.Click += SaveButtonClick;
            Resize += (sender, args) => ApplyNoteWidths();
            Shown += (sender, args) => ApplyNoteWidths();
            AcceptButton = saveButton;
            CancelButton = cancelButton;
            ApplyNoteWidths();
        }

        private void BuildCategories(TabControl categories, string initialCategory)
        {
            foreach (IGrouping<string, SettingDefinition> category in _registry.GetAll()
                .Where(definition => !definition.Key.StartsWith("session.", StringComparison.OrdinalIgnoreCase))
                .GroupBy(definition => string.IsNullOrWhiteSpace(definition.Category) ? "常规" : definition.Category))
            {
                TableLayoutPanel table = new TableLayoutPanel
                {
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    Margin = new Padding(0),
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
                        Margin = new Padding(3, 4, 8, 0)
                    };
                    SettingEditor editor = CreateEditor(definition);
                    editor.SetValue(_store.Get(definition));
                    _editors.Add(definition.Key, editor);
                    table.Controls.Add(label, 0, row);
                    table.Controls.Add(editor.Control, 1, row);
                    AddSettingNote(table, definition.Key);
                }

                categories.TabPages.Add(CreateCategoryPage(category.Key, table));
            }

            if (string.IsNullOrEmpty(initialCategory)) return;
            foreach (TabPage page in categories.TabPages)
            {
                if (!string.Equals(page.Text, initialCategory, StringComparison.Ordinal)) continue;
                categories.SelectedTab = page;
                return;
            }
        }

        /// <summary>把分类参数表放进可纵向滚动的页。内容未超出可见高度时不显示滚动条。</summary>
        /// <param name="title">分类名称，同时作为页签文字。</param>
        /// <param name="table">该分类的参数表。宽度随后跟滚动区域对齐。</param>
        /// <returns>已加入滚动容器的页。</returns>
        private TabPage CreateCategoryPage(string title, TableLayoutPanel table)
        {
            TabPage page = new TabPage(title) { Padding = new Padding(8, 8, 0, 8) };
            Panel scroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };
            table.Location = Point.Empty;
            scroll.Controls.Add(table);
            scroll.Resize += (sender, args) =>
            {
                FitCategoryTable(scroll, table);
                ApplyNoteWidths();
            };
            page.Controls.Add(scroll);
            _categoryScrolls.Add(page, scroll);
            return page;
        }

        /// <summary>让参数表与滚动区域同宽。即将超出时先留出滚动条宽度，避免出现滚动条后宽度再跳一次。</summary>
        /// <param name="scroll">分类页里的滚动容器。</param>
        /// <param name="table">其中的参数表。</param>
        private static void FitCategoryTable(Panel scroll, TableLayoutPanel table)
        {
            int width = scroll.ClientSize.Width;
            if (width <= 1) return;
            int fitted = width;
            if (!scroll.VerticalScroll.Visible
                && table.GetPreferredSize(new Size(width, 0)).Height > scroll.ClientSize.Height)
            {
                fitted = Math.Max(1, width - SystemInformation.VerticalScrollBarWidth);
            }

            if (table.MaximumSize.Width == fitted && table.Width == fitted) return;
            table.MinimumSize = new Size(fitted, 0);
            table.MaximumSize = new Size(fitted, 0);
        }

        /// <summary>指针位于当前分类页时，用滚轮滚动该页。</summary>
        /// <param name="message">尚未分发的滚轮消息。</param>
        /// <returns>已经滚动时为 true，消息不再交给当前焦点控件。</returns>
        private bool ScrollAtCursor(Message message)
        {
            if (_categories == null || IsDisposed) return false;
            TabPage page = _categories.SelectedTab;
            if (page == null || !_categoryScrolls.TryGetValue(page, out Panel scroll) || scroll.IsDisposed)
                return false;
            if (!scroll.VerticalScroll.Visible) return false;
            Point cursor = scroll.PointToClient(Cursor.Position);
            if (!scroll.ClientRectangle.Contains(cursor)) return false;

            int delta = (short)((message.WParam.ToInt64() >> 16) & 0xFFFF);
            int next = -scroll.AutoScrollPosition.Y - delta;
            if (next < 0) next = 0;
            scroll.AutoScrollPosition = new Point(0, next);
            return true;
        }

        /// <summary>在当前设置项下方追加灰色说明。没有说明时不占行。</summary>
        /// <param name="table">该分类的参数表。</param>
        /// <param name="key">设置键，对应 <see cref="AppPresentation.SettingNote"/>。</param>
        private void AddSettingNote(TableLayoutPanel table, string key)
        {
            string text = AppPresentation.SettingNote(key);
            if (string.IsNullOrEmpty(text)) return;

            Label note = new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(3, 0, 8, 4)
            };
            _notes.Add(note);
            int row = table.RowCount++;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(note, 1, row);
        }

        /// <summary>把框下说明收进所在参数表的输入列，避免长句横向撑出窗体。</summary>
        private void ApplyNoteWidths()
        {
            foreach (Label note in _notes)
            {
                if (note.IsDisposed) continue;
                int width = NoteWidth(note.Parent as TableLayoutPanel);
                if (note.MaximumSize.Width == width) continue;
                note.MaximumSize = new Size(width, 0);
            }
        }

        /// <summary>说明放在输入列。列宽尚未算出时按窗体宽度减去名称列估算。</summary>
        /// <param name="table">说明所在的参数表。</param>
        /// <returns>说明标签的最大宽度。</returns>
        private int NoteWidth(TableLayoutPanel table)
        {
            if (table != null)
            {
                int[] columns = table.GetColumnWidths();
                if (columns != null && columns.Length > 1 && columns[1] > 40)
                {
                    return Math.Max(120, columns[1] - 8);
                }
            }

            return Math.Max(120, ClientSize.Width - 270);
        }

        private static SettingEditor CreateEditor(SettingDefinition definition)
        {
            if (definition.ValueType == typeof(Color))
                return CreateColorEditor(definition);

            if (definition.ValueType == typeof(bool))
            {
                CheckBox checkBox = new CheckBox
                {
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 2, 0, 0)
                };
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

        /// <summary>用色块显示颜色。旁边的按钮打开系统调色板，选定后只改界面，要点保存才写入。</summary>
        /// <param name="definition">颜色设置。取值和赋值都用 <see cref="Color"/>。</param>
        /// <returns>色块和选择按钮组成的编辑器。</returns>
        private static SettingEditor CreateColorEditor(SettingDefinition definition)
        {
            Panel chip = new Panel
            {
                Width = 36,
                Height = 22,
                Margin = new Padding(0, 1, 6, 0),
                BorderStyle = BorderStyle.FixedSingle
            };
            Button pick = new Button { Text = "选择", AutoSize = true, Margin = new Padding(0) };
            pick.Click += (sender, args) =>
            {
                using (ColorDialog dialog = new ColorDialog { FullOpen = true, Color = chip.BackColor })
                {
                    if (dialog.ShowDialog(pick.FindForm()) != DialogResult.OK) return;
                    chip.BackColor = Color.FromArgb(dialog.Color.R, dialog.Color.G, dialog.Color.B);
                }
            };

            FlowLayoutPanel row = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            row.Controls.Add(chip);
            row.Controls.Add(pick);
            return new SettingEditor(
                definition,
                row,
                value =>
                {
                    Color color = value is Color chosen ? chosen : Color.Black;
                    chip.BackColor = Color.FromArgb(color.R, color.G, color.B);
                },
                () => Color.FromArgb(chip.BackColor.R, chip.BackColor.G, chip.BackColor.B));
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

        /// <summary>设置窗打开期间，把落在分类页上的滚轮滚动交给该页。</summary>
        private sealed class WheelRedirector : IMessageFilter
        {
            private const int MouseWheelMessage = 0x020A;
            private readonly SettingsForm _form;

            /// <summary>创建只服务于这一张设置窗的滚轮过滤器。</summary>
            /// <param name="form">要滚动的设置窗。窗体释放后由外部移除过滤器。</param>
            public WheelRedirector(SettingsForm form)
            {
                _form = form;
            }

            /// <inheritdoc/>
            public bool PreFilterMessage(ref Message message)
            {
                if (message.Msg != MouseWheelMessage) return false;
                return _form.ScrollAtCursor(message);
            }
        }
    }
}

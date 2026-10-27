using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;

namespace DB2Sheet.UI
{
    /// <summary>导入和编辑 Skill。正文只作为发给模型的上下文，不会执行，也不会打开其中的链接。</summary>
    /// <remarks>
    /// 左侧列表末项是「新增」。上面已有 Skill 时，该项顶部有一条分隔线。点中后这一行保持选中高亮。
    /// 删除在某一条的右键菜单里。右侧上方是名称，下方是内容。
    /// </remarks>
    public sealed class AiSkillSettingsForm : AppForm
    {
        private readonly AiAssistant _assistant;
        private readonly Image _addIcon = DatabaseTreeImageCatalog.LoadActionIcon("icon_add.png");
        private readonly Image _skillIcon = AiModelIcons.Skill(true);
        private readonly ToolTip _tips = new ToolTip();
        private readonly ContextMenuStrip _listMenu = new ContextMenuStrip();
        private readonly AddItem _addItem = new AddItem();
        private readonly AddItemListBox _list;
        private readonly TextBox _name = new TextBox { BorderStyle = BorderStyle.FixedSingle };
        private readonly TextBox _body = new TextBox
        {
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            WordWrap = true,
            BorderStyle = BorderStyle.FixedSingle,
            AcceptsReturn = true
        };
        private bool _binding;
        private string _listTip = string.Empty;
        private string _loadedId;

        /// <summary>程序化改编辑器和选中时，同时暂停列表的退选锚点跟踪。</summary>
        private bool Binding
        {
            get { return _binding; }
            set
            {
                _binding = value;
                if (_list != null)
                    _list.SuppressSelectionTracking = value;
            }
        }

        /// <summary>创建 Skill 设置窗体。</summary>
        /// <param name="assistant">Skill 存储。</param>
        /// <param name="settings">记住窗体尺寸。</param>
        public AiSkillSettingsForm(AiAssistant assistant, ISettingsStore settings)
        {
            _assistant = assistant ?? throw new ArgumentNullException(nameof(assistant));
            InitializeWindow(
                AppPresentation.WindowTitle("Skill 设置"),
                760,
                560,
                new Size(640, 520),
                settings,
                "AiSkills");

            _list = new AddItemListBox(_addIcon);
            _list.SelectableItem = item => item is SkillItem;
            _list.EntryDraw += ListDrawItem;
            _list.MouseMove += ListMouseMove;
            _list.MouseLeave += ListMouseLeave;
            _list.MouseUp += ListMouseUp;
            _list.SelectedIndexChanged += ListSelectedIndexChanged;

            ToolStripMenuItem deleteItem = new ToolStripMenuItem("删除", null, (sender, args) => DeleteSelected());
            deleteItem.ToolTipText = AppPresentation.AiSkillDeleteTip;
            _listMenu.Items.Add(deleteItem);

            Label help = new WrapLabel
            {
                Margin = new Padding(0),
                ForeColor = SystemColors.GrayText,
                Text = AppPresentation.AiSkillHelp,
                Padding = new Padding(0, 0, 0, 8)
            };

            Button importButton = ButtonOf("导入");
            Button saveButton = ButtonOf("保存");
            Button closeButton = ButtonOf("关闭");
            closeButton.DialogResult = DialogResult.Cancel;
            importButton.Click += (sender, args) => Import();
            saveButton.Click += (sender, args) => SaveSelected();

            FlowLayoutPanel actions = CreateActionBar(closeButton, saveButton, importButton);

            TableLayoutPanel editor = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(12, 0, 0, 0)
            };
            editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            editor.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            editor.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _name.Dock = DockStyle.Fill;
            _name.Margin = new Padding(0, 0, 0, 8);
            _body.Dock = DockStyle.Fill;
            editor.Controls.Add(Caption("名称"), 0, 0);
            editor.Controls.Add(_name, 0, 1);
            editor.Controls.Add(Caption("内容"), 0, 2);
            editor.Controls.Add(_body, 0, 3);

            SplitContainer split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6
            };
            split.Panel1.Controls.Add(_list);
            split.Panel2.Controls.Add(editor);

            TableLayoutPanel root = CreateRootTable(
                new RowStyle(SizeType.AutoSize),
                new RowStyle(SizeType.Percent, 100),
                new RowStyle(SizeType.AutoSize));
            root.Controls.Add(help, 0, 0);
            root.Controls.Add(split, 0, 1);
            root.Controls.Add(actions, 0, 2);
            Controls.Add(root);
            CancelButton = closeButton;
            Load += (sender, args) =>
            {
                int available = split.Width - split.SplitterWidth;
                if (available >= 360)
                {
                    try { split.SplitterDistance = 200; }
                    catch (InvalidOperationException)
                    {
                        // 句柄还没按最终宽度布局时跳过，下次打开会再摆。
                    }
                }

                Reload(null);
            };
            FormClosed += (sender, args) =>
            {
                _addIcon.Dispose();
                _tips.Dispose();
                _listMenu.Dispose();
            };
        }

        /// <summary>选中仍停在「新增」时才清空右侧。中间若已改选 Skill，则不打断。</summary>
        private void BeginNewIfAddRow()
        {
            if (IsDisposed || !(_list.SelectedItem is AddItem)) return;
            BeginNew();
        }

        /// <summary>清空右侧，开始填写一条还没保存的 Skill，并让「新增」保持选中。</summary>
        private void BeginNew()
        {
            Binding = true;
            int addIndex = _list.Items.Count - 1;
            if (addIndex >= 0 && _list.Items[addIndex] is AddItem)
            {
                _list.SelectedIndex = addIndex;
                _list.AnchorIndex = addIndex;
            }

            _loadedId = null;
            _name.Text = string.Empty;
            _body.Text = string.Empty;
            Binding = false;
            _name.Focus();
        }

        private void Import()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "文本和 Markdown|*.txt;*.md|所有文件|*.*";
                dialog.Title = "导入 Skill";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    string body = File.ReadAllText(dialog.FileName, Encoding.UTF8);
                    string name = Path.GetFileNameWithoutExtension(dialog.FileName);
                    AiSkill saved = _assistant.Skills.Save(new AiSkill { Name = name, Body = body });
                    _assistant.WriteAction("skill.import", "name=" + saved.Name + "\r\n" + saved.Body, null, null);
                    Reload(saved.Id);
                }
                catch (Exception exception)
                {
                    ExceptionDetailForm.Show(this, exception);
                }
            }
        }

        private void SaveSelected()
        {
            try
            {
                AiSkill saved = _assistant.Skills.Save(new AiSkill
                {
                    Id = _loadedId,
                    Name = _name.Text,
                    Body = _body.Text
                });
                _assistant.WriteAction("skill.save", "name=" + saved.Name + "\r\n" + saved.Body, null, null);
                Reload(saved.Id);
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(this, exception);
            }
        }

        private void DeleteSelected()
        {
            SkillItem item = _list.SelectedItem as SkillItem;
            if (item == null) return;
            if (MessageBox.Show(this, "删除 Skill「" + item.Skill.Name + "」？", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;
            string name = item.Skill.Name;
            string id = item.Skill.Id;
            _assistant.Skills.Delete(id);
            _assistant.WriteAction("skill.delete", "name=" + name, null, null);
            _loadedId = null;
            Reload(null);
        }

        /// <summary>重新读出已保存的 Skill，并在末尾加上「新增」。</summary>
        /// <param name="selectId">要选中的 Skill。为空时有 Skill 就选第一条，没有就进入新建。</param>
        private void Reload(string selectId)
        {
            Binding = true;
            _list.Items.Clear();
            bool any = false;
            foreach (AiSkill skill in _assistant.Skills.GetAll())
            {
                _list.Items.Add(new SkillItem(skill));
                any = true;
            }

            _list.Items.Add(_addItem);
            Binding = false;
            if (!string.IsNullOrEmpty(selectId))
            {
                for (int index = 0; index < _list.Items.Count; index++)
                {
                    SkillItem item = _list.Items[index] as SkillItem;
                    if (item == null || !string.Equals(item.Skill.Id, selectId, StringComparison.OrdinalIgnoreCase))
                        continue;
                    _list.SelectedIndex = index;
                    return;
                }
            }

            if (!any)
            {
                BeginNew();
                return;
            }

            _list.SelectedIndex = 0;
        }

        /// <summary>点「新增」进入新建。右键落在「新增」上时保持原来的选中，不弹出删除。</summary>
        private void ListSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_binding) return;
            if (_list.SelectedItem is AddItem)
            {
                // 等这次点击把选中写完，再清空右侧。「新增」保持选中。
                if (IsHandleCreated)
                    BeginInvoke(new Action(BeginNewIfAddRow));
                else
                    BeginNew();
                return;
            }

            SkillItem item = _list.SelectedItem as SkillItem;
            if (item == null) return;
            _loadedId = item.Skill.Id;
            _name.Text = item.Skill.Name;
            _body.Text = item.Skill.Body;
        }

        /// <summary>鼠标停在「新增」上时显示悬浮说明，停在 Skill 上则收起。</summary>
        private void ListMouseMove(object sender, MouseEventArgs e)
        {
            int index = _list.IndexFromPoint(e.Location);
            string tip = index >= 0 && index < _list.Items.Count && _list.Items[index] is AddItem
                ? AppPresentation.AiSkillAddTip
                : string.Empty;
            if (string.Equals(tip, _listTip, StringComparison.Ordinal)) return;
            _listTip = tip;
            _tips.SetToolTip(_list, tip);
        }

        /// <summary>鼠标离开列表后收起「新增」的悬浮说明。</summary>
        private void ListMouseLeave(object sender, EventArgs e)
        {
            _listTip = string.Empty;
            _tips.SetToolTip(_list, string.Empty);
        }

        /// <summary>只在某一条 Skill 上弹出删除。点在「新增」或空白处不弹出。</summary>
        private void ListMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            int index = _list.IndexFromPoint(e.Location);
            if (index < 0 || index >= _list.Items.Count || !(_list.Items[index] is SkillItem)) return;
            if (_list.SelectedIndex != index)
                _list.SelectedIndex = index;
            _listMenu.Show(_list, e.Location);
        }

        /// <summary>画出 Skill 行。「新增」及其分隔线由共用列表处理。</summary>
        private void ListDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _list.Items.Count) return;

            e.DrawBackground();
            SkillItem item = _list.Items[e.Index] as SkillItem;
            if (item == null) return;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            AddItemListBox.DrawEntry(e.Graphics, e.Font, e.Bounds, item.ToString(), _skillIcon, selected, 24, 4);
            e.DrawFocusRectangle();
        }

        private static Label Caption(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4)
            };
        }

        private static Button ButtonOf(string text)
        {
            return new Button { Text = text, AutoSize = true, Margin = new Padding(6, 0, 0, 0), UseVisualStyleBackColor = true };
        }

        private sealed class SkillItem
        {
            public SkillItem(AiSkill skill) { Skill = skill; }
            public AiSkill Skill { get; }
            public override string ToString() { return Skill.Name; }
        }
    }
}

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
        private const int RowHeight = 28;
        private const int SeparatorBand = 8;

        private readonly AiAssistant _assistant;
        private readonly Image _addIcon = DatabaseTreeImageCatalog.LoadActionIcon("icon_add.png");
        private readonly Image _skillIcon = AiModelIcons.Skill(true);
        private readonly ToolTip _tips = new ToolTip();
        private readonly ContextMenuStrip _listMenu = new ContextMenuStrip();
        private readonly AddSkillItem _addItem = new AddSkillItem();
        private readonly SkillListBox _list = new SkillListBox();
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
        private int _anchorIndex = -1;
        private string _listTip = string.Empty;
        private string _loadedId;

        /// <summary>创建 Skill 设置窗体。</summary>
        /// <param name="assistant">Skill 存储。</param>
        /// <param name="settings">记住窗体尺寸。</param>
        public AiSkillSettingsForm(AiAssistant assistant, ISettingsStore settings)
        {
            _assistant = assistant ?? throw new ArgumentNullException(nameof(assistant));
            Text = AppPresentation.WindowTitle("Skill 设置");
            Width = 760;
            Height = 560;
            MinimumSize = new Size(640, 520);
            StartPosition = FormStartPosition.Manual;
            FormSizeMemory.Attach(this, settings, "AiSkills");

            _list.MeasureItem += ListMeasureItem;
            _list.DrawItem += ListDrawItem;
            _list.MouseMove += ListMouseMove;
            _list.MouseLeave += ListMouseLeave;
            _list.MouseUp += ListMouseUp;
            _list.SelectedIndexChanged += ListSelectedIndexChanged;

            ToolStripMenuItem deleteItem = new ToolStripMenuItem("删除", null, (sender, args) => DeleteSelected());
            deleteItem.ToolTipText = AppPresentation.AiSkillDeleteTip;
            _listMenu.Items.Add(deleteItem);

            Label help = new Label
            {
                AutoSize = false,
                Margin = new Padding(0),
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
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

            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0, 8, 0, 0)
            };
            actions.Controls.Add(closeButton);
            actions.Controls.Add(saveButton);
            actions.Controls.Add(importButton);

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

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 3
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(help, 0, 0);
            root.Controls.Add(split, 0, 1);
            root.Controls.Add(actions, 0, 2);
            Controls.Add(root);
            CancelButton = closeButton;
            root.SizeChanged += (sender, args) => FitHelp(root, help);
            Load += (sender, args) =>
            {
                FitHelp(root, help);
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

        /// <summary>按当前窗宽让顶部说明换行，并把所在行撑到能放下末行。</summary>
        /// <param name="root">容纳说明的表格。宽度含内边距。</param>
        /// <param name="help">Skill 用途说明。底部留有空隙。</param>
        private static void FitHelp(TableLayoutPanel root, Label help)
        {
            int width = root.ClientSize.Width - root.Padding.Horizontal;
            if (width < 1) return;
            Size measured = TextRenderer.MeasureText(
                help.Text,
                help.Font,
                new Size(width, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.Left | TextFormatFlags.TextBoxControl);
            int height = measured.Height + help.Padding.Vertical;
            if (help.Width == width && help.Height == height) return;
            help.Size = new Size(width, height);
        }

        /// <summary>选中仍停在「新增」时才清空右侧。中间若已改选 Skill，则不打断。</summary>
        private void BeginNewIfAddRow()
        {
            if (IsDisposed || !(_list.SelectedItem is AddSkillItem)) return;
            BeginNew();
        }

        /// <summary>清空右侧，开始填写一条还没保存的 Skill，并让「新增」保持选中。</summary>
        private void BeginNew()
        {
            _binding = true;
            int addIndex = _list.Items.Count - 1;
            if (addIndex >= 0 && _list.Items[addIndex] is AddSkillItem)
            {
                _list.SelectedIndex = addIndex;
                _anchorIndex = addIndex;
            }

            _loadedId = null;
            _name.Text = string.Empty;
            _body.Text = string.Empty;
            _binding = false;
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
            _binding = true;
            _list.Items.Clear();
            bool any = false;
            foreach (AiSkill skill in _assistant.Skills.GetAll())
            {
                _list.Items.Add(new SkillItem(skill));
                any = true;
            }

            _list.Items.Add(_addItem);
            _binding = false;
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
            if (_list.SelectedItem is AddSkillItem)
            {
                if (_list.RightButton)
                {
                    _binding = true;
                    _list.SelectedIndex = RestorableIndex();
                    _binding = false;
                    return;
                }

                // 等这次点击把选中写完，再清空右侧。「新增」保持选中。
                if (IsHandleCreated)
                    BeginInvoke(new Action(BeginNewIfAddRow));
                else
                    BeginNew();
                return;
            }

            SkillItem item = _list.SelectedItem as SkillItem;
            if (item == null) return;
            _anchorIndex = _list.SelectedIndex;
            _loadedId = item.Skill.Id;
            _name.Text = item.Skill.Name;
            _body.Text = item.Skill.Body;
        }

        /// <summary>有 Skill 时，「新增」上方多出一条分隔线的高度。</summary>
        private void ListMeasureItem(object sender, MeasureItemEventArgs e)
        {
            bool separated = e.Index > 0 && e.Index < _list.Items.Count && _list.Items[e.Index] is AddSkillItem;
            e.ItemHeight = separated ? RowHeight + SeparatorBand : RowHeight;
        }

        /// <summary>鼠标停在「新增」上时显示悬浮说明，停在 Skill 上则收起。</summary>
        private void ListMouseMove(object sender, MouseEventArgs e)
        {
            int index = _list.IndexFromPoint(e.Location);
            string tip = index >= 0 && index < _list.Items.Count && _list.Items[index] is AddSkillItem
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

        /// <summary>右键没有点在 Skill 上时退回的行。正在新建时退回「新增」，否则为 -1。</summary>
        private int RestorableIndex()
        {
            if (_anchorIndex < 0 || _anchorIndex >= _list.Items.Count) return -1;
            object item = _list.Items[_anchorIndex];
            return item is SkillItem || item is AddSkillItem ? _anchorIndex : -1;
        }

        /// <summary>画出 Skill 行，或末项「新增」及其上方的分隔线。</summary>
        private void ListDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _list.Items.Count) return;
            if (_list.Items[e.Index] is AddSkillItem)
            {
                DrawAddItem(e);
                return;
            }

            e.DrawBackground();
            SkillItem item = _list.Items[e.Index] as SkillItem;
            if (item == null) return;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            DrawEntry(e.Graphics, e.Font, e.Bounds, item.ToString(), _skillIcon, selected);
            e.DrawFocusRectangle();
        }

        /// <summary>选中时只高亮文字行。上面已有 Skill 时，分隔线仍用窗口底色。</summary>
        private void DrawAddItem(DrawItemEventArgs e)
        {
            bool selected = (e.State & DrawItemState.Selected) != 0;
            Rectangle content = e.Bounds;
            if (e.Index > 0)
            {
                e.Graphics.FillRectangle(SystemBrushes.Window, new Rectangle(e.Bounds.Left, e.Bounds.Top, e.Bounds.Width, SeparatorBand));
                int y = e.Bounds.Top + (SeparatorBand / 2);
                e.Graphics.DrawLine(SystemPens.ControlDark, e.Bounds.Left, y, e.Bounds.Right - 1, y);
                content = new Rectangle(e.Bounds.Left, e.Bounds.Top + SeparatorBand, e.Bounds.Width, Math.Max(0, e.Bounds.Height - SeparatorBand));
            }

            e.Graphics.FillRectangle(selected ? SystemBrushes.Highlight : SystemBrushes.Window, content);
            DrawEntry(e.Graphics, e.Font, content, _addItem.ToString(), _addIcon, selected);
            if ((e.State & DrawItemState.Focus) != 0)
                ControlPaint.DrawFocusRectangle(e.Graphics, content);
        }

        /// <summary>在一行里画出 16 像素图标和文字。选中时文字用高亮色。</summary>
        /// <param name="graphics">列表项的画布。</param>
        /// <param name="font">列表字体。</param>
        /// <param name="bounds">图标和文字所在的区域。分隔线不在这里面。</param>
        /// <param name="text">Skill 名称或「新增」。</param>
        /// <param name="icon">左侧图标。没有时只画文字。</param>
        /// <param name="selected">为 true 时文字用系统高亮色。</param>
        private static void DrawEntry(Graphics graphics, Font font, Rectangle bounds, string text, Image icon, bool selected)
        {
            if (icon != null)
                graphics.DrawImage(icon, bounds.Left + 4, bounds.Top + (bounds.Height - icon.Height) / 2, icon.Width, icon.Height);
            Color color = selected ? SystemColors.HighlightText : SystemColors.ControlText;
            TextRenderer.DrawText(
                graphics,
                text,
                font,
                new Rectangle(bounds.Left + 24, bounds.Top, Math.Max(0, bounds.Width - 28), bounds.Height),
                color,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.Left);
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

        /// <summary>在选中变化之前记下是不是右键，避免右键点「新增」被当成开始新建。</summary>
        private sealed class SkillListBox : ListBox
        {
            public SkillListBox()
            {
                Dock = DockStyle.Fill;
                IntegralHeight = false;
                BorderStyle = BorderStyle.FixedSingle;
                DrawMode = DrawMode.OwnerDrawVariable;
                ItemHeight = RowHeight;
            }

            /// <summary>当前这次按下是否为右键。在窗口消息里、选中变化之前写好。</summary>
            public bool RightButton { get; private set; }

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
        }

        /// <summary>列表末项。点它开始填写新 Skill，本身不是一条已保存的 Skill。</summary>
        private sealed class AddSkillItem
        {
            /// <summary>列表里显示的文字。</summary>
            public override string ToString()
            {
                return "新增";
            }
        }

        private sealed class SkillItem
        {
            public SkillItem(AiSkill skill) { Skill = skill; }
            public AiSkill Skill { get; }
            public override string ToString() { return Skill.Name; }
        }
    }
}

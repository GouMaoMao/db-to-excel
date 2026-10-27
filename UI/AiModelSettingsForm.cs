using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;

namespace DB2Sheet.UI
{
    /// <summary>编辑模型配置。协议只有 OpenAI 兼容、Claude 和 Gemini 三类。</summary>
    /// <remarks>
    /// 不在这里填写子型号。测试会向接口拉取型号，并用其中一个发一句不含业务内容的话。
    /// 只有这次测试成功，保存才可用。已保存的密钥以固定掩码显示，不明文解出。
    /// 左侧列表按配置分组，父项展开后缩进显示子型号；子型号仅供查看。列表末项是「新增」，删除在配置行的右键菜单里。
    /// </remarks>
    public sealed class AiModelSettingsForm : AppForm
    {
        private const string KeyMask = "••••••••";
        private const int RowHeight = AddItemListBox.DefaultRowHeight;

        private readonly AiAssistant _assistant;
        private readonly IOperationRunner _operationRunner;
        private readonly Image _addIcon = DatabaseTreeImageCatalog.LoadActionIcon("icon_add.png");
        private readonly ToolTip _tips = new ToolTip();
        private readonly ContextMenuStrip _listMenu = new ContextMenuStrip();
        private readonly AddItem _addItem = new AddItem();
        private readonly AddItemListBox _list;
        private readonly List<AiModelProfile> _profiles = new List<AiModelProfile>();
        private readonly HashSet<string> _expandedProfileIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly TextBox _name = new TextBox { BorderStyle = BorderStyle.FixedSingle };
        private readonly ComboBox _protocol = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _endpoint = new TextBox { BorderStyle = BorderStyle.FixedSingle };
        private readonly TextBox _key = new TextBox { BorderStyle = BorderStyle.FixedSingle, UseSystemPasswordChar = true };
        private readonly CheckBox _default = new CheckBox { Text = "作为默认模型", AutoSize = true, UseVisualStyleBackColor = true };
        private readonly Label _status = new Label
        {
            AutoSize = true,
            Dock = DockStyle.Top,
            Margin = new Padding(0),
            ForeColor = SystemColors.GrayText,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(0, 6, 0, 6)
        };
        private readonly Button _saveButton;
        private bool _binding;
        private bool _busy;
        private bool _keyDirty;
        private string _listTip = string.Empty;
        private string _loadedId;
        private string _loadedKey;
        private string _testedFingerprint;
        private List<string> _testedIds = new List<string>();

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

        /// <summary>创建模型设置窗体。</summary>
        /// <param name="assistant">模型存储、拉列表和探测用的助手。</param>
        /// <param name="operationRunner">显示测试连接进度的运行器。</param>
        /// <param name="settings">记住窗体尺寸。</param>
        public AiModelSettingsForm(AiAssistant assistant, IOperationRunner operationRunner, ISettingsStore settings)
        {
            _assistant = assistant ?? throw new ArgumentNullException(nameof(assistant));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            InitializeWindow(
                AppPresentation.WindowTitle("模型设置"),
                780,
                560,
                new Size(680, 520),
                settings,
                "AiModels");

            _list = new AddItemListBox(_addIcon);
            _list.SelectableItem = item => item is ProfileItem;
            _protocol.Items.Add(AiProtocol.OpenAiCompatible);
            _protocol.Items.Add(AiProtocol.Claude);
            _protocol.Items.Add(AiProtocol.Gemini);
            _protocol.Format += (sender, args) =>
            {
                if (args.ListItem is AiProtocol protocol)
                    args.Value = AiProtocolNames.DisplayName(protocol);
            };
            _protocol.SelectedIndex = 0;
            _list.EntryMeasure += ListMeasureItem;
            _list.EntryDraw += ListDrawItem;
            _list.KeyDown += ListKeyDown;
            _list.MouseMove += ListMouseMove;
            _list.MouseLeave += ListMouseLeave;
            _list.MouseUp += ListMouseUp;

            ToolStripMenuItem deleteItem = new ToolStripMenuItem("删除", null, (sender, args) => DeleteSelected());
            deleteItem.ToolTipText = AppPresentation.AiModelDeleteTip;
            _listMenu.Items.Add(deleteItem);

            Label help = new WrapLabel
            {
                Margin = new Padding(0),
                ForeColor = SystemColors.GrayText,
                Text = AppPresentation.AiProtocolHelp,
                Padding = new Padding(0, 0, 0, 8)
            };
            _list.SelectedIndexChanged += ListSelectedIndexChanged;
            _endpoint.TextChanged += (sender, args) => RefreshSaveState();
            _key.TextChanged += KeyTextChanged;
            _key.Enter += (sender, args) => SelectMask();
            _key.MouseUp += (sender, args) => SelectMask();
            _protocol.SelectedIndexChanged += (sender, args) => RefreshSaveState();

            _saveButton = ButtonOf("保存");
            Button testButton = ButtonOf("检测模型");
            _tips.SetToolTip(testButton, AppPresentation.AiModelTestTip);
            Button closeButton = ButtonOf("关闭");
            closeButton.DialogResult = DialogResult.Cancel;
            _saveButton.Click += (sender, args) => SaveSelected();
            testButton.Click += TestButtonClick;
            _saveButton.Enabled = false;

            FlowLayoutPanel actions = CreateActionBar(closeButton, testButton, _saveButton);

            TableLayoutPanel editor = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 9,
                Padding = new Padding(12, 0, 0, 0)
            };
            editor.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 88));
            editor.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            int row = 0;
            row = AddField(editor, row, "显示名", _name, AppPresentation.AiModelNameNote);
            row = AddField(editor, row, "协议", _protocol, AppPresentation.AiModelProtocolNote);
            row = AddField(editor, row, "接口地址", _endpoint, AppPresentation.AiModelEndpointNote);
            row = AddField(editor, row, "密钥", _key, AppPresentation.AiModelKeyNote);
            editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            editor.Controls.Add(_default, 1, row);

            // 右侧字段比窗口矮时无法压下任何一行，包一层自动滚动面板，矮窗里也能逐项滚到。
            editor.AutoSize = true;
            editor.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            editor.Dock = DockStyle.Top;
            Panel editorScroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };
            editorScroll.Controls.Add(editor);

            SplitContainer split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6
            };
            split.Panel1.Controls.Add(_list);
            split.Panel2.Controls.Add(editorScroll);

            TableLayoutPanel root = CreateRootTable(
                new RowStyle(SizeType.AutoSize),
                new RowStyle(SizeType.Percent, 100),
                new RowStyle(SizeType.AutoSize),
                new RowStyle(SizeType.AutoSize));
            root.Controls.Add(help, 0, 0);
            root.Controls.Add(split, 0, 1);
            root.Controls.Add(_status, 0, 2);
            root.Controls.Add(actions, 0, 3);
            Controls.Add(root);
            CancelButton = closeButton;
            Load += (sender, args) =>
            {
                TrySetSplitter(split);
                Reload(null);
            };
            FormClosed += (sender, args) =>
            {
                _addIcon.Dispose();
                _tips.Dispose();
                _listMenu.Dispose();
            };
            _status.Text = AppPresentation.SaveRequiresTestHint;
        }

        /// <summary>选中仍停在「新增」时才清空右侧。中间若已改选配置，则不打断。</summary>
        private void BeginNewIfAddRow()
        {
            if (IsDisposed || !(_list.SelectedItem is AddItem)) return;
            BeginNew();
        }

        /// <summary>清空右侧，开始填写一条还没保存的配置，并让「新增」保持选中。</summary>
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
            _loadedKey = null;
            _keyDirty = false;
            _name.Text = string.Empty;
            _endpoint.Text = string.Empty;
            _key.Text = string.Empty;
            _default.Checked = _assistant.Models.GetAll().Count == 0;
            _protocol.SelectedIndex = 0;
            Binding = false;
            ClearTest();
            _status.Text = "新建配置。密钥需要填写一次。";
            _name.Focus();
        }

        private void DeleteSelected()
        {
            ProfileItem item = _list.SelectedItem as ProfileItem;
            if (item == null) return;
            if (MessageBox.Show(this, "删除模型「" + item.Profile.DisplayName + "」？", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;
            try
            {
                _assistant.Models.Delete(item.Profile.Id);
                _assistant.WriteAction(
                    "model.delete",
                    "display=" + item.Profile.DisplayName + "\r\nprotocol=" + item.Profile.Protocol,
                    null,
                    null);
                Reload(null);
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(this, exception);
            }
        }

        private void SaveSelected()
        {
            try
            {
                if (!TestStillMatches())
                    throw new InvalidOperationException("请先测试连接成功后再保存。");
                bool keyChanged = _keyDirty;
                AiModelProfile profile = ReadProfile(_testedIds);
                _assistant.Models.Save(profile);
                _assistant.WriteAction(
                    "model.save",
                    "display=" + profile.DisplayName
                    + "\r\nprotocol=" + profile.Protocol
                    + "\r\nendpoint=" + AiEndpoints.Describe(profile.Endpoint)
                    + "\r\nmodels=" + string.Join(",", _testedIds)
                    + "\r\nkeyChanged=" + (keyChanged ? "true" : "false"),
                    null,
                    null);
                Reload(profile.Id);
                _status.Text = "已保存。";
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(this, exception);
            }
        }

        private async void TestButtonClick(object sender, EventArgs e)
        {
            if (_busy) return;
            _busy = true;
            _saveButton.Enabled = false;
            try
            {
                AiModelProfile profile = ReadProfile(null);
                string apiKey = ResolveKey();
                string fingerprint = Fingerprint(profile, apiKey);
                AiProtocol protocol = AiProtocolNames.Parse(profile.Protocol);
                IReadOnlyList<string> ids = await _operationRunner.RunAsync(
                    this,
                    "测试模型连接",
                    false,
                    async (progress, cancellationToken) =>
                    {
                        Stopwatch stopwatch = Stopwatch.StartNew();
                        progress.Report(new OperationProgress
                        {
                            Stage = OperationStage.Connecting,
                            Message = "正在读取可用型号…",
                            IsIndeterminate = true
                        });
                        IReadOnlyList<string> availableIds = await _assistant.ListModelsAsync(protocol, profile.Endpoint, apiKey);
                        string probe = AiAssistant.ChooseProbeModel(availableIds);
                        progress.Report(new OperationProgress
                        {
                            Stage = OperationStage.Connecting,
                            Message = string.Format("已读取 {0:N0} 个型号，将使用“{1}”探测响应。", availableIds.Count, probe),
                            IsIndeterminate = true
                        });
                        progress.Report(new OperationProgress
                        {
                            Stage = OperationStage.Executing,
                            Message = "正在探测模型响应…",
                            IsIndeterminate = true
                        });
                        AiAuditEntry entry = new AiAuditEntry
                        {
                            Action = "model.test",
                            ModelDisplayName = AiModelCaption.Audit(profile.DisplayName, probe),
                            Protocol = profile.Protocol,
                            Endpoint = AiEndpoints.Describe(profile.Endpoint),
                            UserText = "Reply with the single word ok.",
                            Redaction = "无",
                            HasTab = false,
                            Detail = "测试不含业务文本。探测型号=" + probe
                        };
                        await _assistant.AskAsync(
                            profile,
                            apiKey,
                            probe,
                            entry,
                            null,
                            null,
                            entry.UserText,
                            AiThinkingLevels.Automatic,
                            null);
                        progress.Report(new OperationProgress
                        {
                            Stage = OperationStage.Completed,
                            Message = string.Format("模型响应正常，测试完成，耗时 {0:N1} 秒。", stopwatch.Elapsed.TotalSeconds),
                            Percent = 100,
                            Elapsed = stopwatch.Elapsed,
                            IsIndeterminate = false
                        });
                        return availableIds;
                    });
                _testedIds = new List<string>(ids);
                _testedFingerprint = fingerprint;
                _saveButton.Enabled = true;
                _status.Text = "测试成功，可以保存。共 " + ids.Count.ToString() + " 个型号。";
                TextPromptDialog.ShowText(
                    this,
                    "检测结果",
                    string.Format("接口返回 {0:N0} 个型号：\r\n\r\n{1}", ids.Count, string.Join("\r\n", ids)));
            }
            catch (Exception exception)
            {
                ClearTest();
                ExceptionDetailForm.Show(this, exception);
            }
            finally
            {
                _busy = false;
                RefreshSaveState();
            }
        }

        private AiModelProfile ReadProfile(IReadOnlyList<string> modelIds)
        {
            AiEndpoints.EnsureUsable(_endpoint.Text);
            if (string.IsNullOrWhiteSpace(_name.Text))
                throw new ArgumentException("请填写显示名。");
            string protectedKey = !_keyDirty && !string.IsNullOrWhiteSpace(_loadedKey)
                ? _loadedKey
                : AiSecretProtector.Protect(ResolveKey());
            if (string.IsNullOrWhiteSpace(protectedKey))
                throw new ArgumentException("请填写密钥。");
            AiProtocol protocol = _protocol.SelectedItem is AiProtocol selected ? selected : AiProtocol.OpenAiCompatible;
            return new AiModelProfile
            {
                Id = _loadedId,
                DisplayName = _name.Text.Trim(),
                Protocol = AiProtocolNames.ToStorage(protocol),
                Endpoint = _endpoint.Text.Trim(),
                ModelName = string.Empty,
                ModelIds = modelIds == null ? new List<string>() : new List<string>(modelIds),
                ProtectedKey = protectedKey,
                IsDefault = _default.Checked
            };
        }

        private string ResolveKey()
        {
            if (!_keyDirty && !string.IsNullOrWhiteSpace(_loadedKey))
                return AiSecretProtector.Unprotect(_loadedKey);
            if (string.IsNullOrWhiteSpace(_key.Text) || string.Equals(_key.Text, KeyMask, StringComparison.Ordinal))
                throw new ArgumentException("请填写密钥。");
            return _key.Text;
        }

        private void SelectMask()
        {
            if (_binding || _keyDirty) return;
            if (!string.Equals(_key.Text, KeyMask, StringComparison.Ordinal)) return;
            _key.SelectAll();
        }

        private void KeyTextChanged(object sender, EventArgs e)
        {
            if (_binding) return;
            _keyDirty = true;
            RefreshSaveState();
        }

        private void RefreshSaveState()
        {
            if (_binding) return;
            bool ready = TestStillMatches();
            _saveButton.Enabled = ready && !_busy;
            if (!ready && !_busy)
                _status.Text = AppPresentation.SaveRequiresTestHint;
        }

        private bool TestStillMatches()
        {
            if (_testedIds == null || _testedIds.Count == 0 || string.IsNullOrEmpty(_testedFingerprint))
                return false;
            try
            {
                AiModelProfile profile = ReadProfile(_testedIds);
                return string.Equals(_testedFingerprint, Fingerprint(profile, ResolveKey()), StringComparison.Ordinal);
            }
            catch (Exception)
            {
                // 名称或密钥还没填完时，保存保持不可用。
                return false;
            }
        }

        private static string Fingerprint(AiModelProfile profile, string apiKey)
        {
            return (profile.Protocol ?? string.Empty) + "\n" + (profile.Endpoint ?? string.Empty) + "\n" + (apiKey ?? string.Empty);
        }

        private void ClearTest()
        {
            _testedIds = new List<string>();
            _testedFingerprint = null;
            _saveButton.Enabled = false;
        }

        /// <summary>重新读出已保存的配置，并在末尾加上「新增」。</summary>
        /// <param name="selectId">要选中的配置。为空时有配置就选第一条，没有就进入新建。</param>
        private void Reload(string selectId)
        {
            _profiles.Clear();
            _profiles.AddRange(_assistant.Models.GetAll());
            _expandedProfileIds.Clear();
            RebuildVisibleItems(selectId, true);
            if (_profiles.Count == 0 || (!string.IsNullOrEmpty(selectId) && _list.SelectedIndex < 0))
            {
                BeginNew();
                return;
            }

        }

        /// <summary>按展开状态重建配置列表，并保留指定父项的选择。</summary>
        /// <param name="selectId">要保持选择的配置标识；为空时选择第一项。</param>
        /// <param name="loadEditor">是否将所选配置重新载入右侧编辑区。</param>
        private void RebuildVisibleItems(string selectId, bool loadEditor)
        {
            Binding = true;
            _list.Items.Clear();
            int selectedIndex = -1;
            foreach (AiModelProfile profile in _profiles)
            {
                int profileIndex = _list.Items.Add(new ProfileItem(profile));
                if (selectedIndex < 0
                    && (string.IsNullOrEmpty(selectId) || string.Equals(profile.Id, selectId, StringComparison.OrdinalIgnoreCase)))
                    selectedIndex = profileIndex;

                if (_expandedProfileIds.Contains(profile.Id))
                {
                    foreach (string modelId in AiModelCatalog.Ids(profile))
                        _list.Items.Add(new ModelIdItem(profile, modelId));
                }
            }

            _list.Items.Add(_addItem);
            if (selectedIndex >= 0)
                _list.SelectedIndex = selectedIndex;
            Binding = false;

            _list.AnchorIndex = selectedIndex;
            if (loadEditor && selectedIndex >= 0)
                ListSelectedIndexChanged(_list, EventArgs.Empty);
        }

        /// <summary>按配置标识查找当前可见的父项行。</summary>
        /// <param name="profileId">模型配置标识。</param>
        /// <returns>父项行号；未找到时为 -1。</returns>
        private int FindProfileIndex(string profileId)
        {
            for (int index = 0; index < _list.Items.Count; index++)
            {
                ProfileItem item = _list.Items[index] as ProfileItem;
                if (item != null && string.Equals(item.Profile.Id, profileId, StringComparison.OrdinalIgnoreCase))
                    return index;
            }

            return -1;
        }

        /// <summary>处理配置、子型号和「新增」行的选择变化。</summary>
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

            if (_list.SelectedItem is ModelIdItem)
            {
                ModelIdItem modelItem = (ModelIdItem)_list.SelectedItem;
                int profileIndex = FindProfileIndex(modelItem.Profile.Id);
                if (profileIndex >= 0)
                {
                    Binding = true;
                    _list.SelectedIndex = profileIndex;
                    _list.AnchorIndex = profileIndex;
                    Binding = false;
                }
                return;
            }

            ProfileItem item = _list.SelectedItem as ProfileItem;
            if (item == null) return;
            AiModelProfile profile = item.Profile;
            Binding = true;
            _loadedId = profile.Id;
            _loadedKey = profile.ProtectedKey;
            _keyDirty = false;
            _name.Text = profile.DisplayName;
            _endpoint.Text = profile.Endpoint;
            _key.Text = string.IsNullOrWhiteSpace(profile.ProtectedKey) ? string.Empty : KeyMask;
            _default.Checked = profile.IsDefault;
            _protocol.SelectedItem = AiProtocolNames.Parse(profile.Protocol);
            Binding = false;
            ClearTest();
            _status.Text = AppPresentation.SaveRequiresTestHint;
        }

        /// <summary>为配置和子型号行设置行高。「新增」及其分隔线由共用列表处理。</summary>
        private void ListMeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _list.Items.Count)
            {
                e.ItemHeight = RowHeight;
                return;
            }

            e.ItemHeight = _list.Items[e.Index] is ModelIdItem ? RowHeight - 2 : RowHeight;
        }

        /// <summary>鼠标停在新增项或可展开配置上时显示对应说明。</summary>
        private void ListMouseMove(object sender, MouseEventArgs e)
        {
            int index = _list.IndexFromPoint(e.Location);
            string tip = string.Empty;
            if (index >= 0 && index < _list.Items.Count)
            {
                if (_list.Items[index] is AddItem)
                    tip = AppPresentation.AiModelAddTip;
                else
                {
                    ProfileItem profileItem = _list.Items[index] as ProfileItem;
                    if (profileItem != null && AiModelCatalog.Ids(profileItem.Profile).Count > 0)
                        tip = AppPresentation.AiModelGroupTip;
                }
            }

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

        /// <summary>键盘上下移动时跳过仅展示的子型号行。</summary>
        /// <param name="sender">模型列表。</param>
        /// <param name="e">按键事件。</param>
        private void ListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Up && e.KeyCode != Keys.Down && e.KeyCode != Keys.Home && e.KeyCode != Keys.End)
                return;

            int index = _list.SelectedIndex;
            if (e.KeyCode == Keys.Home)
            {
                index = 0;
            }
            else if (e.KeyCode == Keys.End)
            {
                index = _list.Items.Count - 1;
            }
            else
            {
                int direction = e.KeyCode == Keys.Down ? 1 : -1;
                do
                {
                    index += direction;
                }
                while (index >= 0 && index < _list.Items.Count && _list.Items[index] is ModelIdItem);
            }

            if (index >= 0 && index < _list.Items.Count)
                _list.SelectedIndex = index;
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        /// <summary>左键展开或收起配置子型号；右键只在配置父项上弹出删除。</summary>
        private void ListMouseUp(object sender, MouseEventArgs e)
        {
            int index = _list.IndexFromPoint(e.Location);
            if (index < 0 || index >= _list.Items.Count || !(_list.Items[index] is ProfileItem)) return;
            ProfileItem profileItem = (ProfileItem)_list.Items[index];
            if (e.Button == MouseButtons.Left)
            {
                if (AiModelCatalog.Ids(profileItem.Profile).Count == 0) return;
                if (!_expandedProfileIds.Add(profileItem.Profile.Id))
                    _expandedProfileIds.Remove(profileItem.Profile.Id);
                RebuildVisibleItems(profileItem.Profile.Id, false);
                return;
            }

            if (e.Button != MouseButtons.Right) return;
            if (_list.SelectedIndex != index)
                _list.SelectedIndex = index;
            _listMenu.Show(_list, e.Location);
        }

        /// <summary>画出配置父项和子型号行。「新增」及其分隔线由共用列表处理。</summary>
        private void ListDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _list.Items.Count) return;

            e.DrawBackground();
            ModelIdItem modelItem = _list.Items[e.Index] as ModelIdItem;
            if (modelItem != null)
            {
                AddItemListBox.DrawEntry(e.Graphics, e.Font, e.Bounds, modelItem.ModelId, null, false, 42, 20);
                return;
            }

            ProfileItem item = _list.Items[e.Index] as ProfileItem;
            if (item == null) return;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            IReadOnlyList<string> modelIds = AiModelCatalog.Ids(item.Profile);
            if (modelIds.Count > 0)
            {
                string arrow = _expandedProfileIds.Contains(item.Profile.Id) ? "﹀" : "›";
                TextRenderer.DrawText(
                    e.Graphics,
                    arrow,
                    e.Font,
                    new Rectangle(e.Bounds.Left + 4, e.Bounds.Top, 14, e.Bounds.Height),
                    selected ? SystemColors.HighlightText : SystemColors.ControlText,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding);
            }

            Image icon = AiModelIcons.For(item.Profile.DisplayName, modelIds);
            AddItemListBox.DrawEntry(e.Graphics, e.Font, e.Bounds, item.ToString(), icon, selected, 42, 20);
            e.DrawFocusRectangle();
        }

        /// <summary>放一行输入，并在它下面放一行灰色说明。</summary>
        /// <param name="table">右侧编辑表格。</param>
        /// <param name="row">输入所在行。说明占用下一行。</param>
        /// <param name="label">左侧字段名。</param>
        /// <param name="input">输入控件。</param>
        /// <param name="note">输入下方的说明。</param>
        /// <returns>说明行之后的下一行。</returns>
        private static int AddField(TableLayoutPanel table, int row, string label, Control input, string note)
        {
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            table.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 8, 0, 0) }, 0, row);
            input.Dock = DockStyle.Top;
            input.Margin = new Padding(0, 4, 0, 0);
            table.Controls.Add(input, 1, row);
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(new WrapLabel
            {
                Text = note,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 0, 0, 6)
            }, 1, row + 1);
            return row + 2;
        }

        private static Button ButtonOf(string text)
        {
            return new Button { Text = text, AutoSize = true, Margin = new Padding(6, 0, 0, 0), UseVisualStyleBackColor = true };
        }

        private static void TrySetSplitter(SplitContainer split)
        {
            int available = split.Width - split.SplitterWidth;
            if (available < 360) return;
            try
            {
                split.SplitterDistance = 220;
            }
            catch (InvalidOperationException)
            {
                // 句柄还没按最终宽度布局时跳过，下次打开会再摆。
            }
        }

        private sealed class ProfileItem
        {
            public ProfileItem(AiModelProfile profile) { Profile = profile; }
            public AiModelProfile Profile { get; }
            public override string ToString()
            {
                return AiModelCaption.Format(Profile);
            }
        }

        /// <summary>列表中只展示、不单独编辑的子型号行。</summary>
        private sealed class ModelIdItem
        {
            /// <summary>创建只展示的子型号行。</summary>
            /// <param name="profile">所属模型配置。</param>
            /// <param name="modelId">接口返回的子型号标识。</param>
            public ModelIdItem(AiModelProfile profile, string modelId)
            {
                Profile = profile;
                ModelId = modelId;
            }

            /// <summary>获取该子型号所属的模型配置。</summary>
            public AiModelProfile Profile { get; }
            /// <summary>获取接口返回的子型号标识。</summary>
            public string ModelId { get; }
        }
    }
}

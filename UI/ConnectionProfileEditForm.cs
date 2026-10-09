using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Providers;

namespace DB2Sheet.UI
{
    /// <summary>提供连接方案的新建、编辑、参数校验和连接测试界面。</summary>
    /// <remarks>
    /// 窗体本身不保存仓储；确认后通过 <see cref="Profile"/> 输出用户编辑得到的模型。
    /// 类型介绍、框下说明和悬停文案来自 <see cref="AppPresentation"/>。新建默认 SQL Server；选 JDBC 时用横幅提示 Java，测试前再硬拦。
    /// </remarks>
    public sealed class ConnectionProfileEditForm : AppForm
    {
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IOperationRunner _operationRunner;
        private readonly IConnectionProfileRepository _connections;
        private readonly IJdbcEnvironmentStore _jdbcEnvironment;
        private readonly TextBox _name;
        private readonly ComboBox _provider;
        private readonly Label _providerHint;
        private readonly Label _nameLabel;
        private readonly Label _providerLabel;
        private readonly Panel _jdbcBanner;
        private readonly Label _jdbcBannerLabel;
        private readonly TableLayoutPanel _parameters;
        private readonly Label _status;
        private readonly Label _saveHint;
        private readonly Button _testButton;
        private readonly Button _saveButton;
        private readonly ToolTip _tips = new ToolTip();
        private readonly Dictionary<string, ParameterEditor> _editors =
            new Dictionary<string, ParameterEditor>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _storedParameters =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Action _openJdbcEnvironment;
        private readonly Button _jdbcEnvironmentButton;
        private readonly List<Label> _fieldNotes = new List<Label>();
        private readonly string _profileId;
        private readonly bool _isNew;
        private bool _loading;
        private bool _connectionTestPassed;
        private bool _nameOwnedByUser;
        private string _lastAutoName = string.Empty;
        private bool _updatingName;

        /// <summary>创建连接方案编辑窗体。</summary>
        /// <param name="providers">提供数据源类型和参数定义的注册表。</param>
        /// <param name="executionService">用于测试连接的服务。</param>
        /// <param name="operationRunner">负责显示连接测试进度。</param>
        /// <param name="connections">用于校验方案名称是否与其它连接重名。</param>
        /// <param name="jdbcEnvironment">本机 Java 路径。选中 JDBC 和测试连接时做快速检查。</param>
        /// <param name="profile">要编辑的快照；为空表示新建方案。</param>
        /// <param name="openJdbcEnvironment">打开 JDBC 环境窗体的操作；非 JDBC 连接不显示入口。</param>
        public ConnectionProfileEditForm(
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IOperationRunner operationRunner,
            IConnectionProfileRepository connections,
            IJdbcEnvironmentStore jdbcEnvironment,
            ConnectionProfileSnapshot profile = null,
            Action openJdbcEnvironment = null)
        {
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _jdbcEnvironment = jdbcEnvironment ?? throw new ArgumentNullException(nameof(jdbcEnvironment));
            _openJdbcEnvironment = openJdbcEnvironment;
            _profileId = profile?.Id ?? Guid.NewGuid().ToString("N");
            _isNew = profile == null;
            _nameOwnedByUser = !_isNew && !string.IsNullOrWhiteSpace(profile.Name);

            Text = AppPresentation.WindowTitle(_isNew ? "新建连接方案" : "编辑连接方案");
            Width = 1000;
            Height = 640;
            MinimumSize = new Size(800, 480);
            StartPosition = FormStartPosition.CenterParent;

            _name = new TextBox { Dock = DockStyle.Fill };
            _provider = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                DisplayMember = nameof(ProviderItem.DisplayName)
            };
            _nameLabel = new Label { Text = "显示名称", AutoSize = true, Anchor = AnchorStyles.Left };
            _providerLabel = new Label { Text = "数据源类型", AutoSize = true, Anchor = AnchorStyles.Left };
            _providerHint = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                ForeColor = SystemColors.GrayText,
                Padding = new Padding(0, 2, 0, 4)
            };
            _jdbcBannerLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                ForeColor = Color.DarkRed,
                Padding = new Padding(8, 6, 8, 6)
            };
            _jdbcBanner = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Visible = false,
                BackColor = Color.FromArgb(255, 243, 205),
                Padding = new Padding(0, 0, 0, 8)
            };
            _jdbcBanner.Controls.Add(_jdbcBannerLabel);
            _parameters = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 3,
                Padding = new Padding(0, 8, 0, 8)
            };
            _parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            _parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _parameters.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _status = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Padding = new Padding(0, 6, 0, 6)
            };
            _testButton = new Button { Text = "测试连接", AutoSize = true };
            _saveButton = new Button { Text = "保存", AutoSize = true, Enabled = false };
            _saveHint = new Label
            {
                Text = AppPresentation.SaveRequiresTestHint,
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(8, 8, 0, 0)
            };
            _jdbcEnvironmentButton = new Button { Text = "配置 JDBC 环境…", AutoSize = true, Visible = false };
            Button cancelButton = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };

            TableLayoutPanel header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Padding = new Padding(0, 0, 0, 4)
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            header.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            header.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            header.Controls.Add(_nameLabel, 0, 0);
            header.Controls.Add(_name, 1, 0);
            header.Controls.Add(_providerLabel, 0, 1);
            header.Controls.Add(_provider, 1, 1);
            header.SetColumnSpan(_providerHint, 2);
            header.Controls.Add(_providerHint, 0, 2);

            Panel parameterHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            parameterHost.Controls.Add(_parameters);
            parameterHost.Controls.Add(_jdbcBanner);

            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            actions.Controls.Add(cancelButton);
            actions.Controls.Add(_saveButton);
            actions.Controls.Add(_saveHint);
            actions.Controls.Add(_testButton);
            actions.Controls.Add(_jdbcEnvironmentButton);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(header, 0, 0);
            root.Controls.Add(parameterHost, 0, 1);
            root.Controls.Add(_status, 0, 2);
            root.Controls.Add(actions, 0, 3);
            Controls.Add(root);

            AppPresentation.SetCueBanner(_name, AppPresentation.ConnectionNameCue);
            AttachTip(_nameLabel, _name, "name");
            AttachTip(_providerLabel, _provider, "provider");
            AttachTip(null, _testButton, "test");
            AttachTip(null, _saveButton, "save");

            _name.TextChanged += NameTextChanged;
            _provider.SelectedIndexChanged += ProviderSelectedIndexChanged;
            _jdbcEnvironmentButton.Click += JdbcEnvironmentButtonClick;
            _testButton.Click += TestButtonClick;
            _saveButton.Click += SaveButtonClick;
            Shown += (sender, args) =>
            {
                ApplyHintWidths();
                RefreshJdbcJavaUi(false);
            };
            Resize += (sender, args) => ApplyHintWidths();
            FormClosed += (sender, args) => _tips.Dispose();
            AcceptButton = _saveButton;
            CancelButton = cancelButton;

            LoadProviders(profile);
            ApplyHintWidths();
        }

        /// <summary>获取用户验证并确认后的连接方案；取消或尚未保存时为空。</summary>
        public ConnectionProfile Profile { get; private set; }

        /// <summary>按固定顺序装入类型列表。新建默认 SQL Server；编辑保留已保存类型。</summary>
        private void LoadProviders(ConnectionProfileSnapshot profile)
        {
            _loading = true;
            try
            {
                Dictionary<string, IDataSourceProvider> byId = _providers.GetAll()
                    .ToDictionary(item => item.ProviderId, item => item, StringComparer.OrdinalIgnoreCase);
                List<ProviderItem> items = new List<ProviderItem>();
                foreach (string id in AppPresentation.ConnectionProviderOrder)
                {
                    IDataSourceProvider provider;
                    if (byId.TryGetValue(id, out provider))
                    {
                        items.Add(new ProviderItem(provider));
                        byId.Remove(id);
                    }
                }

                foreach (IDataSourceProvider leftover in byId.Values.OrderBy(item => item.DisplayName))
                {
                    items.Add(new ProviderItem(leftover));
                }

                _provider.DataSource = items;
                _name.Text = profile?.Name ?? string.Empty;

                string providerId = profile?.ProviderId;
                int selectedIndex = items.FindIndex(item =>
                    string.Equals(item.Provider.ProviderId, providerId, StringComparison.OrdinalIgnoreCase));
                if (selectedIndex < 0)
                {
                    selectedIndex = items.FindIndex(item =>
                        string.Equals(item.Provider.ProviderId, "sqlserver", StringComparison.OrdinalIgnoreCase));
                }

                _provider.SelectedIndex = selectedIndex >= 0 ? selectedIndex : (items.Count > 0 ? 0 : -1);
                BuildParameterEditors(profile?.Parameters);
                UpdateProviderHint();
                UpdateJdbcEnvironmentButton();
                if (_isNew && !_nameOwnedByUser)
                {
                    ApplySuggestedName();
                }
            }
            finally
            {
                _loading = false;
            }
        }

        private void ProviderSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_loading) return;
            _editors.Clear();
            _storedParameters.Clear();
            BuildParameterEditors(null);
            UpdateProviderHint();
            UpdateJdbcEnvironmentButton();
            InvalidateConnectionTest();
            RefreshJdbcJavaUi(false);
            if (!_nameOwnedByUser)
            {
                ApplySuggestedName();
            }
        }

        private void NameTextChanged(object sender, EventArgs e)
        {
            if (_loading || _updatingName) return;
            if (!string.Equals(_name.Text, _lastAutoName, StringComparison.Ordinal))
            {
                _nameOwnedByUser = true;
            }

            InvalidateConnectionTest();
        }

        private void JdbcEnvironmentButtonClick(object sender, EventArgs e)
        {
            if (_openJdbcEnvironment == null) return;
            _openJdbcEnvironment();
            RefreshJdbcJavaUi(false);
        }

        private void BuildParameterEditors(IReadOnlyDictionary<string, string> values)
        {
            bool restoreLoading = _loading;
            _loading = true;
            _parameters.SuspendLayout();
            try
            {
                CaptureEditorValues();
                if (values != null)
                {
                    foreach (KeyValuePair<string, string> pair in values)
                    {
                        _storedParameters[pair.Key] = pair.Value ?? string.Empty;
                    }
                }

                _parameters.Controls.Clear();
                _parameters.RowStyles.Clear();
                _parameters.RowCount = 0;
                _editors.Clear();
                _fieldNotes.Clear();

                IDataSourceProvider provider = SelectedProvider;
                if (provider == null)
                {
                    _status.Text = "没有可用的数据源 Provider。";
                    return;
                }

                if (string.Equals(provider.ProviderId, JdbcConnectionComposer.ProviderId, StringComparison.OrdinalIgnoreCase))
                {
                    BuildJdbcParameterEditors();
                    return;
                }

                foreach (ParameterDefinition definition in provider.ConnectionParameters)
                {
                    if (!IsParameterVisible(definition)) continue;
                    int row = _parameters.RowCount++;
                    _parameters.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                    Label label = new Label
                    {
                        Text = definition.DisplayName + (definition.IsRequired ? " *" : string.Empty),
                        AutoSize = true,
                        Anchor = AnchorStyles.Left,
                        Margin = new Padding(3, 8, 8, 3)
                    };
                    ParameterEditor editor = CreateEditor(definition);
                    editor.SetValue(CurrentParameterValue(definition));
                    if (definition.ValueType == ParameterValueType.Choice && editor.Control is ComboBox choice)
                    {
                        choice.SelectedIndexChanged += ParameterChoiceChanged;
                    }

                    WireEditorChangeHandlers(editor);
                    AttachTip(label, editor.Control, definition.Key);
                    if (editor.AuxiliaryControl != null)
                    {
                        AttachTip(null, editor.AuxiliaryControl, definition.Key);
                    }

                    _editors.Add(definition.Key, editor);
                    _parameters.Controls.Add(label, 0, row);
                    _parameters.Controls.Add(editor.Control, 1, row);
                    if (editor.AuxiliaryControl != null)
                    {
                        _parameters.Controls.Add(editor.AuxiliaryControl, 2, row);
                    }

                    AddFieldNote(definition.Key);
                }
            }
            finally
            {
                _parameters.ResumeLayout(true);
                _loading = restoreLoading;
            }
        }

        /// <summary>为 JDBC 连接生成专用字段：完整 URL、jar 列表、驱动类检测、用户名和密码。</summary>
        private void BuildJdbcParameterEditors()
        {
            TextBox url = new TextBox { Dock = DockStyle.Fill };
            Label urlLabel = AddLabeledRow("JDBC URL *", url, null);
            ParameterEditor urlEditor = new ParameterEditor(url, null, value => url.Text = value, () => url.Text.Trim());
            urlEditor.SetValue(CurrentParameterValue(JdbcParameterKeys.JdbcUrl));
            WireEditorChangeHandlers(urlEditor);
            AttachTip(urlLabel, url, JdbcParameterKeys.JdbcUrl);
            _editors.Add(JdbcParameterKeys.JdbcUrl, urlEditor);
            AddFieldNote(JdbcParameterKeys.JdbcUrl);

            ListBox jars = new ListBox
            {
                Dock = DockStyle.Fill,
                Height = 88,
                IntegralHeight = false,
                HorizontalScrollbar = true
            };
            Button addJar = new Button { Text = "添加…", AutoSize = true };
            Button removeJar = new Button { Text = "移除", AutoSize = true };
            FlowLayoutPanel jarButtons = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };
            jarButtons.Controls.Add(addJar);
            jarButtons.Controls.Add(removeJar);
            addJar.Click += (sender, args) =>
            {
                using (OpenFileDialog dialog = new OpenFileDialog
                {
                    Filter = "JDBC 驱动 (*.jar)|*.jar|所有文件|*.*",
                    Multiselect = true
                })
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    foreach (string file in dialog.FileNames)
                    {
                        if (string.IsNullOrWhiteSpace(file)) continue;
                        string full = Path.GetFullPath(file);
                        bool exists = jars.Items.Cast<object>().Any(item =>
                            string.Equals(Convert.ToString(item), full, StringComparison.OrdinalIgnoreCase));
                        if (!exists) jars.Items.Add(full);
                    }

                    OnParameterEdited();
                }
            };
            removeJar.Click += (sender, args) =>
            {
                if (jars.SelectedIndices.Count == 0) return;
                while (jars.SelectedIndices.Count > 0)
                {
                    jars.Items.RemoveAt(jars.SelectedIndices[0]);
                }

                OnParameterEdited();
            };
            Label jarsLabel = AddLabeledRow("驱动 jar *", jars, jarButtons);
            ParameterEditor jarsEditor = new ParameterEditor(
                jars,
                jarButtons,
                value =>
                {
                    jars.Items.Clear();
                    foreach (string path in JdbcConnectionComposer.SplitDriverJars(value))
                    {
                        jars.Items.Add(path);
                    }
                },
                () => JdbcConnectionComposer.JoinDriverJars(jars.Items.Cast<object>().Select(item => Convert.ToString(item))));
            jarsEditor.SetValue(CurrentParameterValue(JdbcParameterKeys.DriverJars));
            AttachTip(jarsLabel, jars, JdbcParameterKeys.DriverJars);
            AttachTip(null, addJar, JdbcParameterKeys.DriverJars);
            _editors.Add(JdbcParameterKeys.DriverJars, jarsEditor);
            AddFieldNote(JdbcParameterKeys.DriverJars);

            TextBox driverClass = new TextBox { Dock = DockStyle.Fill };
            Button detectDriver = new Button { Text = "检测驱动类", AutoSize = true };
            detectDriver.Click += (sender, args) =>
            {
                List<string> paths = jars.Items.Cast<object>()
                    .Select(item => Convert.ToString(item))
                    .Where(item => !string.IsNullOrWhiteSpace(item))
                    .ToList();
                if (paths.Count == 0)
                {
                    MessageBox.Show(this, "请先添加驱动 jar。", "连接方案", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                IReadOnlyList<string> names = JdbcDriverClassDetector.Detect(paths);
                if (names.Count == 0)
                {
                    MessageBox.Show(
                        this,
                        AppPresentation.DriverClassMissingMessage,
                        "连接方案",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                if (names.Count == 1)
                {
                    driverClass.Text = names[0];
                    _status.ForeColor = Color.DarkGreen;
                    _status.Text = "已填入驱动类 " + names[0];
                    OnParameterEdited();
                    return;
                }

                using (Form chooser = new Form
                {
                    Text = AppPresentation.WindowTitle("选择驱动类"),
                    Width = 480,
                    Height = 280,
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false,
                    MinimizeBox = false,
                    Font = Font,
                    Icon = Icon
                })
                {
                    ListBox list = new ListBox { Dock = DockStyle.Fill };
                    list.Items.AddRange(names.Cast<object>().ToArray());
                    list.SelectedIndex = 0;
                    Button ok = new Button { Text = "确定", DialogResult = DialogResult.OK, AutoSize = true };
                    Button cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
                    FlowLayoutPanel footer = new FlowLayoutPanel
                    {
                        Dock = DockStyle.Bottom,
                        FlowDirection = FlowDirection.RightToLeft,
                        AutoSize = true,
                        Padding = new Padding(8)
                    };
                    footer.Controls.Add(cancel);
                    footer.Controls.Add(ok);
                    chooser.Controls.Add(list);
                    chooser.Controls.Add(footer);
                    chooser.AcceptButton = ok;
                    chooser.CancelButton = cancel;
                    if (chooser.ShowDialog(this) == DialogResult.OK && list.SelectedItem != null)
                    {
                        driverClass.Text = Convert.ToString(list.SelectedItem);
                        OnParameterEdited();
                    }
                }
            };
            Label driverLabel = AddLabeledRow("驱动类 *", driverClass, detectDriver);
            ParameterEditor driverClassEditor = new ParameterEditor(
                driverClass, detectDriver, value => driverClass.Text = value, () => driverClass.Text.Trim());
            driverClassEditor.SetValue(CurrentParameterValue(JdbcParameterKeys.DriverClass));
            WireEditorChangeHandlers(driverClassEditor);
            AttachTip(driverLabel, driverClass, JdbcParameterKeys.DriverClass);
            AttachTip(null, detectDriver, JdbcParameterKeys.DriverClass);
            _editors.Add(JdbcParameterKeys.DriverClass, driverClassEditor);
            AddFieldNote(JdbcParameterKeys.DriverClass);

            TextBox userName = new TextBox { Dock = DockStyle.Fill };
            Label userLabel = AddLabeledRow("用户名", userName, null);
            ParameterEditor userNameEditor = new ParameterEditor(
                userName, null, value => userName.Text = value, () => userName.Text.Trim());
            userNameEditor.SetValue(CurrentParameterValue(DatabaseParameterKeys.UserName));
            WireEditorChangeHandlers(userNameEditor);
            AttachTip(userLabel, userName, DatabaseParameterKeys.UserName);
            _editors.Add(DatabaseParameterKeys.UserName, userNameEditor);
            AddFieldNote(DatabaseParameterKeys.UserName);

            TextBox password = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
            Label passwordLabel = AddLabeledRow("密码", password, null);
            ParameterEditor passwordEditor = new ParameterEditor(
                password, null, value => password.Text = value, () => password.Text);
            passwordEditor.SetValue(CurrentParameterValue(DatabaseParameterKeys.Password));
            WireEditorChangeHandlers(passwordEditor);
            AttachTip(passwordLabel, password, DatabaseParameterKeys.Password);
            _editors.Add(DatabaseParameterKeys.Password, passwordEditor);
        }

        /// <summary>追加一行标签加编辑器，并返回标签以便挂悬停说明。</summary>
        private Label AddLabeledRow(string labelText, Control editor, Control auxiliary)
        {
            int row = _parameters.RowCount++;
            _parameters.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Label label = new Label
            {
                Text = labelText,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(3, 8, 8, 3)
            };
            _parameters.Controls.Add(label, 0, row);
            _parameters.Controls.Add(editor, 1, row);
            if (auxiliary != null)
            {
                _parameters.Controls.Add(auxiliary, 2, row);
            }

            return label;
        }

        private void ParameterChoiceChanged(object sender, EventArgs e)
        {
            if (_loading) return;
            BuildParameterEditors(null);
            OnParameterEdited();
        }

        /// <summary>名称、类型或连接参数变更后须重新测试才能保存。</summary>
        private void InvalidateConnectionTest()
        {
            if (_loading) return;
            _connectionTestPassed = false;
            _saveButton.Enabled = false;
            _saveHint.Visible = true;
        }

        /// <summary>参数变更后刷新测试状态，并在用户未改名称时更新建议名称。</summary>
        private void OnParameterEdited()
        {
            InvalidateConnectionTest();
            if (!_nameOwnedByUser)
            {
                ApplySuggestedName();
            }
        }

        /// <summary>监听编辑控件变更，使保存按钮在参数改动后重新禁用。</summary>
        private void WireEditorChangeHandlers(ParameterEditor editor)
        {
            if (editor == null) return;
            Control control = editor.Control;
            if (control is TextBox textBox)
            {
                textBox.TextChanged += (sender, args) => OnParameterEdited();
            }
            else if (control is ComboBox comboBox)
            {
                comboBox.SelectedIndexChanged += (sender, args) => OnParameterEdited();
            }
            else if (control is CheckBox checkBox)
            {
                checkBox.CheckedChanged += (sender, args) => OnParameterEdited();
            }
        }

        private void CaptureEditorValues()
        {
            foreach (KeyValuePair<string, ParameterEditor> pair in _editors)
            {
                _storedParameters[pair.Key] = pair.Value.GetValue() ?? string.Empty;
            }
        }

        private bool IsParameterVisible(ParameterDefinition definition)
        {
            if (definition == null || string.IsNullOrEmpty(definition.VisibleWhenKey)) return true;
            return string.Equals(
                CurrentParameterValue(definition.VisibleWhenKey),
                definition.VisibleWhenValue,
                StringComparison.OrdinalIgnoreCase);
        }

        private string CurrentParameterValue(ParameterDefinition definition)
        {
            return CurrentParameterValue(definition.Key, definition.DefaultValue);
        }

        private string CurrentParameterValue(string key, object defaultValue = null)
        {
            if (_editors.TryGetValue(key, out ParameterEditor editor))
            {
                return editor.GetValue() ?? string.Empty;
            }

            if (_storedParameters.TryGetValue(key, out string stored))
            {
                return stored ?? string.Empty;
            }

            if (defaultValue == null && SelectedProvider != null)
            {
                ParameterDefinition definition = SelectedProvider.ConnectionParameters.FirstOrDefault(item =>
                    string.Equals(item.Key, key, StringComparison.OrdinalIgnoreCase));
                defaultValue = definition?.DefaultValue;
            }

            return Convert.ToString(defaultValue, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private void UpdateJdbcEnvironmentButton()
        {
            _jdbcEnvironmentButton.Visible = _openJdbcEnvironment != null && IsJdbcSelected();
        }

        private void UpdateProviderHint()
        {
            _providerHint.Text = AppPresentation.ConnectionProviderHint(SelectedProvider?.ProviderId);
            ApplyHintWidths();
        }

        /// <summary>把类型介绍、缺 Java 横幅和框下说明收窄到可用宽度，避免长句横向撑出窗体。</summary>
        private void ApplyHintWidths()
        {
            int width = Math.Max(120, ClientSize.Width - 48);
            _providerHint.MaximumSize = new Size(width, 0);
            _jdbcBannerLabel.MaximumSize = new Size(Math.Max(120, width - 16), 0);
            _jdbcBanner.Height = _jdbcBannerLabel.PreferredHeight + 12;

            int noteWidth = FieldNoteWidth();
            foreach (Label note in _fieldNotes)
            {
                if (note.IsDisposed || note.MaximumSize.Width == noteWidth) continue;
                note.MaximumSize = new Size(noteWidth, 0);
            }
        }

        /// <summary>框下说明放在参数表的输入列，宽度跟该列走；列宽尚未算出时按窗体宽度估算。</summary>
        /// <returns>说明标签的最大宽度。</returns>
        private int FieldNoteWidth()
        {
            int[] columns = _parameters.GetColumnWidths();
            if (columns != null && columns.Length > 1 && columns[1] > 40)
            {
                return Math.Max(120, columns[1] - 8);
            }

            return Math.Max(120, ClientSize.Width - 220);
        }

        /// <summary>在当前字段下方追加灰色说明。该字段没有说明时不占行。</summary>
        /// <param name="fieldKey">字段键，对应 <see cref="AppPresentation.ConnectionFieldNote"/>。</param>
        private void AddFieldNote(string fieldKey)
        {
            string text = AppPresentation.ConnectionFieldNote(SelectedProvider?.ProviderId, fieldKey);
            if (string.IsNullOrEmpty(text)) return;

            Label note = new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(3, 0, 8, 6)
            };
            _fieldNotes.Add(note);
            int row = _parameters.RowCount++;
            _parameters.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _parameters.Controls.Add(note, 1, row);
        }

        /// <summary>当前是否选中 JDBC 数据源。</summary>
        private bool IsJdbcSelected()
        {
            return string.Equals(
                SelectedProvider?.ProviderId,
                JdbcConnectionComposer.ProviderId,
                StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>按已保存路径加快速档查找 java.exe，不扫描 C 盘。</summary>
        /// <returns>查找结果。</returns>
        private JavaResolution ResolveConfiguredJava()
        {
            JdbcEnvironmentSettings settings = _jdbcEnvironment.Load();
            return JavaRuntimeProbe.Resolve(settings == null ? string.Empty : settings.JavaExecutable);
        }

        /// <summary>刷新 JDBC 的 Java 提示。选中时只更新横幅和状态，不弹窗。</summary>
        /// <param name="promptIfMissing">为 true 时未找到会询问是否打开 JDBC 环境，供测试连接使用。</param>
        /// <returns>已找到可用 java.exe 时为 true。非 JDBC 时直接为 true。</returns>
        private bool RefreshJdbcJavaUi(bool promptIfMissing)
        {
            if (!IsJdbcSelected())
            {
                _jdbcBanner.Visible = false;
                return true;
            }

            JavaResolution resolution = ResolveConfiguredJava();
            if (resolution.Found)
            {
                _jdbcBanner.Visible = false;
                _status.ForeColor = SystemColors.GrayText;
                _status.Text = "已检测到 Java。";
                return true;
            }

            _jdbcBannerLabel.Text = AppPresentation.JdbcJavaBanner;
            _jdbcBanner.Visible = true;
            ApplyHintWidths();

            if (promptIfMissing && _openJdbcEnvironment != null)
            {
                DialogResult answer = MessageBox.Show(
                    this,
                    "JDBC 连接需要本机 Java（JDK 8 或更高）。当前没有找到可用的 Java。是否打开「JDBC 环境」进行检测或指定？",
                    "连接方案",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (answer == DialogResult.Yes)
                {
                    _openJdbcEnvironment();
                    resolution = ResolveConfiguredJava();
                    if (resolution.Found)
                    {
                        _jdbcBanner.Visible = false;
                        _status.ForeColor = SystemColors.GrayText;
                        _status.Text = "已检测到 Java。";
                        return true;
                    }
                }
            }
            else if (promptIfMissing)
            {
                MessageBox.Show(
                    this,
                    JavaRuntimeProbe.MissingJavaMessage,
                    "连接方案",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            _status.ForeColor = Color.DarkRed;
            _status.Text = JavaRuntimeProbe.MissingJavaMessage;
            return false;
        }

        /// <summary>在用户未改名称时，按类型和已填内容写入建议显示名称。</summary>
        private void ApplySuggestedName()
        {
            if (_nameOwnedByUser) return;
            string suggested = BuildSuggestedName();
            if (string.IsNullOrEmpty(suggested)) return;
            _updatingName = true;
            try
            {
                _lastAutoName = suggested;
                if (!string.Equals(_name.Text, suggested, StringComparison.Ordinal))
                {
                    _name.Text = suggested;
                }
            }
            finally
            {
                _updatingName = false;
            }
        }

        /// <summary>按当前类型和参数拼建议名称。缺段则省略。</summary>
        private string BuildSuggestedName()
        {
            IDataSourceProvider provider = SelectedProvider;
            if (provider == null) return string.Empty;

            if (IsJdbcSelected())
            {
                string fromUrl = ExtractJdbcNameHint(CurrentParameterValue(JdbcParameterKeys.JdbcUrl));
                return string.IsNullOrEmpty(fromUrl) ? "JDBC" : "JDBC-" + fromUrl;
            }

            string title = AppPresentation.ConnectionProviderTitle(provider.ProviderId);
            string filePath = CurrentParameterValue(DatabaseParameterKeys.FilePath);
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                return string.IsNullOrWhiteSpace(fileName) ? title : title + "-" + fileName;
            }

            List<string> parts = new List<string> { title };
            string host = CurrentParameterValue(DatabaseParameterKeys.Host);
            if (!string.IsNullOrWhiteSpace(host) &&
                !string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase))
            {
                parts.Add(host.Trim());
            }

            string database = CurrentParameterValue(DatabaseParameterKeys.Database);
            if (!string.IsNullOrWhiteSpace(database))
            {
                parts.Add(database.Trim());
            }

            return string.Join("-", parts);
        }

        /// <summary>从 JDBC URL 里尽量抽出主机或库名，供建议显示名称使用。</summary>
        private static string ExtractJdbcNameHint(string jdbcUrl)
        {
            if (string.IsNullOrWhiteSpace(jdbcUrl)) return string.Empty;
            Match host = Regex.Match(jdbcUrl, @"[@/](?://)?([^:/;?\s]+)");
            if (!host.Success) return string.Empty;
            string value = host.Groups[1].Value.Trim();
            if (string.IsNullOrEmpty(value) ||
                string.Equals(value, "localhost", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return value.Length > 40 ? value.Substring(0, 40) : value;
        }

        private void AttachTip(Control label, Control control, string fieldKey)
        {
            string tip = AppPresentation.ConnectionFieldTip(SelectedProvider?.ProviderId, fieldKey);
            if (string.IsNullOrEmpty(tip)) return;
            if (label != null) _tips.SetToolTip(label, tip);
            if (control != null) _tips.SetToolTip(control, tip);
        }

        private ParameterEditor CreateEditor(ParameterDefinition definition)
        {
            if (definition.ValueType == ParameterValueType.Boolean)
            {
                CheckBox checkBox = new CheckBox { AutoSize = true, Anchor = AnchorStyles.Left };
                return new ParameterEditor(
                    checkBox,
                    null,
                    value => checkBox.Checked = string.Equals(value, "true", StringComparison.OrdinalIgnoreCase),
                    () => checkBox.Checked ? "true" : "false");
            }

            if (definition.ValueType == ParameterValueType.Choice)
            {
                ComboBox choice = new ComboBox
                {
                    Dock = DockStyle.Fill,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                choice.Items.AddRange(definition.Choices.Cast<object>().ToArray());
                return new ParameterEditor(
                    choice,
                    null,
                    value => choice.SelectedItem = definition.Choices.FirstOrDefault(item =>
                        string.Equals(item, value, StringComparison.OrdinalIgnoreCase)),
                    () => Convert.ToString(choice.SelectedItem, CultureInfo.InvariantCulture) ?? string.Empty);
            }

            TextBox textBox = new TextBox
            {
                Dock = DockStyle.Fill,
                UseSystemPasswordChar = definition.ValueType == ParameterValueType.Password
            };
            if (string.Equals(definition.Key, DatabaseParameterKeys.Host, StringComparison.OrdinalIgnoreCase))
            {
                AppPresentation.SetCueBanner(textBox, AppPresentation.ConnectionHostCue);
            }

            Control auxiliary = null;
            if (definition.ValueType == ParameterValueType.FilePath)
            {
                Button browse = new Button { Text = "浏览…", AutoSize = true };
                browse.Click += (sender, args) =>
                {
                    using (OpenFileDialog dialog = new OpenFileDialog { FileName = textBox.Text })
                    {
                        if (dialog.ShowDialog() == DialogResult.OK) textBox.Text = dialog.FileName;
                    }
                };
                auxiliary = browse;
            }

            return new ParameterEditor(textBox, auxiliary, value => textBox.Text = value, () => textBox.Text.Trim());
        }

        private void TestButtonClick(object sender, EventArgs e)
        {
            if (!RefreshJdbcJavaUi(true)) return;
            if (!TryCreateProfile(out ConnectionProfile profile, true)) return;

            try
            {
                ConnectionTestResult result = _operationRunner.Run(
                    this,
                    "测试连接",
                    true,
                    (progress, cancellationToken) => _executionService.TestConnectionAsync(
                        profile.CreateSnapshot(),
                        Guid.NewGuid().ToString("N"),
                        progress,
                        cancellationToken));
                _status.ForeColor = result.Succeeded ? Color.DarkGreen : Color.DarkRed;
                _status.Text = result.Message + "（" + result.Elapsed.TotalSeconds.ToString("0.0") + " 秒）";
                if (result.Succeeded)
                {
                    _connectionTestPassed = true;
                    _saveButton.Enabled = true;
                    _saveHint.Visible = false;
                    return;
                }

                _connectionTestPassed = false;
                _saveButton.Enabled = false;
                _saveHint.Visible = true;
                ExceptionDetailForm.Show(this, "连接测试失败", result.Message);
            }
            catch (OperationCanceledException)
            {
                _connectionTestPassed = false;
                _saveButton.Enabled = false;
                _saveHint.Visible = true;
                _status.ForeColor = SystemColors.GrayText;
                _status.Text = "连接测试已取消。";
            }
            catch (Exception exception)
            {
                _connectionTestPassed = false;
                _saveButton.Enabled = false;
                _saveHint.Visible = true;
                _status.ForeColor = Color.DarkRed;
                _status.Text = exception.Message;
                ExceptionDetailForm.Show(this, exception);
            }
        }

        private void SaveButtonClick(object sender, EventArgs e)
        {
            if (!_connectionTestPassed)
            {
                MessageBox.Show(this, "请先测试连接成功后再保存。", "连接方案", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!TryCreateProfile(out ConnectionProfile profile, true)) return;
            Profile = profile;
            DialogResult = DialogResult.OK;
            Close();
        }

        private bool TryCreateProfile(out ConnectionProfile profile, bool showValidation)
        {
            profile = null;
            IDataSourceProvider provider = SelectedProvider;
            if (string.IsNullOrWhiteSpace(_name.Text))
            {
                return ValidationFailure("请输入显示名称。", showValidation, _name);
            }
            if (provider == null)
            {
                return ValidationFailure("请选择数据源类型。", showValidation, _provider);
            }

            string trimmedName = _name.Text.Trim();
            bool nameTaken = _connections.GetAll().Any(item =>
                !string.Equals(item.Id, _profileId, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.Name, trimmedName, StringComparison.CurrentCultureIgnoreCase));
            if (nameTaken)
            {
                return ValidationFailure("已存在同名连接方案。", showValidation, _name);
            }

            CaptureEditorValues();
            Dictionary<string, string> parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (ParameterDefinition definition in provider.ConnectionParameters)
            {
                parameters[definition.Key] = _storedParameters.TryGetValue(definition.Key, out string value)
                    ? value ?? string.Empty
                    : Convert.ToString(definition.DefaultValue, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            profile = new ConnectionProfile
            {
                Id = _profileId,
                Name = trimmedName,
                ProviderId = provider.ProviderId,
                Parameters = parameters
            };

            IReadOnlyList<string> errors = provider.ValidateConnection(profile.CreateSnapshot());
            if (errors.Count > 0)
            {
                profile = null;
                return ValidationFailure(string.Join(Environment.NewLine, errors), showValidation, null);
            }
            return true;
        }

        private bool ValidationFailure(string message, bool showValidation, Control focusTarget)
        {
            if (showValidation)
            {
                MessageBox.Show(this, message, "连接方案", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                focusTarget?.Focus();
            }
            return false;
        }

        private IDataSourceProvider SelectedProvider => (_provider.SelectedItem as ProviderItem)?.Provider;

        /// <summary>为下拉框包装数据源提供程序及其用户可见标题。</summary>
        private sealed class ProviderItem
        {
            public ProviderItem(IDataSourceProvider provider)
            {
                Provider = provider;
            }

            public IDataSourceProvider Provider { get; }

            /// <summary>下拉显示的标题，来自 AppPresentation，不是 Provider.DisplayName。</summary>
            public string DisplayName => AppPresentation.ConnectionProviderTitle(Provider.ProviderId);
        }

        /// <summary>统一封装不同类型参数控件的取值和赋值操作。</summary>
        private sealed class ParameterEditor
        {
            private readonly Action<string> _setValue;
            private readonly Func<string> _getValue;

            public ParameterEditor(Control control, Control auxiliaryControl, Action<string> setValue, Func<string> getValue)
            {
                Control = control;
                AuxiliaryControl = auxiliaryControl;
                _setValue = setValue;
                _getValue = getValue;
            }

            public Control Control { get; }
            public Control AuxiliaryControl { get; }
            public void SetValue(string value) => _setValue(value);
            public string GetValue() => _getValue();
        }
    }
}

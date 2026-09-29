using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Providers;

namespace DB2Sheet.UI
{
    /// <summary>提供连接方案的新建、编辑、参数校验和连接测试界面。</summary>
    /// <remarks>窗体本身不保存仓储；确认后通过 <see cref="Profile"/> 输出用户编辑得到的模型。</remarks>
    public sealed class ConnectionProfileEditForm : AppForm
    {
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IOperationRunner _operationRunner;
        private readonly IConnectionProfileRepository _connections;
        private readonly TextBox _name;
        private readonly ComboBox _provider;
        private readonly TableLayoutPanel _parameters;
        private readonly Label _status;
        private readonly Button _testButton;
        private readonly Button _saveButton;
        private readonly Dictionary<string, ParameterEditor> _editors =
            new Dictionary<string, ParameterEditor>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _storedParameters =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Action _openJdbcEnvironment;
        private readonly Button _jdbcEnvironmentButton;
        private readonly string _profileId;
        private bool _loading;
        private bool _connectionTestPassed;

        /// <summary>创建连接方案编辑窗体。</summary>
        /// <param name="providers">提供数据源类型和参数定义的注册表。</param>
        /// <param name="executionService">用于测试连接的服务。</param>
        /// <param name="operationRunner">负责显示连接测试进度。</param>
        /// <param name="connections">用于校验方案名称是否与其它连接重名。</param>
        /// <param name="profile">要编辑的快照；为空表示新建方案。</param>
        /// <param name="openJdbcEnvironment">打开 JDBC 环境窗体的操作；非 JDBC 连接不显示入口。</param>
        public ConnectionProfileEditForm(
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IOperationRunner operationRunner,
            IConnectionProfileRepository connections,
            ConnectionProfileSnapshot profile = null,
            Action openJdbcEnvironment = null)
        {
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _openJdbcEnvironment = openJdbcEnvironment;
            _profileId = profile?.Id ?? Guid.NewGuid().ToString("N");

            Text = AppPresentation.WindowTitle(profile == null ? "新建连接方案" : "编辑连接方案");
            Width = 1000;
            Height = 620;
            MinimumSize = new Size(800, 460);
            StartPosition = FormStartPosition.CenterParent;

            _name = new TextBox { Dock = DockStyle.Fill };
            _provider = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                DisplayMember = nameof(ProviderItem.DisplayName)
            };
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
            _jdbcEnvironmentButton = new Button { Text = "配置 JDBC 环境…", AutoSize = true, Visible = false };
            Button cancelButton = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };

            TableLayoutPanel header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Padding = new Padding(0, 0, 0, 8)
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.Controls.Add(new Label { Text = "方案名称", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            header.Controls.Add(_name, 1, 0);
            header.Controls.Add(new Label { Text = "数据源类型", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            header.Controls.Add(_provider, 1, 1);

            Panel parameterHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            parameterHost.Controls.Add(_parameters);

            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft
            };
            actions.Controls.Add(cancelButton);
            actions.Controls.Add(_saveButton);
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

            _name.TextChanged += (sender, args) => InvalidateConnectionTest();
            _provider.SelectedIndexChanged += ProviderSelectedIndexChanged;
            _jdbcEnvironmentButton.Click += (sender, args) => _openJdbcEnvironment?.Invoke();
            _testButton.Click += TestButtonClick;
            _saveButton.Click += SaveButtonClick;
            AcceptButton = _saveButton;
            CancelButton = cancelButton;

            LoadProviders(profile);
        }

        /// <summary>获取用户验证并确认后的连接方案；取消或尚未保存时为空。</summary>
        public ConnectionProfile Profile { get; private set; }

        private void LoadProviders(ConnectionProfileSnapshot profile)
        {
            _loading = true;
            try
            {
                List<ProviderItem> items = _providers.GetAll().Select(item => new ProviderItem(item)).ToList();
                _provider.DataSource = items;
                _name.Text = profile?.Name ?? string.Empty;

                string providerId = profile?.ProviderId;
                int selectedIndex = items.FindIndex(item =>
                    string.Equals(item.Provider.ProviderId, providerId, StringComparison.OrdinalIgnoreCase));
                _provider.SelectedIndex = selectedIndex >= 0 ? selectedIndex : (items.Count > 0 ? 0 : -1);
                BuildParameterEditors(profile?.Parameters);
                UpdateJdbcEnvironmentButton();
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
            UpdateJdbcEnvironmentButton();
            _status.Text = string.Empty;
            InvalidateConnectionTest();
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
                    _editors.Add(definition.Key, editor);
                    _parameters.Controls.Add(label, 0, row);
                    _parameters.Controls.Add(editor.Control, 1, row);
                    if (editor.AuxiliaryControl != null)
                    {
                        _parameters.Controls.Add(editor.AuxiliaryControl, 2, row);
                    }
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
            AddLabeledRow("JDBC URL *", url, null);
            ParameterEditor urlEditor = new ParameterEditor(url, null, value => url.Text = value, () => url.Text.Trim());
            urlEditor.SetValue(CurrentParameterValue(JdbcParameterKeys.JdbcUrl));
            WireEditorChangeHandlers(urlEditor);
            _editors.Add(JdbcParameterKeys.JdbcUrl, urlEditor);

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

                    InvalidateConnectionTest();
                }
            };
            removeJar.Click += (sender, args) =>
            {
                if (jars.SelectedIndices.Count == 0) return;
                while (jars.SelectedIndices.Count > 0)
                {
                    jars.Items.RemoveAt(jars.SelectedIndices[0]);
                }

                InvalidateConnectionTest();
            };
            AddLabeledRow("驱动 jar *", jars, jarButtons);
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
            _editors.Add(JdbcParameterKeys.DriverJars, jarsEditor);

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
                        "这些 jar 里没有找到 META-INF/services/java.sql.Driver 声明，请手工填写驱动类。",
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
                    }
                }
            };
            AddLabeledRow("驱动类 *", driverClass, detectDriver);
            ParameterEditor driverClassEditor = new ParameterEditor(
                driverClass, detectDriver, value => driverClass.Text = value, () => driverClass.Text.Trim());
            driverClassEditor.SetValue(CurrentParameterValue(JdbcParameterKeys.DriverClass));
            WireEditorChangeHandlers(driverClassEditor);
            _editors.Add(JdbcParameterKeys.DriverClass, driverClassEditor);

            TextBox userName = new TextBox { Dock = DockStyle.Fill };
            AddLabeledRow("用户名", userName, null);
            ParameterEditor userNameEditor = new ParameterEditor(
                userName, null, value => userName.Text = value, () => userName.Text.Trim());
            userNameEditor.SetValue(CurrentParameterValue(DatabaseParameterKeys.UserName));
            WireEditorChangeHandlers(userNameEditor);
            _editors.Add(DatabaseParameterKeys.UserName, userNameEditor);

            TextBox password = new TextBox { Dock = DockStyle.Fill, UseSystemPasswordChar = true };
            AddLabeledRow("密码", password, null);
            ParameterEditor passwordEditor = new ParameterEditor(
                password, null, value => password.Text = value, () => password.Text);
            passwordEditor.SetValue(CurrentParameterValue(DatabaseParameterKeys.Password));
            WireEditorChangeHandlers(passwordEditor);
            _editors.Add(DatabaseParameterKeys.Password, passwordEditor);
        }

        private void AddLabeledRow(string labelText, Control editor, Control auxiliary)
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
        }

        private void ParameterChoiceChanged(object sender, EventArgs e)
        {
            if (_loading) return;
            BuildParameterEditors(null);
            InvalidateConnectionTest();
        }

        /// <summary>名称、类型或连接参数变更后须重新测试才能保存。</summary>
        private void InvalidateConnectionTest()
        {
            if (_loading) return;
            _connectionTestPassed = false;
            _saveButton.Enabled = false;
        }

        /// <summary>监听编辑控件变更，使保存按钮在参数改动后重新禁用。</summary>
        private void WireEditorChangeHandlers(ParameterEditor editor)
        {
            if (editor == null) return;
            Control control = editor.Control;
            if (control is TextBox textBox)
            {
                textBox.TextChanged += (sender, args) => InvalidateConnectionTest();
            }
            else if (control is ComboBox comboBox)
            {
                comboBox.SelectedIndexChanged += (sender, args) => InvalidateConnectionTest();
            }
            else if (control is CheckBox checkBox)
            {
                checkBox.CheckedChanged += (sender, args) => InvalidateConnectionTest();
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
            _jdbcEnvironmentButton.Visible = _openJdbcEnvironment != null &&
                string.Equals(SelectedProvider?.ProviderId, "jdbc", StringComparison.OrdinalIgnoreCase);
        }

        private static ParameterEditor CreateEditor(ParameterDefinition definition)
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
                    return;
                }

                _connectionTestPassed = false;
                _saveButton.Enabled = false;
                ExceptionDetailForm.Show(this, "连接测试失败", result.Message);
            }
            catch (OperationCanceledException)
            {
                _connectionTestPassed = false;
                _saveButton.Enabled = false;
                _status.ForeColor = SystemColors.GrayText;
                _status.Text = "连接测试已取消。";
            }
            catch (Exception exception)
            {
                _connectionTestPassed = false;
                _saveButton.Enabled = false;
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
                return ValidationFailure("请输入连接方案名称。", showValidation, _name);
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

        /// <summary>为下拉框包装数据源提供程序及其显示名称。</summary>
        private sealed class ProviderItem
        {
            public ProviderItem(IDataSourceProvider provider)
            {
                Provider = provider;
            }

            public IDataSourceProvider Provider { get; }
            public string DisplayName => Provider.DisplayName;
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

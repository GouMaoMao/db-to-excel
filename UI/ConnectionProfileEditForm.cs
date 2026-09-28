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
    /// <summary>提供连接方案的新建、编辑、参数校验和连接测试界面。</summary>
    /// <remarks>窗体本身不保存仓储；确认后通过 <see cref="Profile"/> 输出用户编辑得到的模型。</remarks>
    public sealed class ConnectionProfileEditForm : AppForm
    {
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IOperationRunner _operationRunner;
        private readonly TextBox _name;
        private readonly ComboBox _provider;
        private readonly TableLayoutPanel _parameters;
        private readonly Label _status;
        private readonly Button _testButton;
        private readonly Button _saveButton;
        private readonly Dictionary<string, ParameterEditor> _editors =
            new Dictionary<string, ParameterEditor>(StringComparer.OrdinalIgnoreCase);
        private readonly string _profileId;
        private bool _loading;

        /// <summary>创建连接方案编辑窗体。</summary>
        /// <param name="providers">提供数据源类型和参数定义的注册表。</param>
        /// <param name="executionService">用于测试连接的服务。</param>
        /// <param name="operationRunner">负责显示连接测试进度。</param>
        /// <param name="profile">要编辑的快照；为空表示新建方案。</param>
        public ConnectionProfileEditForm(
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IOperationRunner operationRunner,
            ConnectionProfileSnapshot profile = null)
        {
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _profileId = profile?.Id ?? Guid.NewGuid().ToString("N");

            Text = AppPresentation.WindowTitle(profile == null ? "新建连接方案" : "编辑连接方案");
            Width = 620;
            Height = 620;
            MinimumSize = new Size(520, 460);
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
            _saveButton = new Button { Text = "保存", AutoSize = true };
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

            _provider.SelectedIndexChanged += ProviderSelectedIndexChanged;
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
            }
            finally
            {
                _loading = false;
            }
        }

        private void ProviderSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_loading) return;
            BuildParameterEditors(null);
            _status.Text = string.Empty;
        }

        private void BuildParameterEditors(IReadOnlyDictionary<string, string> values)
        {
            _parameters.SuspendLayout();
            try
            {
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

                foreach (ParameterDefinition definition in provider.ConnectionParameters)
                {
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
                    string value = values != null && values.TryGetValue(definition.Key, out string stored)
                        ? stored
                        : Convert.ToString(definition.DefaultValue, CultureInfo.InvariantCulture);
                    editor.SetValue(value ?? string.Empty);

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
            }
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
            }
            catch (OperationCanceledException)
            {
                _status.ForeColor = SystemColors.GrayText;
                _status.Text = "连接测试已取消。";
            }
            catch (Exception exception)
            {
                _status.ForeColor = Color.DarkRed;
                _status.Text = exception.Message;
            }
        }

        private void SaveButtonClick(object sender, EventArgs e)
        {
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

            Dictionary<string, string> parameters = _editors.ToDictionary(
                pair => pair.Key,
                pair => pair.Value.GetValue(),
                StringComparer.OrdinalIgnoreCase);
            profile = new ConnectionProfile
            {
                Id = _profileId,
                Name = _name.Text.Trim(),
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

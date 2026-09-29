using System;
using System.Drawing;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.UI
{
    /// <summary>列出连接方案，并提供新建、编辑和删除操作。</summary>
    /// <remarks>编辑结果会保存到仓储；关闭后调用方可通过 <see cref="SelectedProfileId"/> 读取最终选择。</remarks>
    public sealed class ConnectionProfilesForm : AppForm
    {
        private readonly IConnectionProfileRepository _repository;
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IOperationRunner _operationRunner;
        private readonly ListView _profiles;
        private readonly Button _editButton;
        private readonly Button _deleteButton;

        /// <summary>创建连接方案管理窗体。</summary>
        /// <param name="repository">连接方案仓储。</param>
        /// <param name="providers">数据源提供程序注册表。</param>
        /// <param name="executionService">连接测试服务。</param>
        /// <param name="operationRunner">连接测试进度运行器。</param>
        /// <param name="settings">用于记住窗体尺寸的设置存储。</param>
        /// <param name="selectedProfileId">打开时优先选中的方案标识。</param>
        public ConnectionProfilesForm(
            IConnectionProfileRepository repository,
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IOperationRunner operationRunner,
            ISettingsStore settings,
            string selectedProfileId = null)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            Text = AppPresentation.WindowTitle("连接方案管理");
            Width = 720;
            Height = 500;
            MinimumSize = new Size(560, 360);
            StartPosition = FormStartPosition.Manual;
            FormSizeMemory.Attach(this, settings, "Connections");

            _profiles = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = false
            };
            _profiles.Columns.Add("名称", 280);
            _profiles.Columns.Add("数据源类型", 180);
            _profiles.Columns.Add("标识", 200);

            Button addButton = new Button { Text = "新建", AutoSize = true };
            _editButton = new Button { Text = "编辑", AutoSize = true, Enabled = false };
            _deleteButton = new Button { Text = "删除", AutoSize = true, Enabled = false };
            Button closeButton = new Button { Text = "关闭", AutoSize = true, DialogResult = DialogResult.OK };

            FlowLayoutPanel toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight
            };
            toolbar.Controls.Add(addButton);
            toolbar.Controls.Add(_editButton);
            toolbar.Controls.Add(_deleteButton);

            FlowLayoutPanel footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft
            };
            footer.Controls.Add(closeButton);

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
            root.Controls.Add(toolbar, 0, 0);
            root.Controls.Add(_profiles, 0, 1);
            root.Controls.Add(footer, 0, 2);
            Controls.Add(root);

            addButton.Click += (sender, args) => EditProfile(null);
            _editButton.Click += (sender, args) => EditSelected();
            _deleteButton.Click += DeleteSelected;
            _profiles.DoubleClick += (sender, args) => EditSelected();
            _profiles.SelectedIndexChanged += (sender, args) => UpdateActions();
            FormClosed += (sender, args) => _repository.Changed -= RepositoryChanged;
            _repository.Changed += RepositoryChanged;
            AcceptButton = closeButton;
            CancelButton = closeButton;

            ReloadProfiles(selectedProfileId);
        }

        /// <summary>获取列表中当前选中方案的标识；没有选择时为空。</summary>
        public string SelectedProfileId => SelectedProfile?.Id;

        private ConnectionProfileSnapshot SelectedProfile =>
            _profiles.SelectedItems.Count == 1
                ? _profiles.SelectedItems[0].Tag as ConnectionProfileSnapshot
                : null;

        private void ReloadProfiles(string selectedProfileId = null)
        {
            string targetId = selectedProfileId ?? SelectedProfileId;
            _profiles.BeginUpdate();
            try
            {
                _profiles.Items.Clear();
                foreach (ConnectionProfileSnapshot profile in _repository.GetAll())
                {
                    IDataSourceProvider provider = _providers.GetById(profile.ProviderId);
                    ListViewItem item = new ListViewItem(profile.Name) { Tag = profile };
                    item.SubItems.Add(provider?.DisplayName ?? profile.ProviderId);
                    item.SubItems.Add(profile.Id);
                    _profiles.Items.Add(item);
                    if (string.Equals(profile.Id, targetId, StringComparison.OrdinalIgnoreCase))
                    {
                        item.Selected = true;
                        item.Focused = true;
                    }
                }
            }
            finally
            {
                _profiles.EndUpdate();
            }
            UpdateActions();
        }

        private void EditSelected()
        {
            if (SelectedProfile != null) EditProfile(SelectedProfile);
        }

        private void EditProfile(ConnectionProfileSnapshot profile)
        {
            using (ConnectionProfileEditForm form = new ConnectionProfileEditForm(
                _providers,
                _executionService,
                _operationRunner,
                profile))
            {
                if (form.ShowDialog(this) != DialogResult.OK) return;
                _repository.Save(form.Profile);
                ReloadProfiles(form.Profile.Id);
            }
        }

        private void DeleteSelected(object sender, EventArgs e)
        {
            ConnectionProfileSnapshot profile = SelectedProfile;
            if (profile == null) return;

            DialogResult confirmation = MessageBox.Show(
                this,
                "确定删除连接方案“" + profile.Name + "”吗？",
                "删除连接方案",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (confirmation != DialogResult.Yes) return;

            _repository.Delete(profile.Id);
            ReloadProfiles();
        }

        private void RepositoryChanged(object sender, ConnectionProfilesChangedEventArgs e)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(ReloadProfiles), e.ProfileId);
                return;
            }
            ReloadProfiles(e.ProfileId);
        }

        private void UpdateActions()
        {
            bool hasSelection = SelectedProfile != null;
            _editButton.Enabled = hasSelection;
            _deleteButton.Enabled = hasSelection;
        }
    }
}

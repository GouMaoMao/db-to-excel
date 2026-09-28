using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.UI
{
    /// <summary>可复用的连接方案下拉选择控件，并提供进入连接管理的按钮。</summary>
    /// <remarks>控件订阅仓储变更事件并自动刷新；释放时会取消订阅，避免仓储长期持有控件。</remarks>
    public sealed class ConnectionProfileSelector : UserControl
    {
        private readonly IConnectionProfileRepository _repository;
        private readonly ComboBox _profiles;
        private readonly Button _manageButton;
        private bool _disposed;

        /// <summary>创建连接方案选择器。</summary>
        /// <param name="repository">连接方案数据来源。</param>
        public ConnectionProfileSelector(IConnectionProfileRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));

            AutoSize = true;
            MinimumSize = new Size(260, 32);

            _profiles = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                DisplayMember = nameof(ProfileItem.DisplayName)
            };
            _manageButton = new Button
            {
                Text = "管理…",
                AutoSize = true,
                Dock = DockStyle.Fill
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.Controls.Add(_profiles, 0, 0);
            layout.Controls.Add(_manageButton, 1, 0);
            Controls.Add(layout);

            _profiles.SelectedIndexChanged += ProfilesSelectedIndexChanged;
            _manageButton.Click += (sender, args) => ManageRequested?.Invoke(this, EventArgs.Empty);
            _repository.Changed += RepositoryChanged;
            ReloadProfiles();
        }

        /// <summary>用户选择不同连接方案时触发。</summary>
        public event EventHandler SelectedProfileChanged;
        /// <summary>用户点击“管理”按钮时触发。</summary>
        public event EventHandler ManageRequested;

        /// <summary>获取或设置当前选中方案的标识；找不到时清除选择。</summary>
        public string SelectedProfileId
        {
            get => (_profiles.SelectedItem as ProfileItem)?.Profile.Id;
            set => SelectProfile(value);
        }

        /// <summary>获取当前选中的不可变连接快照；没有选择时为空。</summary>
        public ConnectionProfileSnapshot SelectedProfile =>
            (_profiles.SelectedItem as ProfileItem)?.Profile;

        /// <summary>从仓储重新加载选项，并尽量保留原选择。</summary>
        public void ReloadProfiles()
        {
            string selectedId = SelectedProfileId;
            IReadOnlyList<ConnectionProfileSnapshot> profiles = _repository.GetAll();
            List<ProfileItem> items = profiles.Select(profile => new ProfileItem(profile)).ToList();

            _profiles.BeginUpdate();
            try
            {
                _profiles.DataSource = null;
                _profiles.DataSource = items;
                _profiles.DisplayMember = nameof(ProfileItem.DisplayName);
                SelectProfile(selectedId);
                if (_profiles.SelectedIndex < 0 && items.Count > 0)
                {
                    _profiles.SelectedIndex = 0;
                }
            }
            finally
            {
                _profiles.EndUpdate();
            }
        }

        /// <summary>取消仓储事件订阅并释放控件资源。</summary>
        /// <param name="disposing">是否释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && !_disposed)
            {
                _repository.Changed -= RepositoryChanged;
                _disposed = true;
            }
            base.Dispose(disposing);
        }

        private void SelectProfile(string profileId)
        {
            if (string.IsNullOrWhiteSpace(profileId))
            {
                _profiles.SelectedIndex = -1;
                return;
            }

            for (int index = 0; index < _profiles.Items.Count; index++)
            {
                ProfileItem item = _profiles.Items[index] as ProfileItem;
                if (string.Equals(item?.Profile.Id, profileId, StringComparison.OrdinalIgnoreCase))
                {
                    _profiles.SelectedIndex = index;
                    return;
                }
            }

            _profiles.SelectedIndex = -1;
        }

        private void RepositoryChanged(object sender, ConnectionProfilesChangedEventArgs e)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(ReloadProfiles));
                return;
            }
            ReloadProfiles();
        }

        private void ProfilesSelectedIndexChanged(object sender, EventArgs e)
        {
            SelectedProfileChanged?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>为下拉框组合连接快照和用户可见名称。</summary>
        private sealed class ProfileItem
        {
            public ProfileItem(ConnectionProfileSnapshot profile)
            {
                Profile = profile;
                DisplayName = string.IsNullOrWhiteSpace(profile.ProviderId)
                    ? profile.Name
                    : profile.Name + "  (" + profile.ProviderId + ")";
            }

            public ConnectionProfileSnapshot Profile { get; }
            public string DisplayName { get; }
        }
    }
}

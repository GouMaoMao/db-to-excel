using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;

namespace DB2Sheet.UI
{
    /// <summary>数据库连接树。查询窗体和批量刷新窗体共用当前连接、图标和右键菜单。</summary>
    /// <remarks>
    /// 当前连接写入 <c>session.activeConnectionId</c>，不重新测试连接也能显示。
    /// 另一窗体改了当前连接后，本控件通过设置变更刷新高亮。
    /// 双击表或视图时只发出 <see cref="ObjectActivated"/>，由查询窗体决定是否插入名称。
    /// 控件订阅仓储和设置事件，释放时取消订阅，并释放图标和字体。
    /// </remarks>
    public sealed class DatabaseConnectionTree : UserControl
    {
        private readonly IConnectionProfileRepository _connections;
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IOperationRunner _operationRunner;
        private readonly ISettingsStore _settings;
        private readonly Action _openJdbcEnvironment;
        private readonly TreeView _tree;
        private readonly ImageList _images;
        private readonly ContextMenuStrip _nodeMenu;
        private readonly ContextMenuStrip _blankMenu;
        private readonly ContextMenuStrip _copyMenu;
        private Font _commentFont;
        private Font _activeNameFont;
        private ConnectionProfileSnapshot _activeConnection;
        private string _activeDatabase;
        private CancellationTokenSource _metadataCancellation;
        private bool _savingActiveId;
        private bool _disposed;

        /// <summary>创建连接树并恢复已记住的当前连接。</summary>
        /// <param name="connections">连接方案仓储。</param>
        /// <param name="providers">数据源提供程序注册表。</param>
        /// <param name="executionService">连接测试服务。</param>
        /// <param name="operationRunner">测试连接时使用的进度窗体运行器。</param>
        /// <param name="settings">当前连接标识的会话设置。</param>
        /// <param name="openJdbcEnvironment">打开 JDBC 环境窗体；为空时连接编辑不显示该入口。</param>
        public DatabaseConnectionTree(
            IConnectionProfileRepository connections,
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IOperationRunner operationRunner,
            ISettingsStore settings,
            Action openJdbcEnvironment)
        {
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _openJdbcEnvironment = openJdbcEnvironment;

            Dock = DockStyle.Fill;
            _images = DatabaseTreeImageCatalog.CreateImageList();
            _tree = new TreeView
            {
                Dock = DockStyle.Fill,
                DrawMode = TreeViewDrawMode.OwnerDrawText,
                HideSelection = false,
                ImageList = _images,
                ShowNodeToolTips = true
            };
            Controls.Add(_tree);

            _nodeMenu = BuildNodeMenu();
            _blankMenu = BuildBlankMenu();
            _copyMenu = BuildCopyMenu();
            _tree.DrawNode += TreeDrawNode;
            _tree.NodeMouseClick += TreeNodeMouseClick;
            _tree.MouseUp += TreeMouseUp;
            _tree.NodeMouseDoubleClick += TreeNodeMouseDoubleClick;
            _tree.BeforeExpand += TreeBeforeExpand;
            _tree.AfterSelect += TreeAfterSelect;
            _connections.Changed += ConnectionsChanged;
            _settings.Changed += SettingsChanged;
            Disposed += ControlDisposed;

            RestoreActiveConnection();
            Reload(null);
        }

        /// <summary>状态说明变化时触发。查询窗体用来更新底部状态栏。</summary>
        public event EventHandler NoticeChanged;

        /// <summary>选中的连接根节点变化时触发，供宿主启用菜单。</summary>
        public event EventHandler SelectionChanged;

        /// <summary>当前连接变化时触发。批量刷新用它决定执行连接。</summary>
        public event EventHandler ActiveConnectionChanged;

        /// <summary>双击表或视图时触发。参数是对象名称，不含注释。</summary>
        public event EventHandler<DatabaseObjectActivatedEventArgs> ObjectActivated;

        /// <summary>获取最近一条状态说明。</summary>
        public string Notice { get; private set; }

        /// <summary>获取当前连接方案。没有当前连接时为空。</summary>
        public ConnectionProfileSnapshot ActiveConnection => _activeConnection;

        /// <summary>获取是否选中了连接根节点。</summary>
        public bool HasSelectedConnection => SelectedConnectionNode != null;

        /// <summary>获取执行查询应使用的连接。选中了当前连接下的库时，切换到该库。</summary>
        /// <returns>可执行的连接快照。没有当前连接时为空。</returns>
        public ConnectionProfileSnapshot ExecutionConnection()
        {
            if (_activeConnection == null) return null;
            if (string.IsNullOrWhiteSpace(_activeDatabase)) return _activeConnection;
            IDatabaseMetadataProvider provider = _providers.GetById(_activeConnection.ProviderId) as IDatabaseMetadataProvider;
            if (provider == null) return _activeConnection;
            try
            {
                return provider.CreateDatabaseSnapshot(_activeConnection, _activeDatabase);
            }
            catch (Exception)
            {
                // 切库快照失败时仍用当前连接执行，不把目录问题当成没有连接。
                return _activeConnection;
            }
        }

        /// <summary>打开新建连接对话框并保存。</summary>
        public void CreateConnection()
        {
            using (ConnectionProfileEditForm editor = new ConnectionProfileEditForm(
                _providers, _executionService, _operationRunner, _connections, null, _openJdbcEnvironment))
            {
                if (editor.ShowDialog(FindForm()) != DialogResult.OK || editor.Profile == null) return;
                _connections.Save(editor.Profile);
                Reload(editor.Profile.Id);
                SetNotice("连接方案已创建。");
            }
        }

        /// <summary>编辑当前选中的连接根节点。</summary>
        public void EditSelectedConnection()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null) return;
            bool wasActive = string.Equals(_activeConnection?.Id, selected.Id, StringComparison.OrdinalIgnoreCase);
            using (ConnectionProfileEditForm editor = new ConnectionProfileEditForm(
                _providers, _executionService, _operationRunner, _connections, selected, _openJdbcEnvironment))
            {
                if (editor.ShowDialog(FindForm()) != DialogResult.OK || editor.Profile == null) return;
                _connections.Save(editor.Profile);
                Reload(editor.Profile.Id);
                ConnectionProfileSnapshot updated = _connections.GetById(editor.Profile.Id);
                if (wasActive && updated != null) ActivateConnection(updated, true);
                else SetNotice("连接方案已更新。");
            }
        }

        /// <summary>删除当前选中的连接方案。</summary>
        public void DeleteSelectedConnection()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null) return;
            DialogResult result = MessageBox.Show(
                FindForm(),
                "确定删除连接方案“" + selected.Name + "”吗？",
                "删除连接方案",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes) return;
            _connections.Delete(selected.Id);
            if (string.Equals(_activeConnection?.Id, selected.Id, StringComparison.OrdinalIgnoreCase))
            {
                _activeConnection = null;
                _activeDatabase = null;
                RememberActiveConnection(null);
                ActiveConnectionChanged?.Invoke(this, EventArgs.Empty);
            }
            Reload(null);
            SetNotice("连接方案已删除。");
        }

        /// <summary>把选中的连接设为当前连接，并测试能否打开。</summary>
        public void SetSelectedAsCurrent()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null) return;
            ActivateConnection(selected, true);
        }

        /// <summary>刷新选中连接的库列表。没有选中连接时重载连接列表。</summary>
        public void RefreshSelected()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null)
            {
                Reload(null);
                return;
            }
            ActivateConnection(selected, false);
        }

        private ConnectionProfileSnapshot SelectedConnectionNode
        {
            get
            {
                TreeNode node = _tree.SelectedNode;
                return node != null && node.Parent == null
                    ? node.Tag as ConnectionProfileSnapshot
                    : null;
            }
        }

        private void SetNotice(string message)
        {
            Notice = message ?? string.Empty;
            NoticeChanged?.Invoke(this, EventArgs.Empty);
        }

        private void RestoreActiveConnection()
        {
            string id = _settings.Get(CoreSettings.ActiveConnectionId);
            if (string.IsNullOrWhiteSpace(id))
            {
                _activeConnection = null;
                _activeDatabase = null;
                return;
            }

            ConnectionProfileSnapshot connection = _connections.GetById(id);
            if (connection == null)
            {
                _activeConnection = null;
                _activeDatabase = null;
                RememberActiveConnection(null);
                return;
            }

            _activeConnection = connection;
        }

        private void RememberActiveConnection(ConnectionProfileSnapshot connection)
        {
            _savingActiveId = true;
            try
            {
                _settings.Set(CoreSettings.ActiveConnectionId, connection?.Id ?? string.Empty);
            }
            finally
            {
                _savingActiveId = false;
            }
        }

        private void SettingsChanged(object sender, EventArgs e)
        {
            if (_savingActiveId || _disposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(ApplyExternalActiveConnection));
                return;
            }
            ApplyExternalActiveConnection();
        }

        /// <summary>另一窗体改了当前连接时，只更新高亮，不重新测试。</summary>
        private void ApplyExternalActiveConnection()
        {
            if (_disposed) return;
            string id = _settings.Get(CoreSettings.ActiveConnectionId);
            if (string.Equals(id, _activeConnection?.Id, StringComparison.OrdinalIgnoreCase)) return;
            _activeDatabase = null;
            if (string.IsNullOrWhiteSpace(id))
            {
                _activeConnection = null;
            }
            else
            {
                _activeConnection = _connections.GetById(id);
            }
            ApplyTreeVisualState();
            ActiveConnectionChanged?.Invoke(this, EventArgs.Empty);
        }

        private void Reload(string selectedConnectionId)
        {
            string selectedId = selectedConnectionId ?? SelectedConnectionNode?.Id ?? _activeConnection?.Id;
            List<ConnectionProfileSnapshot> connections = _connections.GetAll()
                .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            _tree.BeginUpdate();
            try
            {
                _tree.Nodes.Clear();
                foreach (ConnectionProfileSnapshot connection in connections)
                {
                    bool active = string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase);
                    TreeNode node = new TreeNode(connection.Name)
                    {
                        Tag = connection,
                        ToolTipText = string.IsNullOrWhiteSpace(connection.ProviderId)
                            ? connection.Name
                            : connection.Name + "（" + connection.ProviderId + "）",
                        ImageKey = DatabaseTreeImageCatalog.Provider(connection.ProviderId, active),
                        SelectedImageKey = DatabaseTreeImageCatalog.Provider(connection.ProviderId, active)
                    };
                    EnsureDatabasePlaceholder(node);
                    _tree.Nodes.Add(node);
                    if (string.Equals(connection.Id, selectedId, StringComparison.OrdinalIgnoreCase))
                    {
                        _tree.SelectedNode = node;
                    }
                }
            }
            finally
            {
                _tree.EndUpdate();
            }
            ApplyTreeVisualState();
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }

        private async void ActivateConnection(ConnectionProfileSnapshot connection, bool expandNode)
        {
            if (connection == null) return;
            ResetActiveConnectionVisualState();
            try
            {
                ConnectionTestResult result = await _operationRunner.RunAsync(
                    FindForm(),
                    "测试连接",
                    true,
                    (progress, cancellationToken) => _executionService.TestConnectionAsync(
                        connection,
                        Guid.NewGuid().ToString("N"),
                        progress,
                        cancellationToken));

                if (!result.Succeeded)
                {
                    SetNotice(result.Message + "（" + result.Elapsed.TotalSeconds.ToString("0.0") + " 秒）");
                    ExceptionDetailForm.Show(FindForm(), "连接测试失败", result.Message);
                    return;
                }

                _activeConnection = connection;
                _activeDatabase = null;
                RememberActiveConnection(connection);
                Reload(connection.Id);
                TreeNode node = _tree.Nodes.Cast<TreeNode>().FirstOrDefault(item =>
                    string.Equals((item.Tag as ConnectionProfileSnapshot)?.Id, connection.Id, StringComparison.OrdinalIgnoreCase));
                if (node != null)
                {
                    _tree.SelectedNode = node;
                    EnsureDatabasePlaceholder(node);
                    if (expandNode) node.Expand();
                }

                SetNotice("当前连接已激活：" + connection.Name);
                ApplyTreeVisualState();
                ActiveConnectionChanged?.Invoke(this, EventArgs.Empty);
            }
            catch (OperationCanceledException)
            {
                SetNotice("连接测试已取消。");
            }
            catch (Exception exception)
            {
                SetNotice(exception.Message);
                ExceptionDetailForm.Show(FindForm(), exception);
            }
        }

        private void ResetActiveConnectionVisualState()
        {
            _activeConnection = null;
            _activeDatabase = null;
            foreach (TreeNode node in _tree.Nodes)
            {
                ConnectionProfileSnapshot connection = node.Tag as ConnectionProfileSnapshot;
                if (connection == null) continue;
                node.Collapse();
                node.ImageKey = DatabaseTreeImageCatalog.Provider(connection.ProviderId, false);
                node.SelectedImageKey = node.ImageKey;
            }
            ApplyTreeVisualState();
        }

        private ContextMenuStrip BuildNodeMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("设置为当前连接", null, (sender, args) => SetSelectedAsCurrent());
            menu.Items.Add("编辑连接", null, (sender, args) => EditSelectedConnection());
            menu.Items.Add("新建连接", null, (sender, args) => CreateConnection());
            menu.Items.Add("删除连接", null, (sender, args) => DeleteSelectedConnection());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("复制名称", null, (sender, args) => CopySelectedNodeName());
            menu.Items.Add("刷新", null, (sender, args) => RefreshSelected());
            return menu;
        }

        private ContextMenuStrip BuildBlankMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("新建连接", null, (sender, args) => CreateConnection());
            menu.Items.Add("刷新连接列表", null, (sender, args) => Reload(null));
            return menu;
        }

        private ContextMenuStrip BuildCopyMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("复制名称", null, (sender, args) => CopySelectedNodeName());
            return menu;
        }

        private void CopySelectedNodeName()
        {
            TreeNode node = _tree.SelectedNode;
            if (node == null) return;
            string name = ResolveNodeCopyName(node);
            if (string.IsNullOrEmpty(name)) return;
            try
            {
                Clipboard.SetText(name);
            }
            catch (Exception)
            {
                // 剪贴板偶发占用；复制失败不打断操作。
            }
        }

        private static string FormatObjectCaption(DatabaseObjectMetadata item)
        {
            if (item == null) return string.Empty;
            if (string.IsNullOrWhiteSpace(item.Comment)) return item.Name ?? string.Empty;
            return (item.Name ?? string.Empty) + "  " + item.Comment.Trim();
        }

        private static string ResolveNodeCopyName(TreeNode node)
        {
            if (node == null) return string.Empty;
            if (node.Tag is ConnectionProfileSnapshot connection) return connection.Name ?? string.Empty;
            if (node.Tag is DatabaseNodeTag databaseTag) return databaseTag.DatabaseName ?? string.Empty;
            if (node.Tag is DatabaseObjectMetadata objectMetadata) return objectMetadata.Name ?? string.Empty;
            return node.Text ?? string.Empty;
        }

        private void ConnectionsChanged(object sender, ConnectionProfilesChangedEventArgs e)
        {
            if (_disposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => Reload(e?.ProfileId)));
                return;
            }
            Reload(e?.ProfileId);
        }

        private void TreeDrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node != null && DrawActiveConnectionNode(e)) return;

            DatabaseObjectMetadata item = e.Node == null ? null : e.Node.Tag as DatabaseObjectMetadata;
            if (item == null || string.IsNullOrWhiteSpace(item.Comment))
            {
                e.DrawDefault = true;
                return;
            }

            bool selected = (e.State & TreeNodeStates.Selected) != 0;
            Color backColor = selected ? SystemColors.Highlight : _tree.BackColor;
            Color nameColor = selected ? SystemColors.HighlightText : _tree.ForeColor;
            Color commentColor = selected ? SystemColors.HighlightText : SystemColors.GrayText;
            Font nameFont = _tree.Font;
            Font commentFont = TreeCommentFont(nameFont);
            string name = item.Name ?? string.Empty;
            string comment = "  " + item.Comment.Trim();
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine;
            using (SolidBrush brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }
            TextRenderer.DrawText(e.Graphics, name, nameFont, e.Bounds, nameColor, backColor, flags | TextFormatFlags.EndEllipsis);
            int nameWidth = TextRenderer.MeasureText(e.Graphics, name, nameFont, new Size(int.MaxValue, e.Bounds.Height), flags).Width;
            Rectangle commentBounds = new Rectangle(e.Bounds.X + nameWidth, e.Bounds.Y, Math.Max(0, e.Bounds.Width - nameWidth), e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, comment, commentFont, commentBounds, commentColor, backColor, flags | TextFormatFlags.EndEllipsis);
        }

        private bool DrawActiveConnectionNode(DrawTreeNodeEventArgs e)
        {
            ConnectionProfileSnapshot connection = e.Node.Tag as ConnectionProfileSnapshot;
            if (connection == null) return false;
            if (!string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase)) return false;

            bool selected = (e.State & TreeNodeStates.Selected) != 0;
            Color backColor = selected ? SystemColors.Highlight : Color.FromArgb(255, 243, 205);
            Color nameColor = selected ? SystemColors.HighlightText : _tree.ForeColor;
            Color markColor = selected ? SystemColors.HighlightText : SystemColors.GrayText;
            Font nameFont = TreeActiveNameFont(_tree.Font);
            Font markFont = _tree.Font;
            string name = connection.Name ?? string.Empty;
            const string mark = " [当前]";
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine;
            Rectangle row = new Rectangle(e.Bounds.X, e.Bounds.Y, Math.Max(0, _tree.ClientSize.Width - e.Bounds.X), e.Bounds.Height);
            using (SolidBrush brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, row);
            }

            int markWidth = TextRenderer.MeasureText(e.Graphics, mark, markFont, new Size(int.MaxValue, row.Height), flags).Width;
            int measuredName = TextRenderer.MeasureText(e.Graphics, name, nameFont, new Size(int.MaxValue, row.Height), flags).Width;
            int nameWidth = Math.Min(measuredName, Math.Max(0, row.Width - markWidth));
            Rectangle nameBounds = new Rectangle(row.X, row.Y, nameWidth, row.Height);
            Rectangle markBounds = new Rectangle(row.X + nameWidth, row.Y, Math.Max(0, row.Width - nameWidth), row.Height);
            TextRenderer.DrawText(e.Graphics, name, nameFont, nameBounds, nameColor, backColor, flags | TextFormatFlags.EndEllipsis);
            TextRenderer.DrawText(e.Graphics, mark, markFont, markBounds, markColor, backColor, flags | TextFormatFlags.EndEllipsis);
            return true;
        }

        private Font TreeActiveNameFont(Font treeFont)
        {
            if (_activeNameFont != null
                && _activeNameFont.FontFamily.Name == treeFont.FontFamily.Name
                && Math.Abs(_activeNameFont.Size - treeFont.Size) < 0.1f)
            {
                return _activeNameFont;
            }

            if (_activeNameFont != null) _activeNameFont.Dispose();
            _activeNameFont = new Font(treeFont, FontStyle.Bold);
            return _activeNameFont;
        }

        private Font TreeCommentFont(Font treeFont)
        {
            if (_commentFont != null
                && _commentFont.FontFamily.Name == treeFont.FontFamily.Name
                && Math.Abs(_commentFont.Size - treeFont.Size) < 0.1f)
            {
                return _commentFont;
            }

            if (_commentFont != null) _commentFont.Dispose();
            _commentFont = new Font(treeFont, FontStyle.Italic);
            return _commentFont;
        }

        private void TreeNodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            _tree.SelectedNode = e.Node;
            if (e.Button != MouseButtons.Right || e.Node == null) return;
            if (e.Node.Parent == null && e.Node.Tag is ConnectionProfileSnapshot)
            {
                _nodeMenu.Show(_tree, e.Location);
                return;
            }
            _copyMenu.Show(_tree, e.Location);
        }

        private void TreeMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            if (_tree.GetNodeAt(e.Location) != null) return;
            _blankMenu.Show(_tree, e.Location);
        }

        private void TreeNodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Left || e.Node == null) return;
            DatabaseObjectMetadata item = e.Node.Tag as DatabaseObjectMetadata;
            if (item != null)
            {
                ObjectActivated?.Invoke(this, new DatabaseObjectActivatedEventArgs(item.Name ?? string.Empty));
                return;
            }
            if (e.Node.Parent != null) return;
            if (!(e.Node.Tag is ConnectionProfileSnapshot)) return;
            _tree.SelectedNode = e.Node;
            SetSelectedAsCurrent();
        }

        private void TreeAfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                DatabaseNodeTag databaseTag = e.Node?.Tag as DatabaseNodeTag;
                if (databaseTag != null)
                {
                    if (!string.Equals(_activeConnection?.Id, databaseTag.ConnectionId, StringComparison.OrdinalIgnoreCase))
                        return;
                    _activeDatabase = databaseTag.DatabaseName;
                    ApplyTreeVisualState();
                    SetNotice("当前数据库：" + _activeDatabase);
                    return;
                }

                ConnectionProfileSnapshot connection = e.Node?.Tag as ConnectionProfileSnapshot;
                if (connection != null && string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase))
                {
                    _activeDatabase = null;
                    ApplyTreeVisualState();
                    SetNotice("当前连接已激活：" + connection.Name);
                }
            }
            finally
            {
                SelectionChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        private async void TreeBeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            ConnectionProfileSnapshot connection = e.Node?.Tag as ConnectionProfileSnapshot;
            if (connection != null)
            {
                if (!string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase))
                {
                    e.Cancel = true;
                    await OnUiAsync(() => SetNotice("请先将该连接设置为当前连接。"));
                    return;
                }

                if (HasPlaceholder(e.Node, PlaceholderKinds.Databases))
                {
                    await LoadDatabasesAsync(e.Node, connection);
                }
                return;
            }

            DatabaseNodeTag databaseTag = e.Node?.Tag as DatabaseNodeTag;
            if (databaseTag != null && HasPlaceholder(e.Node, PlaceholderKinds.Objects))
            {
                await LoadDatabaseObjectsAsync(e.Node, databaseTag);
            }
        }

        private async Task LoadDatabasesAsync(TreeNode connectionNode, ConnectionProfileSnapshot connection)
        {
            IDatabaseMetadataProvider provider = _providers.GetById(connection.ProviderId) as IDatabaseMetadataProvider;
            if (provider == null)
            {
                await OnUiAsync(() => SetNodeMessage(connectionNode, "当前连接不支持数据库浏览。", PlaceholderKinds.Message));
                return;
            }

            await OnUiAsync(() => SetNodeMessage(connectionNode, "正在加载数据库…", PlaceholderKinds.Loading));
            CancellationToken token = RenewMetadataCancellationToken();
            try
            {
                IReadOnlyList<DatabaseMetadata> databases = await provider.GetDatabasesAsync(connection, token).ConfigureAwait(true);
                await OnUiAsync(() =>
                {
                    connectionNode.Nodes.Clear();
                    foreach (DatabaseMetadata database in databases.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase))
                    {
                        TreeNode databaseNode = new TreeNode(database.Name)
                        {
                            Tag = new DatabaseNodeTag(connection.Id, database.Name),
                            ImageKey = DatabaseTreeImageCatalog.Database(false),
                            SelectedImageKey = DatabaseTreeImageCatalog.Database(false)
                        };
                        databaseNode.Nodes.Add(new TreeNode("展开以加载对象") { Tag = new PlaceholderTag(PlaceholderKinds.Objects) });
                        connectionNode.Nodes.Add(databaseNode);
                    }

                    if (connectionNode.Nodes.Count == 0)
                    {
                        SetNodeMessage(connectionNode, "未发现可见数据库。", PlaceholderKinds.Message);
                    }
                    ApplyTreeVisualState();
                    SetNotice("数据库列表已更新：" + connection.Name);
                });
            }
            catch (OperationCanceledException)
            {
                await OnUiAsync(() => SetNodeMessage(connectionNode, "数据库加载已取消。", PlaceholderKinds.Message));
            }
            catch (Exception exception)
            {
                await OnUiAsync(() => SetNodeMessage(connectionNode, "数据库加载失败：" + exception.Message, PlaceholderKinds.Message));
            }
        }

        private async Task LoadDatabaseObjectsAsync(TreeNode databaseNode, DatabaseNodeTag databaseTag)
        {
            ConnectionProfileSnapshot connection = _connections.GetById(databaseTag.ConnectionId);
            if (connection == null)
            {
                await OnUiAsync(() => SetNodeMessage(databaseNode, "连接方案不存在。", PlaceholderKinds.Message));
                return;
            }

            IDatabaseMetadataProvider provider = _providers.GetById(connection.ProviderId) as IDatabaseMetadataProvider;
            if (provider == null)
            {
                await OnUiAsync(() => SetNodeMessage(databaseNode, "当前连接不支持对象浏览。", PlaceholderKinds.Message));
                return;
            }

            await OnUiAsync(() => SetNodeMessage(databaseNode, "正在加载对象…", PlaceholderKinds.Loading));
            CancellationToken token = RenewMetadataCancellationToken();
            try
            {
                IReadOnlyList<DatabaseObjectMetadata> objects = await provider.GetDatabaseObjectsAsync(connection, databaseTag.DatabaseName, token).ConfigureAwait(true);
                await OnUiAsync(() =>
                {
                    databaseNode.Nodes.Clear();
                    foreach (IGrouping<string, DatabaseObjectMetadata> schemaGroup in objects
                        .OrderBy(item => item.SchemaName, StringComparer.CurrentCultureIgnoreCase)
                        .ThenBy(item => item.Kind)
                        .ThenBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                        .GroupBy(item => string.IsNullOrWhiteSpace(item.SchemaName) ? "default" : item.SchemaName, StringComparer.OrdinalIgnoreCase))
                    {
                        TreeNode schemaNode = new TreeNode(schemaGroup.Key)
                        {
                            ImageKey = DatabaseTreeImageCatalog.Database(false),
                            SelectedImageKey = DatabaseTreeImageCatalog.Database(false)
                        };
                        TreeNode tablesGroup = new TreeNode("表")
                        {
                            ImageKey = DatabaseTreeImageCatalog.Table(false),
                            SelectedImageKey = DatabaseTreeImageCatalog.Table(false)
                        };
                        TreeNode viewsGroup = new TreeNode("视图")
                        {
                            ImageKey = DatabaseTreeImageCatalog.View(false),
                            SelectedImageKey = DatabaseTreeImageCatalog.View(false)
                        };

                        foreach (DatabaseObjectMetadata item in schemaGroup)
                        {
                            bool isView = item.Kind == DatabaseObjectKind.View;
                            string caption = FormatObjectCaption(item);
                            string nodeText = string.IsNullOrWhiteSpace(item.Comment) ? caption : caption + "  ";
                            TreeNode objectNode = new TreeNode(nodeText)
                            {
                                Tag = item,
                                ToolTipText = caption,
                                ImageKey = isView ? DatabaseTreeImageCatalog.View(false) : DatabaseTreeImageCatalog.Table(false),
                                SelectedImageKey = isView ? DatabaseTreeImageCatalog.View(false) : DatabaseTreeImageCatalog.Table(false)
                            };
                            if (isView) viewsGroup.Nodes.Add(objectNode);
                            else tablesGroup.Nodes.Add(objectNode);
                        }

                        if (tablesGroup.Nodes.Count > 0) schemaNode.Nodes.Add(tablesGroup);
                        if (viewsGroup.Nodes.Count > 0) schemaNode.Nodes.Add(viewsGroup);
                        if (schemaNode.Nodes.Count > 0) databaseNode.Nodes.Add(schemaNode);
                    }

                    if (databaseNode.Nodes.Count == 0)
                    {
                        SetNodeMessage(databaseNode, "该数据库无可见对象。", PlaceholderKinds.Message);
                    }
                    ApplyTreeVisualState();
                });
            }
            catch (OperationCanceledException)
            {
                await OnUiAsync(() => SetNodeMessage(databaseNode, "对象加载已取消。", PlaceholderKinds.Message));
            }
            catch (Exception exception)
            {
                await OnUiAsync(() => SetNodeMessage(databaseNode, "对象加载失败：" + exception.Message, PlaceholderKinds.Message));
            }
        }

        private void EnsureDatabasePlaceholder(TreeNode connectionNode)
        {
            if (connectionNode == null || connectionNode.Nodes.Count > 0) return;
            connectionNode.Nodes.Add(new TreeNode("展开以加载数据库") { Tag = new PlaceholderTag(PlaceholderKinds.Databases) });
        }

        private static bool HasPlaceholder(TreeNode node, string kind)
        {
            return node != null && node.Nodes.Count == 1 && (node.Nodes[0].Tag as PlaceholderTag)?.Kind == kind;
        }

        private static void SetNodeMessage(TreeNode node, string message, string kind)
        {
            if (node == null) return;
            node.Nodes.Clear();
            node.Nodes.Add(new TreeNode(message) { Tag = new PlaceholderTag(kind) });
        }

        private CancellationToken RenewMetadataCancellationToken()
        {
            _metadataCancellation?.Cancel();
            _metadataCancellation?.Dispose();
            _metadataCancellation = new CancellationTokenSource();
            return _metadataCancellation.Token;
        }

        private void ApplyTreeVisualState()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(ApplyTreeVisualState));
                return;
            }

            foreach (TreeNode root in _tree.Nodes)
            {
                ConnectionProfileSnapshot connection = root.Tag as ConnectionProfileSnapshot;
                if (connection == null) continue;
                bool connectionActive = string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase);
                string providerKey = DatabaseTreeImageCatalog.Provider(connection.ProviderId, connectionActive);
                root.Text = connectionActive ? connection.Name + " [当前]" : connection.Name;
                root.ImageKey = providerKey;
                root.SelectedImageKey = providerKey;
                ApplyChildrenState(root, connectionActive);
            }
        }

        private void ApplyChildrenState(TreeNode node, bool connectionActive)
        {
            foreach (TreeNode child in node.Nodes)
            {
                DatabaseNodeTag databaseTag = child.Tag as DatabaseNodeTag;
                if (databaseTag != null)
                {
                    bool databaseActive = connectionActive && string.Equals(_activeDatabase, databaseTag.DatabaseName, StringComparison.OrdinalIgnoreCase);
                    string key = DatabaseTreeImageCatalog.Database(databaseActive);
                    child.ImageKey = key;
                    child.SelectedImageKey = key;
                    ApplyNestedObjectState(child, databaseActive);
                    continue;
                }

                if (child.Tag is DatabaseObjectMetadata objectTag)
                {
                    string objectKey = objectTag.Kind == DatabaseObjectKind.View
                        ? DatabaseTreeImageCatalog.View(false)
                        : DatabaseTreeImageCatalog.Table(false);
                    child.ImageKey = objectKey;
                    child.SelectedImageKey = objectKey;
                }

                ApplyChildrenState(child, connectionActive);
            }
        }

        private void ApplyNestedObjectState(TreeNode databaseNode, bool databaseActive)
        {
            foreach (TreeNode schemaNode in databaseNode.Nodes)
            {
                schemaNode.ImageKey = DatabaseTreeImageCatalog.Database(databaseActive);
                schemaNode.SelectedImageKey = schemaNode.ImageKey;
                foreach (TreeNode groupNode in schemaNode.Nodes)
                {
                    bool viewGroup = string.Equals(groupNode.Text, "视图", StringComparison.Ordinal);
                    groupNode.ImageKey = viewGroup
                        ? DatabaseTreeImageCatalog.View(databaseActive)
                        : DatabaseTreeImageCatalog.Table(databaseActive);
                    groupNode.SelectedImageKey = groupNode.ImageKey;
                    foreach (TreeNode objectNode in groupNode.Nodes)
                    {
                        bool isView = viewGroup || (objectNode.Tag as DatabaseObjectMetadata)?.Kind == DatabaseObjectKind.View;
                        objectNode.ImageKey = isView
                            ? DatabaseTreeImageCatalog.View(databaseActive)
                            : DatabaseTreeImageCatalog.Table(databaseActive);
                        objectNode.SelectedImageKey = objectNode.ImageKey;
                    }
                }
            }
        }

        private Task OnUiAsync(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (_disposed) return Task.CompletedTask;
            if (!InvokeRequired)
            {
                action();
                return Task.CompletedTask;
            }

            TaskCompletionSource<object> completion = new TaskCompletionSource<object>();
            BeginInvoke(new Action(() =>
            {
                try
                {
                    if (!_disposed) action();
                    completion.SetResult(null);
                }
                catch (Exception exception)
                {
                    completion.SetException(exception);
                }
            }));
            return completion.Task;
        }

        private void ControlDisposed(object sender, EventArgs e)
        {
            if (_disposed) return;
            _disposed = true;
            _metadataCancellation?.Cancel();
            _metadataCancellation?.Dispose();
            _metadataCancellation = null;
            _connections.Changed -= ConnectionsChanged;
            _settings.Changed -= SettingsChanged;
            _tree.DrawNode -= TreeDrawNode;
            _tree.NodeMouseClick -= TreeNodeMouseClick;
            _tree.MouseUp -= TreeMouseUp;
            _tree.NodeMouseDoubleClick -= TreeNodeMouseDoubleClick;
            _tree.BeforeExpand -= TreeBeforeExpand;
            _tree.AfterSelect -= TreeAfterSelect;
            _nodeMenu.Dispose();
            _blankMenu.Dispose();
            _copyMenu.Dispose();
            _images.Dispose();
            if (_commentFont != null) _commentFont.Dispose();
            if (_activeNameFont != null) _activeNameFont.Dispose();
        }

        private sealed class DatabaseNodeTag
        {
            public DatabaseNodeTag(string connectionId, string databaseName)
            {
                ConnectionId = connectionId;
                DatabaseName = databaseName;
            }

            public string ConnectionId { get; }
            public string DatabaseName { get; }
        }

        private sealed class PlaceholderTag
        {
            public PlaceholderTag(string kind)
            {
                Kind = kind;
            }

            public string Kind { get; }
        }

        private static class PlaceholderKinds
        {
            public const string Databases = "databases";
            public const string Objects = "objects";
            public const string Loading = "loading";
            public const string Message = "message";
        }
    }

    /// <summary>双击表或视图时传出的对象名称。</summary>
    public sealed class DatabaseObjectActivatedEventArgs : EventArgs
    {
        /// <summary>创建事件参数。</summary>
        /// <param name="objectName">表或视图编码，不含注释。</param>
        public DatabaseObjectActivatedEventArgs(string objectName)
        {
            ObjectName = objectName ?? string.Empty;
        }

        /// <summary>获取对象编码。</summary>
        public string ObjectName { get; }
    }
}

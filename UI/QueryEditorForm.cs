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
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.UI
{
    /// <summary>提供 SQL 方案编辑、只读查询预览以及结果写入 Excel 的主界面。</summary>
    /// <remarks>
    /// 用户输入包括连接方案、SQL、方案名称和目标 Sheet；输出副作用包括保存查询方案和修改当前工作簿。
    /// </remarks>
    public sealed class QueryEditorForm : AppForm
    {
        private readonly IQueryProfileRepository _queries;
        private readonly IConnectionProfileRepository _connections;
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IQueryBufferService _queryBuffer;
        private readonly IExcelResultWriter _writer;
        private readonly IOperationRunner _operationRunner;
        private readonly ISettingsStore _settings;
        private readonly ExcelInterop.Workbook _workbook;
        private readonly TreeView _connectionsTree;
        private readonly ListBox _queryProfiles;
        private readonly SqlEditorControl _sql;
        private readonly DataGridView _preview;
        private readonly Label _queryInfoLabel;
        private readonly Label _sqlQueryNameLabel;
        private readonly ToolStripStatusLabel _status;
        private readonly SplitContainer _resultsSplit;
        private readonly ImageList _treeImages;
        private readonly ContextMenuStrip _connectionNodeMenu;
        private readonly ContextMenuStrip _connectionBlankMenu;
        private readonly ContextMenuStrip _queryProfileMenu;
        private readonly Font _previewHeaderTypeFont;
        private readonly ToolStripMenuItem _editConnectionMenuItem;
        private readonly ToolStripMenuItem _deleteConnectionMenuItem;
        private readonly ToolStripMenuItem _setCurrentConnectionMenuItem;
        private ConnectionProfileSnapshot _activeConnection;
        private string _activeDatabase;
        private string _currentQueryId;
        private string _currentQueryName;
        private string _currentTargetSheet = "查询结果";
        private DateTime _currentCreatedUtc;
        private DateTime? _currentSavedUtc;
        private bool _hasPreview;
        private CancellationTokenSource _metadataCancellation;

        private Task OnUiAsync(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (IsDisposed) return Task.CompletedTask;

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
                    if (!IsDisposed) action();
                    completion.SetResult(null);
                }
                catch (Exception exception)
                {
                    completion.SetException(exception);
                }
            }));
            return completion.Task;
        }

        private void OnUi(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (IsDisposed) return;

            if (!InvokeRequired)
            {
                action();
                return;
            }

            BeginInvoke(action);
        }

        /// <summary>创建 SQL 查询编辑窗体。</summary>
        /// <param name="queries">查询方案仓储。</param>
        /// <param name="connections">连接方案仓储。</param>
        /// <param name="providers">数据源提供程序注册表。</param>
        /// <param name="executionService">查询执行和连接测试服务。</param>
        /// <param name="queryBuffer">用于预览和写入前缓冲查询结果。</param>
        /// <param name="writer">Excel 结果写入器。</param>
        /// <param name="operationRunner">进度窗体运行器。</param>
        /// <param name="settings">查询行数、超时和分块大小设置。</param>
        /// <param name="workbook">当前操作的 Excel 工作簿。</param>
        public QueryEditorForm(
            IQueryProfileRepository queries,
            IConnectionProfileRepository connections,
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IQueryBufferService queryBuffer,
            IExcelResultWriter writer,
            IOperationRunner operationRunner,
            ISettingsStore settings,
            ExcelInterop.Workbook workbook)
        {
            _queries = queries ?? throw new ArgumentNullException(nameof(queries));
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _queryBuffer = queryBuffer ?? throw new ArgumentNullException(nameof(queryBuffer));
            _writer = writer ?? throw new ArgumentNullException(nameof(writer));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));

            Text = AppPresentation.WindowTitle("SQL 查询");
            Rectangle workingArea = Screen.FromControl(this).WorkingArea;
            Width = Math.Min(1320, workingArea.Width - 40);
            Height = Math.Min(860, workingArea.Height - 60);
            MinimumSize = new Size(820, 600);
            StartPosition = FormStartPosition.CenterScreen;

            _treeImages = DatabaseTreeImageCatalog.CreateImageList();
            _connectionsTree = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false,
                ImageList = _treeImages,
                ShowNodeToolTips = true
            };
            _queryProfiles = new ListBox
            {
                Dock = DockStyle.Fill,
                DisplayMember = nameof(QueryItem.DisplayName),
                IntegralHeight = false
            };
            _sql = new SqlEditorControl(new SqlFormattingService()) { Dock = DockStyle.Fill };
            _previewHeaderTypeFont = new Font(AppPresentation.DefaultFontName, AppPresentation.DefaultFontSize, FontStyle.Italic);
            _preview = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 42,
                RowHeadersVisible = true,
                RowHeadersWidth = 56,
                BackgroundColor = SystemColors.Control,
                EnableHeadersVisualStyles = false,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.Gainsboro,
                    ForeColor = SystemColors.ControlText,
                    Font = new Font(AppPresentation.DefaultFontName, AppPresentation.DefaultFontSize, FontStyle.Bold),
                    WrapMode = DataGridViewTriState.True
                }
            };
            _queryInfoLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 70,
                Padding = new Padding(8, 8, 8, 6),
                ForeColor = SystemColors.GrayText,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "查询方案名称：未保存" + Environment.NewLine + "连接方案：未选择" + Environment.NewLine + "保存时间：未保存"
            };
            _sqlQueryNameLabel = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(0, 0, 8, 0),
                Text = "当前方案：未保存"
            };
            _status = new ToolStripStatusLabel { Text = "就绪。", ForeColor = SystemColors.GrayText, Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            _connectionNodeMenu = BuildConnectionNodeMenu();
            _connectionBlankMenu = BuildConnectionBlankMenu();
            _queryProfileMenu = BuildQueryProfileMenu();

            _setCurrentConnectionMenuItem = new ToolStripMenuItem("设置为当前连接", null, (sender, args) => SetSelectedConnectionAsCurrent());
            _editConnectionMenuItem = new ToolStripMenuItem("编辑连接", null, (sender, args) => EditSelectedConnection());
            _deleteConnectionMenuItem = new ToolStripMenuItem("删除连接", null, (sender, args) => DeleteSelectedConnection());
            ToolStripMenuItem newConnectionMenuItem = new ToolStripMenuItem("新建连接", null, (sender, args) => CreateConnection());
            ToolStripMenuItem refreshConnectionMenuItem = new ToolStripMenuItem("刷新", null, (sender, args) => RefreshSelectedConnection());
            ToolStripMenuItem connectionMenu = new ToolStripMenuItem("连接");
            connectionMenu.DropDownItems.AddRange(new ToolStripItem[]
            {
                newConnectionMenuItem,
                _editConnectionMenuItem,
                _deleteConnectionMenuItem,
                new ToolStripSeparator(),
                _setCurrentConnectionMenuItem,
                refreshConnectionMenuItem
            });

            ToolStripMenuItem newMenuItem = new ToolStripMenuItem("新建查询", null, (sender, args) => NewQuery()) { ShortcutKeys = Keys.Control | Keys.N };
            ToolStripMenuItem saveMenuItem = new ToolStripMenuItem("保存方案", null, (sender, args) => SaveQuery()) { ShortcutKeys = Keys.Control | Keys.S };
            ToolStripMenuItem deleteMenuItem = new ToolStripMenuItem("删除方案", null, (sender, args) => DeleteQuery());
            ToolStripMenuItem previewMenuItem = new ToolStripMenuItem("预览", null, (sender, args) => PreviewQuery()) { ShortcutKeys = Keys.Control | Keys.Enter };
            ToolStripMenuItem exportMenuItem = new ToolStripMenuItem("写入 Sheet", null, (sender, args) => ExportQuery()) { ShortcutKeys = Keys.Control | Keys.Shift | Keys.E };
            ToolStripMenuItem closeMenuItem = new ToolStripMenuItem("关闭", null, (sender, args) => Close()) { ShortcutKeys = Keys.Alt | Keys.F4 };
            ToolStripMenuItem queryMenu = new ToolStripMenuItem("查询");
            queryMenu.DropDownItems.AddRange(new ToolStripItem[]
            {
                newMenuItem,
                saveMenuItem,
                deleteMenuItem,
                new ToolStripSeparator(),
                previewMenuItem,
                exportMenuItem,
                new ToolStripSeparator(),
                closeMenuItem
            });

            ToolStripMenuItem formatMenuItem = new ToolStripMenuItem("格式化 SQL", null, (sender, args) => _sql.FormatSql()) { ShortcutKeys = Keys.Control | Keys.Shift | Keys.F };
            ToolStripMenuItem commentMenuItem = new ToolStripMenuItem("注释/取消注释", null, (sender, args) => _sql.ToggleComment()) { ShortcutKeys = Keys.Control | Keys.OemQuestion };
            ToolStripMenuItem editMenu = new ToolStripMenuItem("编辑");
            editMenu.DropDownItems.AddRange(new ToolStripItem[] { formatMenuItem, commentMenuItem });

            MenuStrip menu = new MenuStrip { Dock = DockStyle.Top };
            menu.Items.Add(connectionMenu);
            menu.Items.Add(queryMenu);
            menu.Items.Add(editMenu);
            MainMenuStrip = menu;

            Button previewButton = ButtonOf("预览");
            Button exportButton = ButtonOf("写入 Sheet");
            Button saveButton = ButtonOf("保存方案");
            FlowLayoutPanel sqlActions = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Anchor = AnchorStyles.Right,
                Padding = Padding.Empty,
                Margin = Padding.Empty
            };
            sqlActions.Controls.Add(previewButton);
            sqlActions.Controls.Add(exportButton);
            sqlActions.Controls.Add(saveButton);
            Panel sqlHeader = BuildSqlHeader(sqlActions);

            Panel sqlPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 30, 0, 0) };
            sqlPanel.Controls.Add(_sql);
            sqlPanel.Controls.Add(sqlHeader);

            Panel previewPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 30, 0, 0) };
            previewPanel.Controls.Add(_preview);
            previewPanel.Controls.Add(SectionHeader("结果预览"));

            _resultsSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                Panel2Collapsed = true,
                FixedPanel = FixedPanel.None
            };
            _resultsSplit.Panel1.Controls.Add(sqlPanel);
            _resultsSplit.Panel2.Controls.Add(previewPanel);

            SplitContainer leftSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal
            };
            leftSplit.Panel1.Controls.Add(SectionPanel("数据库连接", _connectionsTree));
            leftSplit.Panel2.Controls.Add(SectionPanel("查询方案", BuildQueryProfilesPanel()));

            SplitContainer workspace = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6
            };
            workspace.Panel1.Controls.Add(leftSplit);
            workspace.Panel2.Controls.Add(_resultsSplit);

            StatusStrip statusStrip = new StatusStrip { Dock = DockStyle.Bottom, SizingGrip = false };
            statusStrip.Items.Add(_status);

            Controls.Add(workspace);
            Controls.Add(statusStrip);
            Controls.Add(menu);

            previewButton.Click += (sender, args) => PreviewQuery();
            exportButton.Click += (sender, args) => ExportQuery();
            saveButton.Click += (sender, args) => SaveQuery();
            _sql.ExecuteSelectionRequested += (sender, args) => PreviewSelectedQuery();
            _queryProfiles.SelectedIndexChanged += QueryProfilesSelectedIndexChanged;
            _queryProfiles.MouseUp += QueryProfilesMouseUp;
            _preview.RowPostPaint += PreviewRowPostPaint;
            _preview.CellPainting += PreviewCellPainting;
            _connectionsTree.NodeMouseClick += ConnectionsTreeNodeMouseClick;
            _connectionsTree.MouseUp += ConnectionsTreeMouseUp;
            _connectionsTree.NodeMouseDoubleClick += ConnectionsTreeNodeMouseDoubleClick;
            _connectionsTree.BeforeExpand += ConnectionsTreeBeforeExpand;
            _connectionsTree.AfterSelect += ConnectionsTreeAfterSelect;
            _connections.Changed += ConnectionsChanged;
            FormClosed += QueryEditorFormFormClosed;
            workspace.SizeChanged += (sender, args) => ApplySplitterPanelMinSizes(workspace, 220, 420);
            Shown += (sender, args) =>
            {
                BeginInvoke(new Action(() =>
                {
                    ApplySplitterPanelMinSizes(workspace, 220, 420);
                    TrySetSplitterDistance(workspace, workspace.Width / 4);
                    TrySetSplitterDistance(leftSplit, leftSplit.Height / 2);
                }));
            };

            RestoreActiveConnection();
            ReloadConnectionsTree();
            ReloadQueries(null, false);
            NewQuery();
            UpdateQueryInfo();
            UpdateConnectionMenuState();
        }

        private void ReloadQueries(string selectedId = null, bool activateSelection = true)
        {
            List<QueryItem> items = _queries.GetAll().Select(profile => new QueryItem(profile)).ToList();
            _queryProfiles.DataSource = null;
            _queryProfiles.DataSource = items;
            _queryProfiles.DisplayMember = nameof(QueryItem.DisplayName);

            if (!activateSelection)
            {
                _queryProfiles.SelectedIndex = -1;
                return;
            }

            QueryItem target = null;
            if (!string.IsNullOrWhiteSpace(selectedId))
            {
                target = items.FirstOrDefault(item =>
                    string.Equals(item.Profile.Id, selectedId, StringComparison.OrdinalIgnoreCase));
            }
            else if (items.Count > 0)
            {
                target = items
                    .OrderByDescending(item => item.Profile.UpdatedUtc)
                    .ThenBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                    .FirstOrDefault();
            }

            if (target != null)
                _queryProfiles.SelectedItem = target;

            if (_queryProfiles.SelectedIndex >= 0)
                return;

            if (items.Count == 0)
            {
                NewQuery();
                return;
            }

            _queryProfiles.SelectedIndex = 0;
        }

        private void NewQuery()
        {
            _queryProfiles.SelectedIndex = -1;
            _currentQueryId = Guid.NewGuid().ToString("N");
            _currentCreatedUtc = DateTime.UtcNow;
            _currentQueryName = string.Empty;
            _currentTargetSheet = "查询结果";
            _currentSavedUtc = null;
            _sql.ClearAndFocus();
            ClearPreviewGrid();
            _hasPreview = false;
            _resultsSplit.Panel2Collapsed = true;
            _status.Text = "新查询方案。";
            UpdateQueryInfo();
        }

        private void QueryProfilesSelectedIndexChanged(object sender, EventArgs e)
        {
            QueryProfile profile = (_queryProfiles.SelectedItem as QueryItem)?.Profile;
            if (profile == null) return;
            _currentQueryId = profile.Id;
            _currentCreatedUtc = profile.CreatedUtc;
            _currentQueryName = profile.Name;
            _sql.SqlText = profile.QueryText;
            _currentTargetSheet = string.IsNullOrWhiteSpace(profile.TargetSheetName) ? "查询结果" : profile.TargetSheetName;
            _currentSavedUtc = profile.UpdatedUtc;
            _activeConnection = _connections.GetById(profile.ConnectionProfileId);
            RememberActiveConnection(_activeConnection);
            ApplyTreeVisualState();
            ClearPreviewGrid();
            _hasPreview = false;
            _resultsSplit.Panel2Collapsed = true;
            _status.Text = "已加载查询方案。";
            UpdateQueryInfo();
        }

        private void SaveQuery()
        {
            if (!ValidateInputs(false, out ConnectionProfileSnapshot connection)) return;
            if (!TextPromptDialog.TryShow(
                this,
                "保存查询方案",
                "请输入方案名称",
                string.IsNullOrWhiteSpace(_currentQueryName) ? "新查询方案" : _currentQueryName,
                out string queryName))
            {
                _status.Text = "已取消保存方案。";
                return;
            }

            _currentQueryName = queryName;
            QueryProfile duplicate = _queries.GetAll().FirstOrDefault(item =>
                string.Equals(item.Name, _currentQueryName, StringComparison.CurrentCultureIgnoreCase) &&
                !string.Equals(item.Id, _currentQueryId, StringComparison.OrdinalIgnoreCase));
            bool overwritten = false;
            if (duplicate != null)
            {
                DialogResult overwrite = MessageBox.Show(
                    this,
                    string.Format("查询方案“{0}”已存在，是否覆盖？", _currentQueryName),
                    "保存查询方案",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
                if (overwrite != DialogResult.Yes)
                {
                    _status.Text = "已取消覆盖保存。";
                    return;
                }

                _currentQueryId = duplicate.Id;
                _currentCreatedUtc = duplicate.CreatedUtc;
                overwritten = true;
            }

            QueryProfile profile = new QueryProfile
            {
                Id = _currentQueryId ?? Guid.NewGuid().ToString("N"),
                Name = _currentQueryName,
                ProviderId = connection.ProviderId,
                ConnectionProfileId = connection.Id,
                QueryText = _sql.SqlText,
                TargetSheetName = _currentTargetSheet,
                CreatedUtc = _currentCreatedUtc == default(DateTime) ? DateTime.UtcNow : _currentCreatedUtc,
                UpdatedUtc = DateTime.UtcNow
            };
            _queries.Save(profile);
            _currentQueryId = profile.Id;
            _currentCreatedUtc = profile.CreatedUtc;
            _currentSavedUtc = profile.UpdatedUtc;
            ReloadQueries(profile.Id);
            _status.Text = overwritten ? "查询方案已覆盖保存。" : "查询方案已保存。";
            UpdateQueryInfo();
        }

        private void DeleteQuery()
        {
            QueryProfile profile = (_queryProfiles.SelectedItem as QueryItem)?.Profile;
            if (profile == null) return;
            if (MessageBox.Show(this, "确定删除查询方案“" + profile.Name + "”吗？", "删除查询方案",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            _queries.Delete(profile.Id);
            ReloadQueries();
            NewQuery();
        }

        private void UpdateQueryInfo()
        {
            string name = string.IsNullOrWhiteSpace(_currentQueryName) ? "未保存" : _currentQueryName;
            string connectionName = "未选择";

            if (!string.IsNullOrWhiteSpace(_currentQueryId))
            {
                QueryProfile profile = _queries.GetById(_currentQueryId);
                if (profile != null)
                {
                    ConnectionProfileSnapshot connection = _connections.GetById(profile.ConnectionProfileId);
                    if (!string.IsNullOrWhiteSpace(connection?.Name)) connectionName = connection.Name;
                }
            }

            if (connectionName == "未选择" && !string.IsNullOrWhiteSpace(_activeConnection?.Name))
                connectionName = _activeConnection.Name;

            string savedTime = _currentSavedUtc.HasValue
                ? _currentSavedUtc.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                : "未保存";

            _queryInfoLabel.Text = string.Format(
                "查询方案名称：{0}{3}连接方案：{1}{3}保存时间：{2}",
                name,
                connectionName,
                savedTime,
                Environment.NewLine);
            _sqlQueryNameLabel.Text = "当前方案：" + name;
        }

        private Control BuildQueryProfilesPanel()
        {
            Panel panel = new Panel { Dock = DockStyle.Fill };
            panel.Controls.Add(_queryProfiles);
            panel.Controls.Add(_queryInfoLabel);
            return panel;
        }

        private Panel BuildSqlHeader(Control actions)
        {
            Panel header = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = SystemColors.ControlLight };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = new Padding(8, 0, 4, 0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            Label title = new Label
            {
                Text = "SQL编辑器（仅限查询）",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 12, 0),
                Font = new Font(AppPresentation.DefaultFontName, AppPresentation.DefaultFontSize, FontStyle.Bold)
            };
            _sqlQueryNameLabel.Margin = new Padding(0, 6, 8, 0);
            actions.Margin = new Padding(0, 2, 0, 0);
            layout.Controls.Add(title, 0, 0);
            layout.Controls.Add(_sqlQueryNameLabel, 1, 0);
            layout.Controls.Add(actions, 2, 0);
            header.Controls.Add(layout);
            return header;
        }

        private void PreviewQuery()
        {
            ExecutePreview(false);
        }

        private void PreviewSelectedQuery()
        {
            ExecutePreview(true);
        }

        private void ExecutePreview(bool selectedOnly)
        {
            string sqlText = selectedOnly ? _sql.SelectedSql : _sql.SqlText;
            if (!ValidateInputs(false, out ConnectionProfileSnapshot connection, sqlText)) return;
            bool hadPreview = _hasPreview;
            try
            {
                DataSourceRequest request = CreateRequest(connection, sqlText, ExecutionPurpose.Preview, _settings.Get(CoreSettings.MaxPreviewRows));
                BufferedQueryResult result = _operationRunner.Run(
                    this, "查询预览", true,
                    (progress, token) => _queryBuffer.ExecuteAsync(request, Guid.NewGuid().ToString("N"), progress, token));
                BindPreview(result);
                _hasPreview = true;
                EnsurePreviewVisible();
                int previewLimit = _settings.Get(CoreSettings.MaxPreviewRows);
                _status.Text = result.IsTruncated
                    ? string.Format(
                        "{0}完成，共 {1:N0} 行，已达上限 {2:N0} 行并截断。可在设置中修改“最大预览行数”。",
                        selectedOnly ? "选中文本预览" : "预览",
                        result.RowCount,
                        previewLimit)
                    : string.Format("{0}完成，共 {1:N0} 行。",
                        selectedOnly ? "选中文本预览" : "预览",
                        result.RowCount);
            }
            catch (OperationCanceledException)
            {
                if (!hadPreview)
                {
                    _hasPreview = false;
                    _resultsSplit.Panel2Collapsed = true;
                }
                _status.Text = selectedOnly ? "选中文本预览已取消。" : "查询预览已取消。";
            }
            catch (Exception exception)
            {
                if (!hadPreview)
                {
                    _hasPreview = false;
                    _resultsSplit.Panel2Collapsed = true;
                }
                ShowError(exception);
            }
        }

        private void ExportQuery()
        {
            if (!TextPromptDialog.TryShow(
                this,
                "写入 Sheet",
                "请输入目标 Sheet 名称",
                string.IsNullOrWhiteSpace(_currentTargetSheet) ? "查询结果" : _currentTargetSheet,
                out string targetSheet))
            {
                _status.Text = "已取消写入。";
                return;
            }

            _currentTargetSheet = targetSheet;
            if (!ValidateInputs(true, out ConnectionProfileSnapshot connection)) return;
            try
            {
                DataSourceRequest request = CreateRequest(connection, _sql.SqlText, ExecutionPurpose.Export, _settings.Get(CoreSettings.MaxExportRows));
                string target = _currentTargetSheet;
                long rows = _operationRunner.Run(this, "写入 Excel", true, async (progress, token) =>
                {
                    BufferedQueryResult result = await _queryBuffer.ExecuteAsync(
                        request, Guid.NewGuid().ToString("N"), progress, token);
                    return _writer.WriteResult(_workbook, target, result, progress, token);
                });
                _status.Text = string.Format("已向“{0}”写入 {1:N0} 行。", target, rows);
            }
            catch (OperationCanceledException)
            {
                _status.Text = "写入已取消。";
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private DataSourceRequest CreateRequest(ConnectionProfileSnapshot connection, string sqlText, ExecutionPurpose purpose, int rowLimit)
        {
            return new DataSourceRequest(
                connection,
                sqlText,
                purpose,
                rowLimit,
                _settings.Get(CoreSettings.ResultBlockSize),
                _settings.Get(CoreSettings.QueryTimeoutSeconds));
        }

        private void EnsurePreviewVisible()
        {
            if (_resultsSplit.Panel2Collapsed)
                _resultsSplit.Panel2Collapsed = false;
            int available = _resultsSplit.ClientSize.Height - _resultsSplit.SplitterWidth;
            TrySetSplitterDistance(_resultsSplit, available / 2);
        }

        /// <summary>在容器拥有足够可用尺寸后应用分割面板最小尺寸，避免构造期触发 WinForms 距离范围异常。</summary>
        /// <param name="split">要设置最小尺寸约束的分割容器。</param>
        /// <param name="panel1MinSize">左侧或上方面板的最小像素尺寸。</param>
        /// <param name="panel2MinSize">右侧或下方面板的最小像素尺寸。</param>
        private static void ApplySplitterPanelMinSizes(SplitContainer split, int panel1MinSize, int panel2MinSize)
        {
            if (split == null || split.IsDisposed) return;

            int total = split.Orientation == Orientation.Vertical ? split.Width : split.Height;
            int available = total - split.SplitterWidth;
            if (available < panel1MinSize + panel2MinSize) return;

            if (split.Panel1MinSize != panel1MinSize)
                split.Panel1MinSize = panel1MinSize;
            if (split.Panel2MinSize != panel2MinSize)
                split.Panel2MinSize = panel2MinSize;
        }

        /// <summary>在满足面板最小尺寸约束时安全设置分割条位置，避免 WinForms 抛出范围异常。</summary>
        /// <param name="split">要设置分割条的容器。</param>
        /// <param name="preferred">期望的分割条位置。</param>
        private static void TrySetSplitterDistance(SplitContainer split, int preferred)
        {
            if (split == null || split.IsDisposed) return;

            int total = split.Orientation == Orientation.Vertical ? split.Width : split.Height;
            if (total <= 0) return;

            int minimum = split.Panel1MinSize;
            int maximum = total - split.Panel2MinSize - split.SplitterWidth;
            if (maximum < minimum) return;

            int distance = Math.Max(minimum, Math.Min(maximum, preferred));
            try
            {
                split.SplitterDistance = distance;
            }
            catch (ArgumentOutOfRangeException)
            {
            }
            catch (InvalidOperationException)
            {
            }
        }

        private void ClearPreviewGrid()
        {
            _preview.Rows.Clear();
            _preview.Columns.Clear();
        }

        private void BindPreview(BufferedQueryResult result)
        {
            _preview.SuspendLayout();
            try
            {
                ClearPreviewGrid();
                if (result == null || result.Columns.Count == 0)
                    return;

                Font headerFont = _preview.ColumnHeadersDefaultCellStyle.Font ?? _preview.Font;
                for (int index = 0; index < result.Columns.Count; index++)
                {
                    ResultColumn column = result.Columns[index];
                    int textWidth = TextRenderer.MeasureText(column.Name ?? string.Empty, headerFont).Width + 24;
                    _preview.Columns.Add(new DataGridViewTextBoxColumn
                    {
                        HeaderText = column.Name,
                        Name = "c" + index,
                        Tag = column,
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                        Width = Math.Max(72, Math.Min(280, textWidth))
                    });
                }

                foreach (object[,] block in result.Blocks)
                {
                    int rowCount = block.GetLength(0);
                    int columnCount = block.GetLength(1);
                    for (int row = 0; row < rowCount; row++)
                    {
                        object[] values = new object[columnCount];
                        for (int column = 0; column < columnCount; column++)
                            values[column] = block[row, column];
                        _preview.Rows.Add(values);
                    }
                }
            }
            finally
            {
                _preview.ResumeLayout();
            }
        }

        private void RestoreActiveConnection()
        {
            string id = _settings.Get(CoreSettings.ActiveConnectionId);
            if (string.IsNullOrWhiteSpace(id))
                return;

            ConnectionProfileSnapshot connection = _connections.GetById(id);
            if (connection == null)
            {
                RememberActiveConnection(null);
                return;
            }

            _activeConnection = connection;
        }

        private void RememberActiveConnection(ConnectionProfileSnapshot connection)
        {
            _settings.Set(CoreSettings.ActiveConnectionId, connection?.Id ?? string.Empty);
        }

        private bool ValidateInputs(bool requireTarget, out ConnectionProfileSnapshot connection, string sqlText = null)
        {
            connection = ResolveExecutionConnection();
            string sql = sqlText ?? _sql.SqlText;
            string message = null;
            if (connection == null) message = "请选择连接方案。";
            else if (string.IsNullOrWhiteSpace(sql)) message = "请输入 SQL。";
            else if (requireTarget && string.IsNullOrWhiteSpace(_currentTargetSheet)) message = "请输入目标 Sheet 名称。";
            if (message == null) return true;
            MessageBox.Show(this, message, "SQL 查询", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        private static string FormatPreviewTypeName(ResultColumn column)
        {
            if (column == null) return string.Empty;
            Type dataType = column.GetDataType();
            return dataType == null ? string.Empty : dataType.Name;
        }

        private ContextMenuStrip BuildConnectionNodeMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("设置为当前连接", null, (sender, args) => SetSelectedConnectionAsCurrent());
            menu.Items.Add("编辑连接", null, (sender, args) => EditSelectedConnection());
            menu.Items.Add("新建连接", null, (sender, args) => CreateConnection());
            menu.Items.Add("删除连接", null, (sender, args) => DeleteSelectedConnection());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("刷新", null, (sender, args) => RefreshSelectedConnection());
            return menu;
        }

        private ContextMenuStrip BuildConnectionBlankMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("新建连接", null, (sender, args) => CreateConnection());
            menu.Items.Add("刷新连接列表", null, (sender, args) => ReloadConnectionsTree());
            return menu;
        }

        private ContextMenuStrip BuildQueryProfileMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("新建", null, (sender, args) => NewQuery());
            menu.Items.Add("保存", null, (sender, args) => SaveQuery());
            menu.Items.Add("删除", null, (sender, args) => DeleteQuery());
            return menu;
        }

        private void ConnectionsChanged(object sender, ConnectionProfilesChangedEventArgs e)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => ReloadConnectionsTree(e?.ProfileId)));
                return;
            }
            ReloadConnectionsTree(e?.ProfileId);
        }

        private void QueryEditorFormFormClosed(object sender, FormClosedEventArgs e)
        {
            _metadataCancellation?.Cancel();
            _metadataCancellation?.Dispose();
            _metadataCancellation = null;
            _connections.Changed -= ConnectionsChanged;
            _connectionsTree.NodeMouseClick -= ConnectionsTreeNodeMouseClick;
            _connectionsTree.MouseUp -= ConnectionsTreeMouseUp;
            _connectionsTree.NodeMouseDoubleClick -= ConnectionsTreeNodeMouseDoubleClick;
            _connectionsTree.BeforeExpand -= ConnectionsTreeBeforeExpand;
            _connectionsTree.AfterSelect -= ConnectionsTreeAfterSelect;
            _queryProfiles.MouseUp -= QueryProfilesMouseUp;
            _preview.RowPostPaint -= PreviewRowPostPaint;
            _preview.CellPainting -= PreviewCellPainting;
            _connectionNodeMenu.Dispose();
            _connectionBlankMenu.Dispose();
            _queryProfileMenu.Dispose();
            _treeImages.Dispose();
            _previewHeaderTypeFont.Dispose();
        }

        private void PreviewCellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex < 0) return;

            e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
            ResultColumn column = _preview.Columns[e.ColumnIndex].Tag as ResultColumn;
            string name = column == null ? _preview.Columns[e.ColumnIndex].HeaderText : column.Name;
            string typeName = FormatPreviewTypeName(column);
            Rectangle bounds = e.CellBounds;
            bounds.Inflate(-4, -2);
            int nameHeight = Math.Max(14, bounds.Height / 2);
            Rectangle nameBounds = new Rectangle(bounds.X, bounds.Y, bounds.Width, nameHeight);
            Rectangle typeBounds = new Rectangle(bounds.X, bounds.Y + nameHeight - 1, bounds.Width, bounds.Height - nameHeight + 1);
            TextRenderer.DrawText(
                e.Graphics,
                name,
                e.CellStyle.Font,
                nameBounds,
                e.CellStyle.ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            if (!string.IsNullOrWhiteSpace(typeName))
            {
                TextRenderer.DrawText(
                    e.Graphics,
                    typeName,
                    _previewHeaderTypeFont,
                    typeBounds,
                    SystemColors.GrayText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            }

            e.Handled = true;
        }

        private void PreviewRowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            string rowNumber = (e.RowIndex + 1).ToString();
            Rectangle bounds = new Rectangle(
                e.RowBounds.Left,
                e.RowBounds.Top,
                _preview.RowHeadersWidth - 6,
                e.RowBounds.Height);
            TextRenderer.DrawText(
                e.Graphics,
                rowNumber,
                _preview.RowHeadersDefaultCellStyle.Font ?? _preview.Font,
                bounds,
                _preview.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        }

        private void QueryProfilesMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;

            int index = _queryProfiles.IndexFromPoint(e.Location);
            if (index >= 0 && index < _queryProfiles.Items.Count)
            {
                _queryProfiles.SelectedIndex = index;
            }

            _queryProfileMenu.Show(_queryProfiles, e.Location);
        }

        private void ConnectionsTreeNodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            _connectionsTree.SelectedNode = e.Node;
            if (e.Button == MouseButtons.Right && e.Node?.Parent == null && e.Node?.Tag is ConnectionProfileSnapshot)
            {
                _connectionNodeMenu.Show(_connectionsTree, e.Location);
            }
        }

        private void ConnectionsTreeMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            TreeNode node = _connectionsTree.GetNodeAt(e.Location);
            if (node != null) return;
            _connectionBlankMenu.Show(_connectionsTree, e.Location);
        }

        private void ConnectionsTreeNodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            if (e.Node?.Parent != null) return;
            if (!(e.Node?.Tag is ConnectionProfileSnapshot)) return;
            _connectionsTree.SelectedNode = e.Node;
            SetSelectedConnectionAsCurrent();
        }

        private ConnectionProfileSnapshot SelectedConnectionNode
        {
            get
            {
                TreeNode node = _connectionsTree.SelectedNode;
                return node != null && node.Parent == null
                    ? node.Tag as ConnectionProfileSnapshot
                    : null;
            }
        }

        private void ReloadConnectionsTree(string selectedConnectionId = null)
        {
            string selectedId = selectedConnectionId ?? SelectedConnectionNode?.Id ?? _activeConnection?.Id;
            List<ConnectionProfileSnapshot> connections = _connections.GetAll().OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
            _connectionsTree.BeginUpdate();
            try
            {
                _connectionsTree.Nodes.Clear();
                foreach (ConnectionProfileSnapshot connection in connections)
                {
                    bool active = string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase);
                    TreeNode node = new TreeNode(connection.Name)
                    {
                        Tag = connection,
                        ToolTipText = string.IsNullOrWhiteSpace(connection.ProviderId) ? connection.Name : connection.Name + "（" + connection.ProviderId + "）",
                        ImageKey = DatabaseTreeImageCatalog.Provider(connection.ProviderId, active),
                        SelectedImageKey = DatabaseTreeImageCatalog.Provider(connection.ProviderId, active)
                    };
                    EnsureDatabasePlaceholder(node);
                    _connectionsTree.Nodes.Add(node);
                    if (string.Equals(connection.Id, selectedId, StringComparison.OrdinalIgnoreCase))
                    {
                        _connectionsTree.SelectedNode = node;
                    }
                }
            }
            finally
            {
                _connectionsTree.EndUpdate();
            }
            ApplyTreeVisualState();
            UpdateConnectionMenuState();
        }

        private void CreateConnection()
        {
            using (ConnectionProfileEditForm editor = new ConnectionProfileEditForm(_providers, _executionService, _operationRunner, null))
            {
                if (editor.ShowDialog(this) != DialogResult.OK || editor.Profile == null) return;
                _connections.Save(editor.Profile);
                ReloadConnectionsTree(editor.Profile.Id);
                _status.Text = "连接方案已创建。";
            }
        }

        private void EditSelectedConnection()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null) return;
            bool wasActive = string.Equals(_activeConnection?.Id, selected.Id, StringComparison.OrdinalIgnoreCase);
            using (ConnectionProfileEditForm editor = new ConnectionProfileEditForm(_providers, _executionService, _operationRunner, selected))
            {
                if (editor.ShowDialog(this) != DialogResult.OK || editor.Profile == null) return;
                _connections.Save(editor.Profile);
                ReloadConnectionsTree(editor.Profile.Id);
                ConnectionProfileSnapshot updated = _connections.GetById(editor.Profile.Id);
                if (wasActive && updated != null)
                {
                    ActivateConnection(updated, true);
                }
                else
                {
                    _status.Text = "连接方案已更新。";
                }
            }
        }

        private void DeleteSelectedConnection()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null) return;
            DialogResult result = MessageBox.Show(
                this,
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
            }
            ReloadConnectionsTree();
            _status.Text = "连接方案已删除。";
        }

        private void SetSelectedConnectionAsCurrent()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null) return;
            ActivateConnection(selected, true);
        }

        private void RefreshSelectedConnection()
        {
            ConnectionProfileSnapshot selected = SelectedConnectionNode;
            if (selected == null)
            {
                ReloadConnectionsTree();
                return;
            }
            ActivateConnection(selected, false);
        }

        private async void ActivateConnection(ConnectionProfileSnapshot connection, bool expandNode)
        {
            if (connection == null) return;
            ResetActiveConnectionVisualState();
            try
            {
                ConnectionTestResult result = await _operationRunner.RunAsync(
                    this,
                    "测试连接",
                    true,
                    (progress, cancellationToken) => _executionService.TestConnectionAsync(
                        connection,
                        Guid.NewGuid().ToString("N"),
                        progress,
                        cancellationToken));

                if (!result.Succeeded)
                {
                    _status.Text = result.Message + "（" + result.Elapsed.TotalSeconds.ToString("0.0") + " 秒）";
                    return;
                }

                _activeConnection = connection;
                _activeDatabase = null;
                RememberActiveConnection(connection);
                ReloadConnectionsTree(connection.Id);
                TreeNode node = _connectionsTree.Nodes.Cast<TreeNode>().FirstOrDefault(item =>
                    string.Equals((item.Tag as ConnectionProfileSnapshot)?.Id, connection.Id, StringComparison.OrdinalIgnoreCase));
                if (node != null)
                {
                    _connectionsTree.SelectedNode = node;
                    EnsureDatabasePlaceholder(node);
                    if (expandNode) node.Expand();
                }

                _status.Text = "当前连接已激活：" + connection.Name;
                ApplyTreeVisualState();
            }
            catch (OperationCanceledException)
            {
                _status.Text = "连接测试已取消。";
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private void ResetActiveConnectionVisualState()
        {
            _activeConnection = null;
            _activeDatabase = null;
            foreach (TreeNode node in _connectionsTree.Nodes)
            {
                ConnectionProfileSnapshot connection = node.Tag as ConnectionProfileSnapshot;
                if (connection == null) continue;
                node.Collapse();
                node.ImageKey = DatabaseTreeImageCatalog.Provider(connection.ProviderId, false);
                node.SelectedImageKey = node.ImageKey;
            }
            ApplyTreeVisualState();
        }

        private void ConnectionsTreeAfterSelect(object sender, TreeViewEventArgs e)
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
                    _status.Text = "当前数据库：" + _activeDatabase;
                    return;
                }

                ConnectionProfileSnapshot connection = e.Node?.Tag as ConnectionProfileSnapshot;
                if (connection != null && string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase))
                {
                    _activeDatabase = null;
                    ApplyTreeVisualState();
                    _status.Text = "当前连接已激活：" + connection.Name;
                }
            }
            finally
            {
                UpdateConnectionMenuState();
            }
        }

        private async void ConnectionsTreeBeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            ConnectionProfileSnapshot connection = e.Node?.Tag as ConnectionProfileSnapshot;
            if (connection != null)
            {
                if (!string.Equals(_activeConnection?.Id, connection.Id, StringComparison.OrdinalIgnoreCase))
                {
                    e.Cancel = true;
                    await OnUiAsync(() => _status.Text = "请先将该连接设置为当前连接。");
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
            IDatabaseQueryProvider provider = _providers.GetById(connection.ProviderId) as IDatabaseQueryProvider;
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
                    _status.Text = "数据库列表已更新：" + connection.Name;
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

            IDatabaseQueryProvider provider = _providers.GetById(connection.ProviderId) as IDatabaseQueryProvider;
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
                            TreeNode objectNode = new TreeNode(item.Name)
                            {
                                Tag = item,
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
            if (connectionNode == null) return;
            if (connectionNode.Nodes.Count > 0) return;
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

        private ConnectionProfileSnapshot ResolveExecutionConnection()
        {
            if (_activeConnection == null) return null;
            if (string.IsNullOrWhiteSpace(_activeDatabase)) return _activeConnection;
            IDatabaseQueryProvider provider = _providers.GetById(_activeConnection.ProviderId) as IDatabaseQueryProvider;
            if (provider == null) return _activeConnection;
            try
            {
                return provider.CreateDatabaseSnapshot(_activeConnection, _activeDatabase);
            }
            catch
            {
                return _activeConnection;
            }
        }

        private void ApplyTreeVisualState()
        {
            if (InvokeRequired)
            {
                OnUi(ApplyTreeVisualState);
                return;
            }

            foreach (TreeNode root in _connectionsTree.Nodes)
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

        private void ShowError(Exception exception)
        {
            _status.Text = exception.Message;
            MessageBox.Show(this, exception.Message, AppPresentation.DisplayName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void UpdateConnectionMenuState()
        {
            bool hasSelection = SelectedConnectionNode != null;
            _editConnectionMenuItem.Enabled = hasSelection;
            _deleteConnectionMenuItem.Enabled = hasSelection;
            _setCurrentConnectionMenuItem.Enabled = hasSelection;
        }

        private static Button ButtonOf(string text) => new Button { Text = text, AutoSize = true };

        private static Panel SectionPanel(string title, Control content)
        {
            Panel panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 30, 0, 0) };
            panel.Controls.Add(content);
            panel.Controls.Add(SectionHeader(title));
            return panel;
        }

        private static Panel SectionHeader(string title)
        {
            Panel header = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = SystemColors.ControlLight };
            header.Controls.Add(new Label
            {
                Text = title,
                AutoSize = true,
                Font = new Font(AppPresentation.DefaultFontName, AppPresentation.DefaultFontSize, FontStyle.Bold),
                Location = new Point(8, 7)
            });
            return header;
        }

        /// <summary>为查询方案下拉框包装模型和显示名称。</summary>
        private sealed class QueryItem
        {
            public QueryItem(QueryProfile profile) { Profile = profile; }
            public QueryProfile Profile { get; }
            public string DisplayName => Profile.Name;
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
}

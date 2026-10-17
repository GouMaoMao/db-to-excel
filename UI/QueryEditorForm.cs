using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
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
    /// 每个页签是一份查询：SQL、方案名称和上次预览结果都留在该页。预览表格只有一份，切换页签时重画当前页的结果。
    /// 用户输入包括连接方案、SQL 和方案名称。写入时另问目标工作表，并可把 SQL 写入 SQL 页。
    /// 输出副作用包括保存查询方案和修改当前工作簿。页签在界面线程上增减，不访问 Excel。
    /// </remarks>
    public sealed class QueryEditorForm : AppForm
    {
        private readonly IQueryProfileRepository _queries;
        private readonly IConnectionProfileRepository _connections;
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IQueryBufferService _queryBuffer;
        private readonly IExcelResultWriter _writer;
        private readonly ISqlSheetTaskReader _taskReader;
        private readonly IOperationRunner _operationRunner;
        private readonly ISettingsStore _settings;
        private readonly IJdbcEnvironmentStore _jdbcEnvironment;
        private readonly Action _openJdbcEnvironment;
        private readonly ExcelInterop.Workbook _workbook;
        private readonly DatabaseConnectionTree _connectionTree;
        private readonly ListBox _queryProfiles;
        private readonly ISqlFormattingService _formatting;
        private readonly TabControl _queryTabs;
        private readonly Label _emptyEditorHint;
        private readonly DataGridView _preview;
        private readonly Label _sqlQueryNameLabel;
        private readonly ToolTip _profileTip = new ToolTip();
        private readonly Timer _profileTipTimer = new Timer();
        private int _profileTipIndex = -1;
        private string _profileTipText = string.Empty;
        private readonly Label _sqlLimitHint;
        private readonly Font _sqlLimitHintFont;
        private readonly ToolStripStatusLabel _status;
        private readonly SplitContainer _workspace;
        private readonly SplitContainer _leftSplit;
        private readonly SplitContainer _resultsSplit;
        private readonly List<Image> _actionIcons = new List<Image>();
        private readonly Image _queryScriptIcon;
        private readonly ContextMenuStrip _queryProfileMenu;
        private readonly ContextMenuStrip _queryTabMenu;
        private readonly ToolStripMenuItem _tabSaveMenuItem;
        private readonly ToolStripMenuItem _tabRenameMenuItem;
        private readonly Font _previewHeaderTypeFont;
        private readonly ToolStripMenuItem _editConnectionMenuItem;
        private readonly ToolStripMenuItem _deleteConnectionMenuItem;
        private readonly ToolStripMenuItem _setCurrentConnectionMenuItem;
        private bool _bindingProfiles;
        private bool _adjustingTabs;
        private bool _closePromptActive;
        private bool _applyingWorkspaceSplit;
        private bool _userDraggingWorkspaceSplit;

        /// <summary>创建 SQL 查询编辑窗体。</summary>
        /// <param name="queries">查询方案仓储。</param>
        /// <param name="connections">连接方案仓储。</param>
        /// <param name="providers">数据源提供程序注册表。</param>
        /// <param name="executionService">查询执行和连接测试服务。</param>
        /// <param name="queryBuffer">用于预览和写入前缓冲查询结果。</param>
        /// <param name="writer">Excel 结果写入器。</param>
        /// <param name="taskReader">SQL 页读写器。勾选写入 SQL 时用来追加或更新任务列。</param>
        /// <param name="operationRunner">进度窗体运行器。</param>
        /// <param name="settings">查询行数、超时和分块大小设置。</param>
        /// <param name="workbook">当前操作的 Excel 工作簿。</param>
        /// <param name="jdbcEnvironment">本机 Java 路径，传给连接树里的编辑窗。</param>
        /// <param name="openJdbcEnvironment">打开 JDBC 环境窗体；为空时连接编辑不显示该入口。</param>
        public QueryEditorForm(
            IQueryProfileRepository queries,
            IConnectionProfileRepository connections,
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IQueryBufferService queryBuffer,
            IExcelResultWriter writer,
            ISqlSheetTaskReader taskReader,
            IOperationRunner operationRunner,
            ISettingsStore settings,
            ExcelInterop.Workbook workbook,
            IJdbcEnvironmentStore jdbcEnvironment,
            Action openJdbcEnvironment = null)
        {
            _queries = queries ?? throw new ArgumentNullException(nameof(queries));
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _queryBuffer = queryBuffer ?? throw new ArgumentNullException(nameof(queryBuffer));
            _writer = writer ?? throw new ArgumentNullException(nameof(writer));
            _taskReader = taskReader ?? throw new ArgumentNullException(nameof(taskReader));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _jdbcEnvironment = jdbcEnvironment ?? throw new ArgumentNullException(nameof(jdbcEnvironment));
            _openJdbcEnvironment = openJdbcEnvironment;

            Text = AppPresentation.WindowTitle("SQL 查询");
            Rectangle workingArea = FormSizeMemory.CursorWorkingArea();
            Width = Math.Min(1320, workingArea.Width - 40);
            Height = Math.Min(860, workingArea.Height - 60);
            MinimumSize = new Size(820, 600);
            StartPosition = FormStartPosition.Manual;
            FormSizeMemory.Attach(this, _settings, "QueryEditor");

            _connectionTree = new DatabaseConnectionTree(
                _connections, _providers, _executionService, _operationRunner, _settings, _jdbcEnvironment, _openJdbcEnvironment);
            _queryScriptIcon = DatabaseTreeImageCatalog.LoadActionIcon("icon_script.png");
            int queryLineHeight = TextRenderer.MeasureText("方案", Font).Height;
            _queryProfiles = new ListBox
            {
                Dock = DockStyle.Fill,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = Math.Max(20, queryLineHeight + 4),
                DisplayMember = nameof(QueryItem.DisplayName),
                IntegralHeight = false
            };
            _formatting = new SqlFormattingService();
            _queryTabs = new TabControl { Dock = DockStyle.Fill };
            _emptyEditorHint = new Label
            {
                Dock = DockStyle.Fill,
                Text = "请新建查询",
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = SystemColors.GrayText,
                BackColor = SystemColors.Window,
                Visible = false
            };
            _queryTabMenu = BuildQueryTabMenu(out _tabSaveMenuItem, out _tabRenameMenuItem);
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
            _sqlQueryNameLabel = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(0, 0, 8, 0),
                Text = "当前方案：未保存"
            };
            _status = new ToolStripStatusLabel { Text = "就绪。", ForeColor = SystemColors.GrayText, Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            _queryProfileMenu = BuildQueryProfileMenu();

            _setCurrentConnectionMenuItem = new ToolStripMenuItem("设置为当前连接", null, (sender, args) => _connectionTree.SetSelectedAsCurrent());
            _editConnectionMenuItem = new ToolStripMenuItem("编辑连接", null, (sender, args) => _connectionTree.EditSelectedConnection());
            _deleteConnectionMenuItem = new ToolStripMenuItem("删除连接", null, (sender, args) => _connectionTree.DeleteSelectedConnection());
            ToolStripMenuItem newConnectionMenuItem = new ToolStripMenuItem("新建连接", null, (sender, args) => _connectionTree.CreateConnection());
            ToolStripMenuItem refreshConnectionMenuItem = new ToolStripMenuItem("刷新", null, (sender, args) => _connectionTree.RefreshSelected());
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
            ToolStripMenuItem renameMenuItem = new ToolStripMenuItem("重命名", null, (sender, args) => RenameQuery());
            ToolStripMenuItem deleteMenuItem = new ToolStripMenuItem("删除方案", null, (sender, args) => DeleteQuery());
            ToolStripMenuItem previewMenuItem = new ToolStripMenuItem("预览", null, (sender, args) => PreviewQuery()) { ShortcutKeys = Keys.Control | Keys.Enter };
            ToolStripMenuItem exportMenuItem = new ToolStripMenuItem("写入 Sheet", null, (sender, args) => ExportQuery()) { ShortcutKeys = Keys.Control | Keys.Shift | Keys.E };
            ToolStripMenuItem closeMenuItem = new ToolStripMenuItem("关闭", null, (sender, args) => Close()) { ShortcutKeys = Keys.Alt | Keys.F4 };
            ToolStripMenuItem queryMenu = new ToolStripMenuItem("查询");
            queryMenu.DropDownItems.AddRange(new ToolStripItem[]
            {
                newMenuItem,
                saveMenuItem,
                renameMenuItem,
                deleteMenuItem,
                new ToolStripSeparator(),
                previewMenuItem,
                exportMenuItem,
                new ToolStripSeparator(),
                closeMenuItem
            });

            ToolStripMenuItem formatMenuItem = new ToolStripMenuItem("格式化 SQL", null, (sender, args) => FormatActiveSql()) { ShortcutKeys = Keys.Control | Keys.Shift | Keys.F };
            ToolStripMenuItem commentMenuItem = new ToolStripMenuItem("注释/取消注释", null, (sender, args) => ToggleActiveComment()) { ShortcutKeys = Keys.Control | Keys.OemQuestion };
            ToolStripMenuItem editMenu = new ToolStripMenuItem("编辑");
            editMenu.DropDownItems.AddRange(new ToolStripItem[] { formatMenuItem, commentMenuItem });

            MenuStrip menu = new MenuStrip { Dock = DockStyle.Top };
            menu.Items.Add(connectionMenu);
            menu.Items.Add(queryMenu);
            menu.Items.Add(editMenu);
            MainMenuStrip = menu;

            Button previewButton = ButtonOf("预览", "icon_execute.png");
            Button exportButton = ButtonOf("写入 Sheet", "icon_writeexcel.png");
            Button saveButton = ButtonOf("保存方案", "icon_save.png");
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
            _sqlLimitHintFont = new Font(AppPresentation.DefaultFontName, 8.25f, FontStyle.Regular);
            _sqlLimitHint = new Label
            {
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 48,
                ForeColor = SystemColors.GrayText,
                Font = _sqlLimitHintFont,
                Padding = new Padding(8, 6, 8, 0),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Panel editorHost = new Panel { Dock = DockStyle.Fill };
            editorHost.Controls.Add(_queryTabs);
            editorHost.Controls.Add(_emptyEditorHint);
            Panel sqlPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 30, 0, 0) };
            sqlPanel.Controls.Add(editorHost);
            sqlPanel.Controls.Add(_sqlLimitHint);
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

            _leftSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal
            };
            _leftSplit.Panel1.Controls.Add(SectionPanel("数据库连接", _connectionTree));
            _leftSplit.Panel2.Controls.Add(SectionPanel("查询方案", _queryProfiles));

            _workspace = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6
            };
            _workspace.Panel1.Controls.Add(_leftSplit);
            _workspace.Panel2.Controls.Add(_resultsSplit);

            StatusStrip statusStrip = new StatusStrip { Dock = DockStyle.Bottom, SizingGrip = false };
            statusStrip.Items.Add(_status);

            Controls.Add(_workspace);
            Controls.Add(statusStrip);
            Controls.Add(menu);

            previewButton.Click += (sender, args) => PreviewQuery();
            exportButton.Click += (sender, args) => ExportQuery();
            saveButton.Click += (sender, args) => SaveQuery();
            _queryTabs.MouseDown += QueryTabsMouseDown;
            _queryTabs.MouseUp += QueryTabsMouseUp;
            _queryTabs.SelectedIndexChanged += QueryTabsSelectedIndexChanged;
            _emptyEditorHint.MouseUp += EmptyEditorMouseUp;
            _queryProfiles.SelectedIndexChanged += QueryProfilesSelectedIndexChanged;
            _queryProfiles.MouseUp += QueryProfilesMouseUp;
            _queryProfiles.MouseMove += QueryProfilesMouseMove;
            _queryProfiles.MouseLeave += QueryProfilesMouseLeave;
            _queryProfiles.DrawItem += QueryProfilesDrawItem;
            _profileTipTimer.Interval = 500;
            _profileTipTimer.Tick += ProfileTipTimerTick;
            _preview.RowPostPaint += PreviewRowPostPaint;
            _preview.CellPainting += PreviewCellPainting;
            _connectionTree.NoticeChanged += (sender, args) => _status.Text = _connectionTree.Notice;
            _connectionTree.SelectionChanged += (sender, args) => UpdateConnectionMenuState();
            FormClosing += QueryEditorFormFormClosing;
            FormClosed += QueryEditorFormFormClosed;
            Activated += (sender, args) => UpdateSqlLimitHint();
            UpdateSqlLimitHint();
            _workspace.SizeChanged += (sender, args) => ApplyWorkspaceSplit();
            _workspace.SplitterMoving += (sender, args) => _userDraggingWorkspaceSplit = true;
            _workspace.SplitterMoved += (sender, args) => RememberWorkspaceSplit(_workspace);

            ReloadQueries(null);
            NewQuery();
            UpdateQueryInfo();
            UpdateConnectionMenuState();
        }

        /// <summary>按仓储重绑左侧方案列表，并选中指定方案。选中不会把 SQL 再灌进编辑器。</summary>
        /// <param name="selectedId">要高亮的方案标识。为空时取消列表选中。</param>
        private void ReloadQueries(string selectedId)
        {
            List<QueryItem> items = _queries.GetAll().Select(profile => new QueryItem(profile)).ToList();
            _bindingProfiles = true;
            try
            {
                _queryProfiles.DataSource = null;
                _queryProfiles.DataSource = items;
                _queryProfiles.DisplayMember = nameof(QueryItem.DisplayName);
                if (string.IsNullOrWhiteSpace(selectedId))
                {
                    _queryProfiles.SelectedIndex = -1;
                    return;
                }

                for (int index = 0; index < items.Count; index++)
                {
                    if (!string.Equals(items[index].Profile.Id, selectedId, StringComparison.OrdinalIgnoreCase))
                        continue;
                    _queryProfiles.SelectedIndex = index;
                    return;
                }

                _queryProfiles.SelectedIndex = -1;
            }
            finally
            {
                _bindingProfiles = false;
            }
        }

        /// <summary>新开一个空白页签，标题用尚未占用的「新建查询」序号。</summary>
        private void NewQuery()
        {
            AddTab(NextNewQueryName(), string.Empty, Guid.NewGuid().ToString("N"), DateTime.UtcNow, null, false, null);
            _status.Text = "新查询方案。";
        }

        /// <summary>左侧列表点选时打开或激活对应页签，不改写已经打开的 SQL。</summary>
        private void QueryProfilesSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_bindingProfiles) return;
            QueryProfile profile = (_queryProfiles.SelectedItem as QueryItem)?.Profile;
            if (profile == null) return;
            OpenProfile(profile);
        }

        /// <summary>把方案打开到页签。同一标识已经打开时只切换过去。</summary>
        /// <param name="profile">要打开的已保存方案。</param>
        private void OpenProfile(QueryProfile profile)
        {
            QueryTab existing = FindTab(profile.Id);
            if (existing != null)
            {
                if (ActiveTab != existing)
                {
                    SelectTab(existing);
                    _status.Text = "已打开查询方案。";
                }
                return;
            }

            AddTab(profile.Name, profile.QueryText, profile.Id, profile.CreatedUtc, profile.UpdatedUtc, true, profile.CreatedBy);
            _status.Text = "已打开查询方案。";
        }

        /// <summary>保存当前页签。没有页签、取消或校验失败时返回 false，页签保持打开。</summary>
        private void SaveQuery()
        {
            SaveTab(ActiveTab);
        }

        /// <summary>把指定页签的 SQL 写入查询方案。覆盖其他方案时关掉占用该标识的页签，避免两页共用一份方案。</summary>
        /// <param name="tab">要保存的页签。为空时提示新建查询。</param>
        /// <returns>已经写入仓储时为 true。取消、校验失败或没有页签时为 false。</returns>
        private bool SaveTab(QueryTab tab)
        {
            if (tab == null)
            {
                PromptNewQuery();
                return false;
            }

            if (ActiveTab != tab)
                SelectTab(tab);
            if (!ValidateInputs(out ConnectionProfileSnapshot connection)) return false;
            if (!TextPromptDialog.TryShow(
                this,
                "保存查询方案",
                "请输入方案名称",
                string.IsNullOrWhiteSpace(tab.Name) ? "新查询方案" : tab.Name,
                out string queryName))
            {
                _status.Text = "已取消保存方案。";
                return false;
            }

            QueryProfile duplicate = _queries.GetAll().FirstOrDefault(item =>
                string.Equals(item.Name, queryName, StringComparison.CurrentCultureIgnoreCase) &&
                !string.Equals(item.Id, tab.Id, StringComparison.OrdinalIgnoreCase));
            bool overwritten = false;
            if (duplicate != null)
            {
                DialogResult overwrite = MessageBox.Show(
                    this,
                    string.Format("查询方案“{0}”已存在，是否覆盖？", queryName),
                    "保存查询方案",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
                if (overwrite != DialogResult.Yes)
                {
                    _status.Text = "已取消覆盖保存。";
                    return false;
                }

                CloseTabs(duplicate.Id, false, tab);
                tab.Id = duplicate.Id;
                tab.CreatedUtc = duplicate.CreatedUtc;
                tab.CreatedBy = duplicate.CreatedBy;
                overwritten = true;
            }

            string person = CurrentPerson();
            string createdBy = overwritten || tab.IsPersisted ? tab.CreatedBy : person;
            DateTime createdUtc = tab.CreatedUtc == default(DateTime) ? DateTime.UtcNow : tab.CreatedUtc;
            DateTime updatedUtc = DateTime.UtcNow;
            string connectionName = string.IsNullOrWhiteSpace(connection.Name) ? "未选择" : connection.Name;
            tab.Editor.SqlText = QuerySchemeNote.Apply(
                tab.Editor.SqlText,
                queryName,
                connectionName,
                createdBy,
                createdUtc,
                person,
                updatedUtc);
            string sql = tab.Editor.SqlText;
            QueryProfile profile = new QueryProfile
            {
                Id = string.IsNullOrWhiteSpace(tab.Id) ? Guid.NewGuid().ToString("N") : tab.Id,
                Name = queryName,
                ProviderId = connection.ProviderId,
                ConnectionProfileId = connection.Id,
                QueryText = sql,
                CreatedUtc = createdUtc,
                UpdatedUtc = updatedUtc,
                CreatedBy = createdBy,
                UpdatedBy = person
            };
            _queries.Save(profile);
            QueryProfile stored = _queries.GetById(profile.Id);
            tab.Id = profile.Id;
            tab.Name = queryName;
            tab.CreatedUtc = profile.CreatedUtc;
            tab.CreatedBy = createdBy;
            tab.SavedUtc = stored != null ? stored.UpdatedUtc : profile.UpdatedUtc;
            tab.SavedSql = sql;
            tab.IsPersisted = true;
            RefreshTabTitle(tab);
            ReloadQueries(profile.Id);
            _status.Text = overwritten ? "查询方案已覆盖保存。" : "查询方案已保存。";
            UpdateQueryInfo();
            return true;
        }

        /// <summary>重命名当前页签。已保存的方案只改名称，不把编辑器里尚未保存的 SQL 写回磁盘。</summary>
        private void RenameQuery()
        {
            RenameTab(ActiveTab);
        }

        /// <summary>修改页签标题。已保存时同步仓储名称；与其他方案重名时沿用覆盖确认。</summary>
        /// <param name="tab">要重命名的页签。为空时提示新建查询。</param>
        /// <remarks>
        /// 未保存的页签确认覆盖后只收下这个名字，不删除已有方案，真正写入留到保存。
        /// 已保存的页签确认覆盖后删除同名方案并关掉它的页签，当前页仍保留编辑器中的未保存 SQL。
        /// </remarks>
        private void RenameTab(QueryTab tab)
        {
            if (tab == null)
            {
                PromptNewQuery();
                return;
            }

            if (ActiveTab != tab)
                SelectTab(tab);
            if (!TextPromptDialog.TryShow(this, "重命名查询方案", "请输入方案名称", tab.Name, out string queryName))
            {
                _status.Text = "已取消重命名。";
                return;
            }

            if (string.Equals(queryName, tab.Name, StringComparison.CurrentCulture))
                return;

            QueryProfile duplicate = _queries.GetAll().FirstOrDefault(item =>
                string.Equals(item.Name, queryName, StringComparison.CurrentCultureIgnoreCase) &&
                !string.Equals(item.Id, tab.Id, StringComparison.OrdinalIgnoreCase));
            if (duplicate != null)
            {
                DialogResult overwrite = MessageBox.Show(
                    this,
                    string.Format("查询方案“{0}”已存在，是否覆盖？", queryName),
                    "保存查询方案",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
                if (overwrite != DialogResult.Yes)
                {
                    _status.Text = "已取消重命名。";
                    return;
                }

                if (tab.IsPersisted)
                {
                    CloseTabs(duplicate.Id, false, tab);
                    _queries.Delete(duplicate.Id);
                }
            }

            tab.Name = queryName;
            if (!tab.IsPersisted)
            {
                RefreshTabTitle(tab);
                UpdateQueryInfo();
                _status.Text = "查询方案已重命名。";
                return;
            }

            QueryProfile stored = _queries.GetById(tab.Id);
            if (stored == null)
            {
                tab.IsPersisted = false;
                tab.SavedUtc = null;
                RefreshTabTitle(tab);
                ReloadQueries(null);
                UpdateQueryInfo();
                _status.Text = "查询方案已重命名。";
                return;
            }

            stored.Name = queryName;
            _queries.Save(stored);
            QueryProfile refreshed = _queries.GetById(tab.Id);
            tab.SavedUtc = refreshed != null ? refreshed.UpdatedUtc : tab.SavedUtc;
            RefreshTabTitle(tab);
            ReloadQueries(tab.Id);
            UpdateQueryInfo();
            _status.Text = "查询方案已重命名。";
        }

        /// <summary>删除左侧选中的方案，并关掉打开着的对应页签。不另开空白页。</summary>
        private void DeleteQuery()
        {
            QueryProfile profile = (_queryProfiles.SelectedItem as QueryItem)?.Profile;
            if (profile == null) return;
            if (MessageBox.Show(this, "确定删除查询方案“" + profile.Name + "”吗？", "删除查询方案",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            _queries.Delete(profile.Id);
            CloseTabs(profile.Id, false, null);
            QueryTab active = ActiveTab;
            ReloadQueries(active != null && active.IsPersisted ? active.Id : null);
            _status.Text = "查询方案已删除。";
        }

        /// <summary>按当前页签刷新标题栏上的当前方案。没有页签时显示未保存。</summary>
        private void UpdateQueryInfo()
        {
            QueryTab tab = ActiveTab;
            string name = tab == null || string.IsNullOrWhiteSpace(tab.Name) ? "未保存" : tab.Name;
            _sqlQueryNameLabel.Text = "当前方案：" + name;
        }

        /// <summary>读取当前 Windows 用户名，以及 Excel 或 WPS 选项里的用户名。</summary>
        /// <returns>拼好的修改人。表格用户名读不出来时只保留 Windows 用户名，不中断保存。</returns>
        private string CurrentPerson()
        {
            string spreadsheet = string.Empty;
            try
            {
                if (_workbook != null && _workbook.Application != null)
                    spreadsheet = _workbook.Application.UserName;
            }
            catch (Exception)
            {
                // 表格进程读不出用户名时仍继续保存，修改人只保留 Windows 用户名。
            }

            return QuerySchemeNote.FormatPerson(Environment.UserName, spreadsheet);
        }

        /// <summary>把连接标识换成悬停里要显示的名称。</summary>
        /// <param name="connectionId">方案记录的连接标识。</param>
        /// <returns>连接名称。没有标识时为「未选择」，标识已不存在时为「连接已删除」。</returns>
        private string ConnectionDisplayName(string connectionId)
        {
            if (string.IsNullOrWhiteSpace(connectionId)) return "未选择";
            ConnectionProfileSnapshot connection = _connections.GetById(connectionId);
            if (connection == null || string.IsNullOrWhiteSpace(connection.Name)) return "连接已删除";
            return connection.Name;
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
            QueryTab tab = ActiveTab;
            string sqlText = tab == null ? string.Empty : (selectedOnly ? tab.Editor.SelectedSql : tab.Editor.SqlText);
            if (!ValidateInputs(out ConnectionProfileSnapshot connection, sqlText)) return;
            bool hadPreview = tab.HasPreview;
            try
            {
                DataSourceRequest request = CreateRequest(connection, sqlText, ExecutionPurpose.Preview, _settings.Get(CoreSettings.MaxPreviewRows));
                BufferedQueryResult result = _operationRunner.Run(
                    this, "查询预览", true,
                    (progress, token) => _queryBuffer.ExecuteAsync(request, Guid.NewGuid().ToString("N"), progress, token));
                tab.Preview = result;
                tab.HasPreview = true;
                if (ActiveTab == tab)
                {
                    BindPreview(result);
                    EnsurePreviewVisible();
                }
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
                    tab.HasPreview = false;
                    tab.Preview = null;
                    if (ActiveTab == tab)
                        _resultsSplit.Panel2Collapsed = true;
                }
                _status.Text = selectedOnly ? "选中文本预览已取消。" : "查询预览已取消。";
            }
            catch (Exception exception)
            {
                if (!hadPreview)
                {
                    tab.HasPreview = false;
                    tab.Preview = null;
                    if (ActiveTab == tab)
                        _resultsSplit.Panel2Collapsed = true;
                }
                ShowError(exception);
            }
        }

        /// <summary>把当前 SQL 的结果写入用户确认的工作表。表已存在时先确认覆盖。</summary>
        /// <remarks>
        /// 勾选写入 SQL 时，结果写成功后才把 SQL 按列写入 SQL 页。取消或覆盖被拒绝时两者都不写。
        /// 写入成功且「保存后美化」为开时，再走一遍美化进度窗。这次美化取消或失败不撤销已经写入的结果和 SQL。
        /// </remarks>
        private void ExportQuery()
        {
            QueryTab tab = ActiveTab;
            if (!ValidateInputs(out ConnectionProfileSnapshot connection)) return;
            if (!ExportSheetDialog.TryShow(this, tab.Name, out string targetSheet, out bool writeSql))
            {
                _status.Text = "已取消写入。";
                return;
            }

            try
            {
                if (WorksheetExists(_workbook, targetSheet)
                    && MessageBox.Show(
                        this,
                        string.Format("工作表“{0}”已存在，是否覆盖？", targetSheet),
                        "写入 Sheet",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                {
                    _status.Text = "已取消写入。";
                    return;
                }

                string sqlText = tab.Editor.SqlText;
                DataSourceRequest request = CreateRequest(connection, sqlText, ExecutionPurpose.Export, _settings.Get(CoreSettings.MaxExportRows));
                long rows = _operationRunner.Run(this, "写入 Excel", true, async (progress, token) =>
                {
                    string operationId = Guid.NewGuid().ToString("N");
                    BufferedQueryResult result = await _queryBuffer.ExecuteAsync(
                        request, operationId, progress, token);
                    long written = _writer.WriteResult(_workbook, targetSheet, result, progress, token, operationId);
                    if (writeSql)
                        _taskReader.WriteTaskColumn(_workbook, targetSheet, sqlText);
                    return written;
                });
                _status.Text = writeSql
                    ? string.Format("已向“{0}”写入 {1:N0} 行，并已把 SQL 写入工作表 SQL。", targetSheet, rows)
                    : string.Format("已向“{0}”写入 {1:N0} 行。", targetSheet, rows);
                if (writeSql)
                    BeautifySavedSql(targetSheet, rows);
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

        /// <summary>按「保存后美化」在写入成功后再美化整张 SQL 页。</summary>
        /// <param name="targetSheet">刚写入的结果表名，只用于完成后的状态文字。</param>
        /// <param name="rows">已写入的结果行数。</param>
        /// <remarks>开关关闭、没有可美化的列、取消或失败时不把这次写入当成失败。取消和失败改写状态栏。</remarks>
        private void BeautifySavedSql(string targetSheet, long rows)
        {
            if (!_settings.Get(CoreSettings.BeautifyAfterSave)) return;
            try
            {
                int count = _taskReader.CountStyleColumns(_workbook);
                if (count == 0) return;
                _operationRunner.Run(
                    this,
                    "美化格式",
                    true,
                    (progress, token) => _taskReader.FormatSqlSheet(_workbook, progress, token));
                _status.Text = string.Format("已向“{0}”写入 {1:N0} 行，并已美化 SQL 页。", targetSheet, rows);
            }
            catch (OperationCanceledException)
            {
                _status.Text = "SQL 已写入，美化已取消。";
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        /// <summary>判断当前工作簿是否已有同名工作表。比较时忽略大小写，不激活工作表。</summary>
        /// <param name="workbook">当前工作簿。</param>
        /// <param name="name">要查找的工作表名称。</param>
        /// <returns>存在时为 true。会释放枚举到的 COM 引用。</returns>
        private static bool WorksheetExists(ExcelInterop.Workbook workbook, string name)
        {
            foreach (object item in workbook.Worksheets)
            {
                ExcelInterop.Worksheet worksheet = item as ExcelInterop.Worksheet;
                if (worksheet == null) continue;
                bool match = string.Equals(worksheet.Name, name, StringComparison.OrdinalIgnoreCase);
                Marshal.FinalReleaseComObject(worksheet);
                if (match) return true;
            }

            return false;
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

        /// <summary>
        /// 窗口仍隐藏时按最终客户区摆左右分栏，并把左栏上下分成两半。
        /// 显示之后再改分割条会把旧位置留在屏幕上。
        /// </summary>
        protected override void PrepareFirstShow()
        {
            ApplyWorkspaceSplit();
            TrySetSplitterDistance(_leftSplit, _leftSplit.Height / 2);
        }

        /// <summary>按已记住的千分比摆左右分栏。容器还没有宽度时不改距离。</summary>
        private void ApplyWorkspaceSplit()
        {
            _userDraggingWorkspaceSplit = false;
            _applyingWorkspaceSplit = true;
            try
            {
                ApplySplitterPanelMinSizes(_workspace, 140, 420);
                TrySetSplitterDistance(_workspace, WorkspaceSplitDistance(_workspace));
            }
            finally
            {
                _applyingWorkspaceSplit = false;
            }
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

        private int WorkspaceSplitDistance(SplitContainer workspace)
        {
            int available = workspace.Width - workspace.SplitterWidth;
            if (available <= 0) return 0;
            int ratio = _settings.Get(CoreSettings.WorkspaceSplitRatio);
            return ratio > 0 ? available * ratio / 1000 : available / 6;
        }

        private void RememberWorkspaceSplit(SplitContainer workspace)
        {
            if (!_userDraggingWorkspaceSplit) return;
            _userDraggingWorkspaceSplit = false;
            if (_applyingWorkspaceSplit || workspace == null || workspace.IsDisposed) return;
            int available = workspace.Width - workspace.SplitterWidth;
            if (available <= 0) return;
            int ratio = (int)Math.Round(workspace.SplitterDistance * 1000d / available);
            ratio = Math.Max(1, Math.Min(999, ratio));
            if (ratio == _settings.Get(CoreSettings.WorkspaceSplitRatio)) return;
            _settings.Set(CoreSettings.WorkspaceSplitRatio, ratio);
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


        /// <summary>检查已有页签、已选择当前连接并且 SQL 非空。失败时弹出说明并返回 false。</summary>
        /// <param name="connection">当前要执行的连接。检查失败时为 null。</param>
        /// <param name="sqlText">要检查的 SQL。为空时使用当前页签全文。</param>
        /// <returns>可以继续预览、保存或写入时为 true。</returns>
        /// <remarks>没有页签时提示新建查询。没有连接方案时指向连接管理新建；有方案但未设当前连接时指向左侧列表右键。</remarks>
        private bool ValidateInputs(out ConnectionProfileSnapshot connection, string sqlText = null)
        {
            connection = null;
            QueryTab tab = ActiveTab;
            if (tab == null)
            {
                PromptNewQuery();
                return false;
            }

            connection = _connectionTree.ExecutionConnection();
            string sql = sqlText ?? tab.Editor.SqlText;
            string message = null;
            if (connection == null)
            {
                message = _connectionTree.HasConnectionProfiles
                    ? AppPresentation.SetCurrentConnectionPrompt
                    : AppPresentation.NoConnectionPrompt;
            }
            else if (QuerySchemeNote.LacksStatement(sql)) message = "请输入 SQL。";
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

        private ContextMenuStrip BuildQueryProfileMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("新建", null, (sender, args) => NewQuery());
            menu.Items.Add("保存", null, (sender, args) => SaveQuery());
            menu.Items.Add("删除", null, (sender, args) => DeleteQuery());
            return menu;
        }

        private void QueryEditorFormFormClosed(object sender, FormClosedEventArgs e)
        {
            FormClosing -= QueryEditorFormFormClosing;
            _queryTabs.MouseDown -= QueryTabsMouseDown;
            _queryTabs.MouseUp -= QueryTabsMouseUp;
            _queryTabs.SelectedIndexChanged -= QueryTabsSelectedIndexChanged;
            _emptyEditorHint.MouseUp -= EmptyEditorMouseUp;
            _queryProfiles.MouseUp -= QueryProfilesMouseUp;
            _queryProfiles.MouseMove -= QueryProfilesMouseMove;
            _queryProfiles.MouseLeave -= QueryProfilesMouseLeave;
            _queryProfiles.DrawItem -= QueryProfilesDrawItem;
            _profileTipTimer.Tick -= ProfileTipTimerTick;
            _profileTipTimer.Stop();
            _profileTipTimer.Dispose();
            _profileTip.Dispose();
            _preview.RowPostPaint -= PreviewRowPostPaint;
            _preview.CellPainting -= PreviewCellPainting;
            _queryProfileMenu.Dispose();
            _queryTabMenu.Dispose();
            foreach (Image icon in _actionIcons)
                icon.Dispose();
            _actionIcons.Clear();
            _queryScriptIcon.Dispose();
            _previewHeaderTypeFont.Dispose();
            _sqlLimitHintFont.Dispose();
        }

        /// <summary>在 SQL 标题和编辑框之间说明预览、导出和超时限制。激活时按当前设置刷新。</summary>
        private void UpdateSqlLimitHint()
        {
            _sqlLimitHint.Text = string.Format(
                "最大预览行数：{0}（可修改）。最大导出 Excel 行：{1}（可修改）。查询超时：{2} 秒（可修改）。",
                _settings.Get(CoreSettings.MaxPreviewRows),
                _settings.Get(CoreSettings.MaxExportRows),
                _settings.Get(CoreSettings.QueryTimeoutSeconds));
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
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding;
            TextRenderer.DrawText(e.Graphics, name, e.CellStyle.Font, nameBounds, e.CellStyle.ForeColor, flags);
            if (!string.IsNullOrWhiteSpace(typeName))
            {
                TextRenderer.DrawText(e.Graphics, typeName, _previewHeaderTypeFont, typeBounds, SystemColors.GrayText, flags);
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

        /// <summary>在查询方案名称左侧绘制脚本图标。列表仍按名称绑定，自绘只负责图标和省略过长名称。</summary>
        private void QueryProfilesDrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            e.DrawBackground();
            QueryItem item = _queryProfiles.Items[e.Index] as QueryItem;
            string name = item != null ? item.DisplayName : Convert.ToString(_queryProfiles.Items[e.Index]);
            const int iconSize = 16;
            const int gap = 4;
            int iconX = e.Bounds.X + gap;
            int iconY = e.Bounds.Y + Math.Max(0, (e.Bounds.Height - iconSize) / 2);
            e.Graphics.DrawImage(_queryScriptIcon, iconX, iconY, iconSize, iconSize);

            Rectangle textBounds = new Rectangle(
                iconX + iconSize + gap,
                e.Bounds.Y,
                Math.Max(0, e.Bounds.Right - iconX - iconSize - gap),
                e.Bounds.Height);
            TextRenderer.DrawText(
                e.Graphics,
                name ?? string.Empty,
                e.Font,
                textBounds,
                e.ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            e.DrawFocusRectangle();
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

        /// <summary>鼠标换到另一项时重新计时，停约半秒后显示该项的方案说明。</summary>
        private void QueryProfilesMouseMove(object sender, MouseEventArgs e)
        {
            int index = _queryProfiles.IndexFromPoint(e.Location);
            if (index == _profileTipIndex) return;

            HideProfileTip();
            _profileTipIndex = index;
            if (index < 0 || index >= _queryProfiles.Items.Count) return;
            QueryItem item = _queryProfiles.Items[index] as QueryItem;
            if (item == null || item.Profile == null) return;

            QueryProfile profile = item.Profile;
            _profileTipText = QuerySchemeNote.FormatLines(
                profile.Name,
                ConnectionDisplayName(profile.ConnectionProfileId),
                profile.CreatedBy,
                profile.CreatedUtc,
                profile.UpdatedBy,
                profile.UpdatedUtc);
            _profileTipTimer.Start();
        }

        /// <summary>离开方案列表时关掉悬停说明。</summary>
        private void QueryProfilesMouseLeave(object sender, EventArgs e)
        {
            _profileTipIndex = -1;
            _profileTipText = string.Empty;
            HideProfileTip();
        }

        /// <summary>在当前项下方显示已经准备好的方案说明。</summary>
        private void ProfileTipTimerTick(object sender, EventArgs e)
        {
            _profileTipTimer.Stop();
            if (_profileTipIndex < 0 || _profileTipIndex >= _queryProfiles.Items.Count) return;
            if (string.IsNullOrEmpty(_profileTipText)) return;
            Rectangle bounds = _queryProfiles.GetItemRectangle(_profileTipIndex);
            _profileTip.Show(_profileTipText, _queryProfiles, _queryProfiles.ClientSize.Width + 2, bounds.Top, 15000);
        }

        /// <summary>停掉悬停计时并隐藏说明。移到另一项或离开列表时调用。</summary>
        private void HideProfileTip()
        {
            _profileTipTimer.Stop();
            _profileTip.Hide(_queryProfiles);
        }

        private void ShowError(Exception exception)
        {
            _status.Text = exception.Message;
            ExceptionDetailForm.Show(this, exception);
        }

        private void UpdateConnectionMenuState()
        {
            bool hasSelection = _connectionTree.HasSelectedConnection;
            _editConnectionMenuItem.Enabled = hasSelection;
            _deleteConnectionMenuItem.Enabled = hasSelection;
            _setCurrentConnectionMenuItem.Enabled = hasSelection;
        }

        /// <summary>创建标题栏操作按钮，图标在文字左侧，高度留在 30 像素标题栏内。</summary>
        /// <param name="text">按钮文字。</param>
        /// <param name="iconFileName">Resources 目录下的图标文件名。返回按钮持有的图像在窗体关闭时释放。</param>
        /// <returns>已设置图标和紧凑边距的按钮。</returns>
        private Button ButtonOf(string text, string iconFileName)
        {
            Image icon = CreateActionIcon(iconFileName);
            _actionIcons.Add(icon);
            return new Button
            {
                Text = text,
                Image = icon,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                ImageAlign = ContentAlignment.MiddleLeft,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(2, 0, 2, 0),
                Margin = new Padding(0, 0, 3, 0),
                UseVisualStyleBackColor = true
            };
        }

        /// <summary>加载 16 像素操作图标，并在右侧留出 4 像素透明间隔，使文字不贴住图标。</summary>
        /// <param name="iconFileName">Resources 目录下的图标文件名。</param>
        /// <returns>调用方负责释放的位图。宽度为 20 像素。</returns>
        private static Image CreateActionIcon(string iconFileName)
        {
            using (Image source = DatabaseTreeImageCatalog.LoadActionIcon(iconFileName))
            {
                Bitmap spaced = new Bitmap(source.Width + 4, source.Height);
                try
                {
                    using (Graphics graphics = Graphics.FromImage(spaced))
                    {
                        graphics.Clear(Color.Transparent);
                        graphics.DrawImage(source, 0, 0, source.Width, source.Height);
                    }

                    return spaced;
                }
                catch
                {
                    spaced.Dispose();
                    throw;
                }
            }
        }

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

        /// <summary>新开页签并选中。载入的文本记为未修改基线。</summary>
        /// <param name="name">页签标题使用的方案名称。</param>
        /// <param name="sql">写入编辑器的 SQL。新建时为空。</param>
        /// <param name="id">方案标识。新建时是尚未落盘的新标识。</param>
        /// <param name="createdUtc">方案创建时间。</param>
        /// <param name="savedUtc">上次保存时间。尚未保存时为空。</param>
        /// <param name="persisted">该标识已经在查询方案仓储中时为 true。</param>
        /// <param name="createdBy">创建人。新建或旧方案没有记录时为空，第一次保存才写入。</param>
        /// <returns>新建的页签。编辑器由页签持有，关闭页签时释放。</returns>
        private QueryTab AddTab(string name, string sql, string id, DateTime createdUtc, DateTime? savedUtc, bool persisted, string createdBy)
        {
            SqlEditorControl editor = new SqlEditorControl(_formatting) { Dock = DockStyle.Fill };
            editor.SqlText = sql ?? string.Empty;
            QueryTab tab = new QueryTab(editor)
            {
                Id = id,
                Name = name,
                CreatedUtc = createdUtc,
                CreatedBy = createdBy,
                SavedUtc = savedUtc,
                SavedSql = editor.SqlText,
                IsPersisted = persisted
            };
            editor.SqlTextChanged += DocumentSqlTextChanged;
            editor.ExecuteSelectionRequested += DocumentExecuteSelectionRequested;

            TabPage page = new TabPage
            {
                Tag = tab,
                Padding = Padding.Empty,
                UseVisualStyleBackColor = true
            };
            tab.Page = page;
            page.Controls.Add(editor);
            RefreshTabTitle(tab);

            _adjustingTabs = true;
            try
            {
                _queryTabs.TabPages.Add(page);
                _queryTabs.SelectedTab = page;
            }
            finally
            {
                _adjustingTabs = false;
            }

            ShowActiveTab();
            editor.FocusEditor();
            return tab;
        }

        /// <summary>切换到指定页签。已经是当前页时只同步预览和方案信息。</summary>
        /// <param name="tab">要显示的页签。</param>
        private void SelectTab(QueryTab tab)
        {
            if (tab == null || tab.Page == null) return;
            if (_queryTabs.SelectedTab == tab.Page)
            {
                ShowActiveTab();
                return;
            }

            _queryTabs.SelectedTab = tab.Page;
        }

        /// <summary>关掉同一方案标识的页签。覆盖或删除时不再询问未保存修改。</summary>
        /// <param name="id">方案标识。</param>
        /// <param name="prompt">有未保存修改时是否询问。</param>
        /// <param name="except">不关闭的页签。没有时为空。</param>
        private void CloseTabs(string id, bool prompt, QueryTab except)
        {
            foreach (QueryTab tab in Tabs())
            {
                if (tab == except) continue;
                if (!string.Equals(tab.Id, id, StringComparison.OrdinalIgnoreCase)) continue;
                if (!TryCloseTab(tab, prompt)) return;
            }
        }

        /// <summary>关闭一个页签。有修改且需要询问时，保存失败或取消则留下该页。</summary>
        /// <param name="tab">要关闭的页签。</param>
        /// <param name="prompt">为 true 时，SQL 与上次保存不一致才询问。</param>
        /// <returns>页签已关闭时为 true。</returns>
        private bool TryCloseTab(QueryTab tab, bool prompt)
        {
            if (tab == null) return true;
            if (prompt && tab.IsDirty)
            {
                if (ActiveTab != tab)
                    SelectTab(tab);
                DialogResult answer = MessageBox.Show(
                    this,
                    string.Format("查询“{0}”尚未保存，是否保存？", tab.Name),
                    "SQL 查询",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button3);
                if (answer == DialogResult.Cancel) return false;
                if (answer == DialogResult.Yes && !SaveTab(tab)) return false;
            }

            RemoveTab(tab);
            if (_queryTabs.TabCount > 0)
                _status.Text = "已关闭查询。";
            return true;
        }

        /// <summary>从页签条移除并释放编辑器。最后一个页签关掉后显示新建提示。</summary>
        /// <param name="tab">要移除的页签。编辑器随页释放。</param>
        private void RemoveTab(QueryTab tab)
        {
            if (tab == null || tab.Page == null) return;
            tab.Editor.SqlTextChanged -= DocumentSqlTextChanged;
            tab.Editor.ExecuteSelectionRequested -= DocumentExecuteSelectionRequested;
            TabPage page = tab.Page;
            tab.Page = null;
            _adjustingTabs = true;
            try
            {
                _queryTabs.TabPages.Remove(page);
            }
            finally
            {
                _adjustingTabs = false;
            }

            page.Dispose();
            if (_queryTabs.TabCount == 0) ShowEmptyEditor();
            else ShowActiveTab();
        }

        /// <summary>按当前页签重画预览并同步左侧选中。没有页签时改为空状态。</summary>
        private void ShowActiveTab()
        {
            QueryTab tab = ActiveTab;
            if (tab == null)
            {
                ShowEmptyEditor();
                return;
            }

            _queryTabs.Visible = true;
            _emptyEditorHint.Visible = false;
            if (tab.HasPreview && tab.Preview != null)
            {
                BindPreview(tab.Preview);
                if (_resultsSplit.Panel2Collapsed)
                    EnsurePreviewVisible();
            }
            else
            {
                ClearPreviewGrid();
                _resultsSplit.Panel2Collapsed = true;
            }

            SyncProfileSelection(tab);
            UpdateQueryInfo();
        }

        /// <summary>隐藏页签条，在编辑区显示「请新建查询」，并折叠预览。</summary>
        private void ShowEmptyEditor()
        {
            _queryTabs.Visible = false;
            _emptyEditorHint.Visible = true;
            ClearPreviewGrid();
            _resultsSplit.Panel2Collapsed = true;
            SyncProfileSelection(null);
            UpdateQueryInfo();
            _status.Text = "请新建查询。";
        }

        /// <summary>让左侧列表高亮当前已保存方案。未保存或没有页签时取消选中，且不触发打开。</summary>
        /// <param name="tab">当前页签。为空表示编辑区是空的。</param>
        private void SyncProfileSelection(QueryTab tab)
        {
            _bindingProfiles = true;
            try
            {
                if (tab == null || !tab.IsPersisted)
                {
                    _queryProfiles.SelectedIndex = -1;
                    return;
                }

                for (int index = 0; index < _queryProfiles.Items.Count; index++)
                {
                    QueryItem item = _queryProfiles.Items[index] as QueryItem;
                    if (item != null && string.Equals(item.Profile.Id, tab.Id, StringComparison.OrdinalIgnoreCase))
                    {
                        _queryProfiles.SelectedIndex = index;
                        return;
                    }
                }

                _queryProfiles.SelectedIndex = -1;
            }
            finally
            {
                _bindingProfiles = false;
            }
        }

        /// <summary>把方案名、未保存标记和关闭标记写进页签文字。系统按这段文字计算宽度。</summary>
        /// <param name="tab">要更新标题的页签。</param>
        private void RefreshTabTitle(QueryTab tab)
        {
            if (tab == null || tab.Page == null || tab.Page.IsDisposed) return;
            string caption = tab.TabTitle + "  ×";
            if (string.Equals(tab.Page.Text, caption, StringComparison.Ordinal)) return;
            tab.Page.Text = caption;
        }

        /// <summary>编辑器文本变化时刷新对应页签的未保存标记。</summary>
        private void DocumentSqlTextChanged(object sender, EventArgs e)
        {
            if (IsDisposed || Disposing) return;
            SqlEditorControl editor = sender as SqlEditorControl;
            if (editor == null) return;
            foreach (QueryTab tab in Tabs())
            {
                if (tab.Editor != editor) continue;
                RefreshTabTitle(tab);
                return;
            }
        }

        /// <summary>只对当前页签响应选区预览，避免隐藏页签的编辑器触发执行。</summary>
        private void DocumentExecuteSelectionRequested(object sender, EventArgs e)
        {
            QueryTab tab = ActiveTab;
            if (tab != null && tab.Editor == sender)
                PreviewSelectedQuery();
        }

        /// <summary>格式化当前页签的 SQL。没有页签时提示新建。</summary>
        private void FormatActiveSql()
        {
            QueryTab tab = ActiveTab;
            if (tab == null)
            {
                PromptNewQuery();
                return;
            }

            tab.Editor.FormatSql();
        }

        /// <summary>注释或取消注释当前页签的 SQL。没有页签时提示新建。</summary>
        private void ToggleActiveComment()
        {
            QueryTab tab = ActiveTab;
            if (tab == null)
            {
                PromptNewQuery();
                return;
            }

            tab.Editor.ToggleComment();
        }

        /// <summary>提示用户先新建查询。</summary>
        private void PromptNewQuery()
        {
            MessageBox.Show(this, "请新建查询。", "SQL 查询", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>关闭窗体前逐个询问未保存的页签。任一次取消或保存失败则不关窗。</summary>
        /// <param name="sender">窗体。</param>
        /// <param name="e">关闭参数。取消时窗体保持打开。</param>
        private void QueryEditorFormFormClosing(object sender, FormClosingEventArgs e)
        {
            if (_closePromptActive) return;
            _closePromptActive = true;
            try
            {
                foreach (QueryTab tab in Tabs())
                {
                    if (tab.Page == null || tab.Editor.IsDisposed || !tab.IsDirty) continue;
                    if (ActiveTab != tab)
                        SelectTab(tab);
                    DialogResult answer = MessageBox.Show(
                        this,
                        string.Format("查询“{0}”尚未保存，是否保存？", tab.Name),
                        "SQL 查询",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button3);
                    if (answer == DialogResult.Cancel || (answer == DialogResult.Yes && !SaveTab(tab)))
                    {
                        e.Cancel = true;
                        return;
                    }
                }
            }
            finally
            {
                _closePromptActive = false;
            }
        }

        /// <summary>在按下页签文字里的叉时关掉该页。关闭推迟到鼠标消息结束，避免拆页签时打断页签控件自己的按下处理。</summary>
        private void QueryTabsMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            int index = TabIndexAt(e.Location);
            if (index < 0) return;
            if (!CloseGlyphBounds(_queryTabs.GetTabRect(index)).Contains(e.Location)) return;
            QueryTab tab = _queryTabs.TabPages[index].Tag as QueryTab;
            if (tab == null) return;
            BeginInvoke(new Action(() =>
            {
                if (IsDisposed || tab.Page == null) return;
                TryCloseTab(tab, true);
            }));
        }

        /// <summary>在页签上右键时先激活该页，再显示新建、保存、重命名。点在页签空白处时只有新建可用。</summary>
        private void QueryTabsMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            int index = TabIndexAt(e.Location);
            if (index >= 0 && _queryTabs.SelectedIndex != index)
                _queryTabs.SelectedIndex = index;
            _tabSaveMenuItem.Enabled = index >= 0;
            _tabRenameMenuItem.Enabled = index >= 0;
            _queryTabMenu.Show(_queryTabs, e.Location);
        }

        /// <summary>没有页签时，在提示文字上右键仍可新建查询。</summary>
        private void EmptyEditorMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            _tabSaveMenuItem.Enabled = false;
            _tabRenameMenuItem.Enabled = false;
            _queryTabMenu.Show(_emptyEditorHint, e.Location);
        }

        /// <summary>页签切换后恢复该页自己的预览，或在没有预览时折叠预览区。</summary>
        private void QueryTabsSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_adjustingTabs) return;
            ShowActiveTab();
        }

        /// <summary>按坐标找到页签序号。不在任何页签上时返回 -1。</summary>
        /// <param name="location">相对页签控件的坐标。</param>
        private int TabIndexAt(Point location)
        {
            for (int index = 0; index < _queryTabs.TabCount; index++)
            {
                if (_queryTabs.GetTabRect(index).Contains(location)) return index;
            }

            return -1;
        }

        /// <summary>计算页签文字右侧关闭标记所占的区域。宽度按「 ×」的字面测量。</summary>
        /// <param name="tabBounds">页签的客户区矩形。</param>
        private Rectangle CloseGlyphBounds(Rectangle tabBounds)
        {
            int width = TextRenderer.MeasureText(" ×", _queryTabs.Font).Width;
            width = Math.Max(16, width);
            return new Rectangle(tabBounds.Right - width, tabBounds.Top, width, tabBounds.Height);
        }

        /// <summary>创建页签右键菜单。保存和重命名作用于右键时激活的那一页。</summary>
        /// <param name="saveItem">保存项，没有点中页签时禁用。</param>
        /// <param name="renameItem">重命名项，没有点中页签时禁用。</param>
        private ContextMenuStrip BuildQueryTabMenu(out ToolStripMenuItem saveItem, out ToolStripMenuItem renameItem)
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("新建", null, (sender, args) => NewQuery());
            saveItem = new ToolStripMenuItem("保存", null, (sender, args) => SaveQuery());
            renameItem = new ToolStripMenuItem("重命名", null, (sender, args) => RenameQuery());
            menu.Items.Add(saveItem);
            menu.Items.Add(renameItem);
            return menu;
        }

        /// <summary>取得当前页签。没有选中页时为空。</summary>
        private QueryTab ActiveTab
        {
            get { return _queryTabs.SelectedTab == null ? null : _queryTabs.SelectedTab.Tag as QueryTab; }
        }

        /// <summary>复制当前打开的页签。调用方可以随后关闭其中某一页。</summary>
        private List<QueryTab> Tabs()
        {
            List<QueryTab> tabs = new List<QueryTab>();
            foreach (TabPage page in _queryTabs.TabPages)
            {
                QueryTab tab = page.Tag as QueryTab;
                if (tab != null) tabs.Add(tab);
            }

            return tabs;
        }

        /// <summary>按方案标识查找已经打开的页签。</summary>
        /// <param name="id">方案标识。</param>
        /// <returns>找到的页签。没有时为空。</returns>
        private QueryTab FindTab(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return null;
            foreach (QueryTab tab in Tabs())
            {
                if (string.Equals(tab.Id, id, StringComparison.OrdinalIgnoreCase))
                    return tab;
            }

            return null;
        }

        /// <summary>取「新建查询」后尚未被打开页签或已保存方案占用的最小正整数名称。</summary>
        private string NextNewQueryName()
        {
            HashSet<string> used = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            foreach (QueryTab tab in Tabs())
            {
                if (!string.IsNullOrWhiteSpace(tab.Name))
                    used.Add(tab.Name);
            }

            foreach (QueryProfile profile in _queries.GetAll())
            {
                if (!string.IsNullOrWhiteSpace(profile.Name))
                    used.Add(profile.Name);
            }

            int number = 1;
            string name = "新建查询" + number;
            while (used.Contains(name))
            {
                number++;
                name = "新建查询" + number;
            }

            return name;
        }

        /// <summary>为查询方案列表包装模型和显示名称。</summary>
        private sealed class QueryItem
        {
            public QueryItem(QueryProfile profile) { Profile = profile; }
            public QueryProfile Profile { get; }
            public string DisplayName => Profile.Name;
        }

        /// <summary>一份打开的查询。SQL 在自己的编辑器里，预览结果只留在内存里，切走时不另挂表格。</summary>
        /// <remarks>
        /// <see cref="SavedSql"/> 是载入或上次保存后的编辑器文本。与当前文本不一致时标题加 *。
        /// <see cref="IsPersisted"/> 为 false 时，重命名只改标题，不写仓储。
        /// </remarks>
        private sealed class QueryTab
        {
            /// <summary>创建页签状态并持有它的编辑器。</summary>
            /// <param name="editor">该页的 SQL 编辑器。由页签关闭时释放。</param>
            public QueryTab(SqlEditorControl editor)
            {
                Editor = editor ?? throw new ArgumentNullException(nameof(editor));
                SavedSql = string.Empty;
            }

            /// <summary>该页的 SQL 编辑器。</summary>
            public SqlEditorControl Editor { get; }

            /// <summary>承载编辑器的页。移除页签后置空，避免再改已释放的标题。</summary>
            public TabPage Page { get; set; }

            /// <summary>方案标识。新建时尚未落盘，保存后与仓储一致。</summary>
            public string Id { get; set; }

            /// <summary>页签上显示的方案名称，不含未保存标记。</summary>
            public string Name { get; set; }

            /// <summary>方案创建时间（UTC）。</summary>
            public DateTime CreatedUtc { get; set; }

            /// <summary>创建人。第一次落盘前为空；旧方案没有记录时保持为空。</summary>
            public string CreatedBy { get; set; }

            /// <summary>上次保存时间（UTC）。尚未保存时为空。</summary>
            public DateTime? SavedUtc { get; set; }

            /// <summary>载入或上次保存后的编辑器文本，用来判断有没有未保存修改。</summary>
            public string SavedSql { get; set; }

            /// <summary>方案已经写入仓储时为 true。</summary>
            public bool IsPersisted { get; set; }

            /// <summary>该页上次预览的缓冲结果。没有预览过时为空。</summary>
            public BufferedQueryResult Preview { get; set; }

            /// <summary>该页是否有可恢复的预览。取消首次预览后为 false。</summary>
            public bool HasPreview { get; set; }

            /// <summary>编辑器文本与 <see cref="SavedSql"/> 不一致时为 true。空的新建查询为 false。</summary>
            public bool IsDirty
            {
                get { return !string.Equals(Editor.SqlText ?? string.Empty, SavedSql ?? string.Empty, StringComparison.Ordinal); }
            }

            /// <summary>页签标题。有未保存修改时在名称末尾加 *。</summary>
            public string TabTitle
            {
                get { return (string.IsNullOrWhiteSpace(Name) ? "未保存" : Name) + (IsDirty ? "*" : string.Empty); }
            }
        }

    }
}

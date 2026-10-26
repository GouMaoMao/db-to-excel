using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Excel;
using DB2Sheet.Models;
using DB2Sheet.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.UI
{
    /// <summary>提供 SQL 方案编辑、只读查询预览以及结果写入 Excel 的主界面。</summary>
    /// <remarks>
    /// 每个页签是一份查询：SQL、方案名称和上次预览结果都留在该页。预览表格只有一份，切换页签时重画当前页的结果。
    /// 用户输入包括连接方案、SQL 和方案名称。写入时另问目标工作表，并可把 SQL 写入 SQL 页。
    /// 预览右键把已显示的结果复制到剪贴板，或另存为 xlsx。另存不重新查询，也不修改当前工作簿。
    /// 输出副作用包括保存查询方案和修改当前工作簿。页签在界面线程上增减，不访问 Excel。
    /// 第一次显示保持透明，分栏按最终尺寸摆完后再画出来。
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
        private readonly ImageList _queryTabIcons;
        private readonly Panel _editorSurface;
        private readonly Button _emptyQueryButton;
        private readonly Button _addQueryButton;
        private Panel _queryTabBar;
        private Panel _queryTabHost;
        private bool _placingQueryButton;
        private readonly DataGridView _preview;
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
        private readonly ContextMenuStrip _previewMenu;
        private ToolStripMenuItem _previewCopyItem;
        private ToolStripMenuItem _previewCopyHeadersItem;
        private ToolStripMenuItem _previewSelectAllItem;
        private ToolStripMenuItem _previewExportItem;
        private ToolStripMenuItem _previewCopyTitleItem;
        private int _previewHeaderColumn = -1;
        private bool _previewMenuFromMouse;
        private bool _previewMenuHeader;
        private bool _suppressPreviewMenu;
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
        private readonly AiAssistant _assistant;
        private readonly AiChatControl _aiChat;
        private readonly SplitContainer _editorSplit;
        private AiChatForm _aiChatForm;
        private bool _applyingAiSplit;
        private bool _userDraggingAiSplit;
        private bool _adjustingAiFrame;
        private int _toggleChatWidth = -1;

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
        /// <param name="assistant">AI 提问、模型、Skill 和审计。不执行 SQL。</param>
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
            AiAssistant assistant,
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
            _assistant = assistant ?? throw new ArgumentNullException(nameof(assistant));
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
            _queryTabs = CloseMark.CreateTabs();
            _queryTabIcons = DatabaseTreeImageCatalog.CreateSingleIconList("icon_script.png");
            _queryTabs.ImageList = _queryTabIcons;
            _queryTabs.Dock = DockStyle.Fill;
            _editorSurface = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Window };
            _emptyQueryButton = HintButton.Create(AppPresentation.EmptyQueryCaption);
            _emptyQueryButton.Click += (sender, args) => NewQuery();
            _profileTip.SetToolTip(_emptyQueryButton, AppPresentation.EmptyQueryTip);
            _editorSurface.Controls.Add(_emptyQueryButton);
            _editorSurface.Resize += (sender, args) => CenterEmptyQueryButton();
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
                ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable,
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
            _status = new ToolStripStatusLabel { Text = "就绪。", ForeColor = SystemColors.GrayText, Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            _queryProfileMenu = BuildQueryProfileMenu();
            _previewMenu = BuildPreviewMenu();
            _preview.ContextMenuStrip = _previewMenu;

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

            ToolStripMenuItem aiMenuItem = new ToolStripMenuItem("AI", null, (sender, args) => ToggleAiPanel());
            aiMenuItem.ToolTipText = AppPresentation.ToggleAiTip;
            MenuStrip menu = new MenuStrip { Dock = DockStyle.Top, ShowItemToolTips = true };
            menu.Items.Add(connectionMenu);
            menu.Items.Add(queryMenu);
            menu.Items.Add(editMenu);
            menu.Items.Add(aiMenuItem);
            MainMenuStrip = menu;

            Button previewButton = IconAction("icon_execute.png", AppPresentation.PreviewQueryTip, 28);
            Button exportButton = IconAction("icon_writeexcel.png", AppPresentation.ExportQueryTip, 28);
            Button saveButton = IconAction("icon_save.png", AppPresentation.SaveQueryTip, 28);
            Panel sqlHeader = BuildSqlHeader();
            _sqlLimitHintFont = new Font(AppPresentation.DefaultFontName, 8.25f, FontStyle.Regular);
            _sqlLimitHint = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                ForeColor = SystemColors.GrayText,
                Font = _sqlLimitHintFont,
                Padding = new Padding(8, 6, 8, 0),
                TextAlign = ContentAlignment.MiddleLeft
            };
            FlowLayoutPanel hintActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 6, 8, 0),
                Margin = Padding.Empty,
                BackColor = SystemColors.Control
            };
            hintActions.Controls.Add(previewButton);
            hintActions.Controls.Add(exportButton);
            hintActions.Controls.Add(saveButton);
            Panel hintRow = new Panel { Dock = DockStyle.Top, Height = 48, BackColor = SystemColors.Control };
            hintRow.Controls.Add(_sqlLimitHint);
            hintRow.Controls.Add(hintActions);

            Image addQueryIcon = DatabaseTreeImageCatalog.LoadActionIcon("icon_add.png");
            _actionIcons.Add(addQueryIcon);
            _addQueryButton = new Button
            {
                Dock = DockStyle.Right,
                Width = 28,
                Height = 28,
                Image = addQueryIcon,
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Control,
                Cursor = Cursors.Hand,
                Text = string.Empty,
                UseVisualStyleBackColor = false
            };
            _addQueryButton.FlatAppearance.BorderSize = 0;
            _addQueryButton.FlatAppearance.MouseOverBackColor = SystemColors.ControlDark;
            _addQueryButton.FlatAppearance.MouseDownBackColor = SystemColors.ControlDarkDark;
            _addQueryButton.Click += (sender, args) => NewQuery();
            _profileTip.SetToolTip(_addQueryButton, AppPresentation.NewQueryTip);
            _queryTabBar = new Panel { Dock = DockStyle.Top, Height = 28, BackColor = SystemColors.Control };
            _queryTabHost = new Panel { Dock = DockStyle.Left, Width = 28, Height = 28, BackColor = SystemColors.Control };
            _queryTabHost.Controls.Add(_queryTabs);
            _queryTabHost.Controls.Add(_addQueryButton);
            _queryTabBar.Controls.Add(_queryTabHost);
            _queryTabBar.Resize += (sender, args) => PlaceNewQueryButton();
            _queryTabs.HandleCreated += (sender, args) => PlaceNewQueryButton();
            Panel editorHost = new Panel { Dock = DockStyle.Fill };
            editorHost.Controls.Add(_editorSurface);
            editorHost.Controls.Add(_queryTabBar);
            Panel sqlPanel = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Control };
            sqlPanel.Controls.Add(editorHost);
            sqlPanel.Controls.Add(hintRow);
            sqlPanel.Controls.Add(sqlHeader);

            Panel previewPanel = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Control };
            previewPanel.Controls.Add(_preview);
            previewPanel.Controls.Add(SectionHeader("结果预览"));

            _resultsSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                Panel2Collapsed = true,
                FixedPanel = FixedPanel.None
            };
            _resultsSplit.Panel1.Controls.Add(LightBorder.Wrap(sqlPanel));
            _resultsSplit.Panel2.Controls.Add(LightBorder.Wrap(previewPanel));

            _aiChat = new AiChatControl(_assistant, _operationRunner, _settings, ReadEditorForAi, ReadMentions, ToggleAiPanel);
            _editorSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6,
                Panel2Collapsed = true
            };
            _editorSplit.Panel1.Controls.Add(_resultsSplit);
            _editorSplit.Panel2.Controls.Add(LightBorder.Wrap(_aiChat));

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
            _workspace.Panel2.Controls.Add(_editorSplit);

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
            _editorSurface.MouseUp += EmptyEditorMouseUp;
            _emptyQueryButton.MouseUp += EmptyEditorMouseUp;
            _queryProfiles.SelectedIndexChanged += QueryProfilesSelectedIndexChanged;
            _queryProfiles.MouseUp += QueryProfilesMouseUp;
            _queryProfiles.MouseMove += QueryProfilesMouseMove;
            _queryProfiles.MouseLeave += QueryProfilesMouseLeave;
            _queryProfiles.DrawItem += QueryProfilesDrawItem;
            _profileTipTimer.Interval = 500;
            _profileTipTimer.Tick += ProfileTipTimerTick;
            _preview.RowPostPaint += PreviewRowPostPaint;
            _preview.CellPainting += PreviewCellPainting;
            _preview.MouseDown += PreviewMouseDown;
            _preview.KeyDown += PreviewGridKeyDown;
            _previewMenu.Opening += PreviewMenuOpening;
            _connectionTree.NoticeChanged += (sender, args) => _status.Text = _connectionTree.Notice;
            _connectionTree.SelectionChanged += (sender, args) => UpdateConnectionMenuState();
            FormClosing += QueryEditorFormFormClosing;
            FormClosed += QueryEditorFormFormClosed;
            Shown += (sender, args) => PlaceNewQueryButton();
            Activated += (sender, args) =>
            {
                UpdateSqlLimitHint();
                _aiChat.ReloadChoices();
            };
            UpdateSqlLimitHint();
            _workspace.SizeChanged += (sender, args) =>
            {
                if (!_adjustingAiFrame)
                    ApplyWorkspaceSplit();
            };
            _workspace.SplitterMoving += (sender, args) => _userDraggingWorkspaceSplit = true;
            _workspace.SplitterMoved += (sender, args) => RememberWorkspaceSplit(_workspace);
            _editorSplit.SizeChanged += (sender, args) =>
            {
                if (!_adjustingAiFrame)
                    ApplyAiSplit();
            };
            _editorSplit.SplitterMoving += (sender, args) => _userDraggingAiSplit = true;
            _editorSplit.SplitterMoved += (sender, args) => RememberAiSplit();

            ReloadQueries(null);
            NewQuery();
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
                _status.Text = "查询方案已重命名。";
                return;
            }

            stored.Name = queryName;
            _queries.Save(stored);
            QueryProfile refreshed = _queries.GetById(tab.Id);
            tab.SavedUtc = refreshed != null ? refreshed.UpdatedUtc : tab.SavedUtc;
            RefreshTabTitle(tab);
            ReloadQueries(tab.Id);
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

        /// <summary>SQL 编辑区标题行。只放固定标题。</summary>
        /// <returns>已排好的标题行面板。</returns>
        private Panel BuildSqlHeader()
        {
            Panel header = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = SystemColors.ControlLight };
            Label title = new Label
            {
                Text = "SQL编辑器（仅限只读查询）",
                AutoSize = true,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 6, 12, 0),
                Font = new Font(AppPresentation.DefaultFontName, AppPresentation.DefaultFontSize, FontStyle.Bold)
            };
            header.Controls.Add(title);
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
                ShowError(exception, sqlText);
            }
        }

        /// <summary>把当前 SQL 的结果写入用户确认的工作表。表已存在时先确认覆盖。</summary>
        /// <remarks>
        /// 勾选写入 SQL 时，结果写成功后才把 SQL 按列写入 SQL 页。取消或覆盖被拒绝时两者都不写。
        /// 写入成功且「保存后美化」为开时，再走一遍美化进度窗。这次美化取消或失败不撤销已经写入的结果和 SQL。
        /// 正常结束后弹出与状态栏相同的一句，必须点确定。美化失败只留错误详情，不再叠一层写入成功框。取消不弹这层框。
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

            string sqlText = tab.Editor.SqlText;
            bool failedDuringQuery = false;
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

                DataSourceRequest request = CreateRequest(connection, sqlText, ExecutionPurpose.Export, _settings.Get(CoreSettings.MaxExportRows));
                failedDuringQuery = true;
                bool truncated = false;
                long rows = _operationRunner.Run(this, "写入 Excel", true, async (progress, token) =>
                {
                    string operationId = Guid.NewGuid().ToString("N");
                    BufferedQueryResult result = await _queryBuffer.ExecuteAsync(
                        request, operationId, progress, token);
                    failedDuringQuery = false;
                    truncated = result.IsTruncated;
                    long written = _writer.WriteResult(_workbook, targetSheet, result, progress, token, operationId);
                    if (writeSql)
                    {
                        progress.Report(new OperationProgress
                        {
                            Stage = OperationStage.Writing,
                            TargetSheetName = "SQL",
                            Message = "结果已写入，正在保存本次 SQL 任务列…",
                            Percent = 100,
                            IsIndeterminate = false
                        });
                        _taskReader.WriteTaskColumn(_workbook, targetSheet, sqlText);
                        progress.Report(new OperationProgress
                        {
                            Stage = OperationStage.Completed,
                            TargetSheetName = "SQL",
                            Message = "SQL 任务列保存完成。",
                            Percent = 100,
                            IsIndeterminate = false
                        });
                    }
                    return written;
                });
                SavedSqlBeautify beautify = writeSql ? BeautifySavedSql() : SavedSqlBeautify.Skipped;
                if (beautify == SavedSqlBeautify.Failed)
                    return;
                ShowOutcome(
                    "写入 Sheet",
                    DescribeSheetWrite(targetSheet, rows, truncated, writeSql, beautify),
                    truncated);
            }
            catch (OperationCanceledException)
            {
                _status.Text = "写入已取消。";
            }
            catch (Exception exception)
            {
                ShowError(exception, failedDuringQuery ? sqlText : null);
            }
        }

        /// <summary>按「保存后美化」在写入成功后再美化整张 SQL 页。</summary>
        /// <returns>美化是否做完。失败时错误详情已经弹出。</returns>
        /// <remarks>
        /// 开关关闭或没有可美化的列时返回跳过，不把这次写入当成失败。
        /// 取消或失败不撤销已经写入的结果和 SQL。完成说明由调用方统一弹出。
        /// </remarks>
        private SavedSqlBeautify BeautifySavedSql()
        {
            if (!_settings.Get(CoreSettings.BeautifyAfterSave))
                return SavedSqlBeautify.Skipped;
            try
            {
                int count = _taskReader.CountStyleColumns(_workbook);
                if (count == 0)
                    return SavedSqlBeautify.Skipped;
                _operationRunner.Run(
                    this,
                    "美化格式",
                    true,
                    (progress, token) => _taskReader.FormatSqlSheet(_workbook, progress, token));
                return SavedSqlBeautify.Done;
            }
            catch (OperationCanceledException)
            {
                return SavedSqlBeautify.Cancelled;
            }
            catch (Exception exception)
            {
                ShowError(exception);
                return SavedSqlBeautify.Failed;
            }
        }

        /// <summary>组装写入结束后的一句说明。截断时写出上限，并指向「最大导出行数」。</summary>
        /// <param name="targetSheet">刚写入的结果表名。</param>
        /// <param name="rows">已写入的数据行数，不含标题。</param>
        /// <param name="truncated">查询结果达到最大导出行数后被截断时为 true。</param>
        /// <param name="writeSql">这次是否把 SQL 写入了工作表 SQL。</param>
        /// <param name="beautify">保存后美化的结果。失败不会走到这里。</param>
        /// <returns>与状态栏和确定框相同的一句。</returns>
        private string DescribeSheetWrite(
            string targetSheet, long rows, bool truncated, bool writeSql, SavedSqlBeautify beautify)
        {
            if (beautify == SavedSqlBeautify.Cancelled)
            {
                if (!truncated)
                    return "SQL 已写入，美化已取消。";
                return string.Format(
                    "SQL 已写入，美化已取消。已达上限 {0:N0} 行并截断。可在设置中修改“最大导出行数”。",
                    _settings.Get(CoreSettings.MaxExportRows));
            }

            string extra = beautify == SavedSqlBeautify.Done
                ? "，并已美化 SQL 页"
                : writeSql
                    ? "，并已把 SQL 写入工作表 SQL"
                    : string.Empty;
            if (!truncated)
                return string.Format("已向“{0}”写入 {1:N0} 行{2}。", targetSheet, rows, extra);

            return string.Format(
                "已向“{0}”写入 {1:N0} 行{2}，已达上限 {3:N0} 行并截断。可在设置中修改“最大导出行数”。",
                targetSheet,
                rows,
                extra,
                _settings.Get(CoreSettings.MaxExportRows));
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

        /// <summary>首次显示在最终布局完成前保持完全透明。</summary>
        protected override bool DeferFirstShow
        {
            get { return true; }
        }

        /// <summary>
        /// 窗口已经按最终尺寸显示，但仍然完全透明。
        /// 在这里摆左右分栏，并把左栏上下分成两半。显示之后再改分割条会把旧位置留在屏幕上。
        /// </summary>
        protected override void PrepareFirstShow()
        {
            ApplyWorkspaceSplit();
            TrySetSplitterDistance(_leftSplit, _leftSplit.Height / 2);
            ApplyAiSplit();
            if (_settings.Get(CoreSettings.AiChatVisible))
                _aiChat.EnsureConversation();
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

        /// <summary>展开或收起右侧 AI 对话。展开时加宽窗体，收起时收回；宽度只有拖动分割条才记住。</summary>
        private void ToggleAiPanel()
        {
            bool visible = !_settings.Get(CoreSettings.AiChatVisible);
            _settings.Set(CoreSettings.AiChatVisible, visible);
            if (visible)
                _toggleChatWidth = WidenForAiPanel();
            else
                NarrowAfterAiPanel();
            ApplyAiSplit();
            _toggleChatWidth = -1;
            if (visible)
                _aiChat.EnsureConversation();
        }

        /// <summary>
        /// 按当前 SQL 列宽度算出对话栏，并把窗体向右加宽。
        /// 左栏像素宽度保持不变。最大化或右侧没有余地时不加宽，多出来的宽度由 SQL 列让出。
        /// </summary>
        /// <returns>希望的对话栏像素宽度。SQL 列还没有宽度时为 -1。</returns>
        private int WidenForAiPanel()
        {
            if (_editorSplit == null || _editorSplit.IsDisposed || !_editorSplit.Panel2Collapsed) return -1;
            int sqlWidth = _editorSplit.Width;
            if (sqlWidth <= 0) return -1;
            int chatWidth = ChatWidthBesideSql(sqlWidth);
            if (WindowState == FormWindowState.Maximized) return chatWidth;

            int growth = chatWidth + _editorSplit.SplitterWidth;
            Rectangle area = Screen.FromControl(this).WorkingArea;
            int room = Math.Max(0, area.Right - Bounds.Right);
            int applied = Math.Min(growth, room);
            if (applied <= 0) return chatWidth;

            int left = _workspace.SplitterDistance;
            _adjustingAiFrame = true;
            try
            {
                Width += applied;
                TrySetSplitterDistance(_workspace, left);
            }
            finally
            {
                _adjustingAiFrame = false;
            }

            return chatWidth;
        }

        /// <summary>收起对话栏时把窗体收回对话栏和分割条所占的宽度，SQL 列像素宽度保持不变。</summary>
        private void NarrowAfterAiPanel()
        {
            if (WindowState == FormWindowState.Maximized) return;
            if (_editorSplit == null || _editorSplit.IsDisposed || _editorSplit.Panel2Collapsed) return;
            int remove = _editorSplit.Width - _editorSplit.SplitterDistance;
            if (remove <= 0) return;
            int left = _workspace.SplitterDistance;
            int floor = Math.Max(MinimumSize.Width, 200);
            _adjustingAiFrame = true;
            try
            {
                Width = Math.Max(floor, Width - remove);
                TrySetSplitterDistance(_workspace, left);
            }
            finally
            {
                _adjustingAiFrame = false;
            }
        }

        /// <summary>由 SQL 列宽度反推对话栏宽度。还没拖过时为 SQL 列的 2/3。</summary>
        /// <param name="sqlWidth">对话栏收起时的 SQL 列宽度。</param>
        /// <returns>对话栏像素宽度。</returns>
        private int ChatWidthBesideSql(int sqlWidth)
        {
            int ratio = _settings.Get(CoreSettings.AiChatSplitRatio);
            if (ratio > 0 && ratio < 1000)
                return Math.Max(1, sqlWidth * ratio / (1000 - ratio));
            return Math.Max(1, sqlWidth * 2 / 3);
        }

        /// <summary>按已记住的千分比摆右侧对话栏。还没记住时，对话栏占 SQL 列与对话栏合计宽度的 2/5，即 SQL 列的 2/3。</summary>
        private void ApplyAiSplit()
        {
            if (_applyingAiSplit || _editorSplit == null || _editorSplit.IsDisposed) return;
            _applyingAiSplit = true;
            try
            {
                bool visible = _settings.Get(CoreSettings.AiChatVisible);
                if (_editorSplit.Panel2Collapsed == visible)
                    _editorSplit.Panel2Collapsed = !visible;
                if (!visible) return;

                _userDraggingAiSplit = false;
                ApplySplitterPanelMinSizes(_editorSplit, 420, 280);
                int available = _editorSplit.Width - _editorSplit.SplitterWidth;
                if (available <= 0) return;
                int chatWidth = _toggleChatWidth;
                if (chatWidth < 0)
                {
                    int ratio = _settings.Get(CoreSettings.AiChatSplitRatio);
                    chatWidth = ratio > 0 ? available * ratio / 1000 : available * 2 / 5;
                }

                TrySetSplitterDistance(_editorSplit, available - chatWidth);
            }
            finally
            {
                _applyingAiSplit = false;
            }
        }

        /// <summary>用户拖动右侧分割条后，把对话栏宽度写成千分比。</summary>
        private void RememberAiSplit()
        {
            if (!_userDraggingAiSplit) return;
            _userDraggingAiSplit = false;
            if (_applyingAiSplit || _editorSplit == null || _editorSplit.IsDisposed || _editorSplit.Panel2Collapsed) return;
            int available = _editorSplit.Width - _editorSplit.SplitterWidth;
            if (available <= 0) return;
            int chatWidth = available - _editorSplit.SplitterDistance;
            int ratio = (int)Math.Round(chatWidth * 1000d / available);
            ratio = Math.Max(1, Math.Min(999, ratio));
            if (ratio == _settings.Get(CoreSettings.AiChatSplitRatio)) return;
            _settings.Set(CoreSettings.AiChatSplitRatio, ratio);
        }

        /// <summary>让对话控件刷新输入框上方的当前页签引用。</summary>
        private void RefreshAiReference()
        {
            if (_aiChat != null && !_aiChat.IsDisposed)
                _aiChat.NotifyEditorChanged();
            if (_aiChatForm != null && !_aiChatForm.IsDisposed)
                _aiChatForm.NotifyEditorChanged();
        }

        /// <summary>读取当前页签全文和当前连接的数据库类型。没有页签时标明为空。</summary>
        /// <returns>发送当时的编辑器快照。</returns>
        private AiEditorSnapshot ReadEditorForAi()
        {
            QueryTab tab = ActiveTab;
            ConnectionProfileSnapshot connection = _connectionTree.ActiveConnection;
            return new AiEditorSnapshot
            {
                HasTab = tab != null,
                TabName = tab == null ? string.Empty : tab.TabTitle,
                Sql = tab == null ? string.Empty : tab.Editor.SqlText,
                ProviderId = connection == null ? string.Empty : connection.ProviderId
            };
        }

        /// <summary>列出已保存方案，并带上当前页已经有的预览。供输入 @ 时选用。</summary>
        /// <returns>方案和预览快照。预览按行截断，不重新查询。</returns>
        private AiMentionSource ReadMentions()
        {
            List<AiSavedQuery> queries = new List<AiSavedQuery>();
            foreach (QueryProfile profile in _queries.GetAll())
            {
                if (profile == null) continue;
                queries.Add(new AiSavedQuery { Name = profile.Name, Sql = profile.QueryText });
            }

            QueryTab tab = ActiveTab;
            bool hasPreview = tab != null && tab.HasPreview && tab.Preview != null;
            return new AiMentionSource
            {
                Queries = queries,
                HasPreview = hasPreview,
                PreviewLabel = tab == null ? string.Empty : tab.TabTitle,
                PreviewText = hasPreview ? FormatPreviewForAi(tab.Preview) : string.Empty
            };
        }

        /// <summary>把预览收成制表符分隔的文本。最多 30 行、约 16KB。</summary>
        /// <param name="result">当前页已有的预览。</param>
        /// <returns>可放进引用的文本。没有列时为空。</returns>
        private static string FormatPreviewForAi(BufferedQueryResult result)
        {
            const int maxRows = 30;
            const int maxChars = 16 * 1024;
            if (result == null || result.Columns.Count == 0) return string.Empty;
            StringBuilder builder = new StringBuilder();
            for (int index = 0; index < result.Columns.Count; index++)
            {
                if (index > 0)
                    builder.Append('\t');
                builder.Append(result.Columns[index].Name ?? string.Empty);
            }

            builder.AppendLine();
            int written = 0;
            foreach (object[,] block in result.Blocks)
            {
                int rows = block.GetLength(0);
                int columns = block.GetLength(1);
                for (int row = 0; row < rows && written < maxRows && builder.Length < maxChars; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        if (column > 0)
                            builder.Append('\t');
                        object value = block[row, column];
                        builder.Append(value == null || value is DBNull ? string.Empty : Convert.ToString(value));
                    }

                    builder.AppendLine();
                    written++;
                }
            }

            if (written < result.RowCount || result.IsTruncated || builder.Length >= maxChars)
                builder.Append("（预览已截断，以上为前 ").Append(written.ToString()).Append(" 行。）");
            if (builder.Length <= maxChars)
                return builder.ToString();
            return builder.ToString(0, maxChars) + Environment.NewLine + "（预览已截断。）";
        }

        /// <summary>弹出独立对话窗，新建「SQL 报错」对话并填入脱敏草稿。</summary>
        /// <param name="sql">失败的那次 SQL。</param>
        /// <param name="errorMessage">异常消息，不含堆栈。</param>
        private void OpenAiErrorChat(string sql, string errorMessage)
        {
            if (_aiChatForm == null || _aiChatForm.IsDisposed)
            {
                _aiChatForm = new AiChatForm(_assistant, _operationRunner, _settings, ReadEditorForAi, ReadMentions);
                _aiChatForm.FormClosed += (sender, args) => _aiChatForm = null;
            }

            _aiChatForm.StartErrorConversation(AiTextRedactor.BuildErrorDraft(sql, errorMessage));
            if (!_aiChatForm.Visible)
                _aiChatForm.Show(this);
            else
                _aiChatForm.Activate();
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

        /// <summary>建立预览表右键菜单。点单元格和点列标题时显示的项不同，打开前再切换。</summary>
        /// <returns>挂到预览表上的菜单。窗体关闭时释放。</returns>
        private ContextMenuStrip BuildPreviewMenu()
        {
            _previewCopyItem = PreviewMenuItem("复制", "Ctrl+C", AppPresentation.PreviewCopyTip, (sender, args) => CopyPreview(false));
            _previewCopyHeadersItem = PreviewMenuItem("复制带标题", "Ctrl+Shift+C", AppPresentation.PreviewCopyHeadersTip, (sender, args) => CopyPreview(true));
            _previewSelectAllItem = PreviewMenuItem("全选", "Ctrl+A", AppPresentation.PreviewSelectAllTip, (sender, args) => SelectEntirePreview());
            _previewExportItem = PreviewMenuItem("导出为 Excel", null, AppPresentation.PreviewExportTip, (sender, args) => ExportPreview());
            _previewCopyTitleItem = PreviewMenuItem("复制标题", null, AppPresentation.PreviewCopyTitleTip, (sender, args) => CopyPreviewTitle());
            ContextMenuStrip menu = new ContextMenuStrip { ShowItemToolTips = true };
            menu.Items.AddRange(new ToolStripItem[]
            {
                _previewCopyItem,
                _previewCopyHeadersItem,
                _previewSelectAllItem,
                _previewExportItem,
                _previewCopyTitleItem
            });
            return menu;
        }

        /// <summary>创建一项预览菜单。快捷键只显示，不注册到窗体，避免抢走 SQL 编辑器的按键。</summary>
        /// <param name="text">菜单文字。</param>
        /// <param name="shortcut">显示用的快捷键。没有时为空。</param>
        /// <param name="tip">悬浮说明。概述和详述用换行分开。</param>
        /// <param name="click">点击时在界面线程执行的动作。</param>
        /// <returns>尚未加入菜单的项。</returns>
        private static ToolStripMenuItem PreviewMenuItem(string text, string shortcut, string tip, EventHandler click)
        {
            ToolStripMenuItem item = new ToolStripMenuItem(text, null, click)
            {
                AutoToolTip = false,
                ToolTipText = tip
            };
            if (!string.IsNullOrEmpty(shortcut))
                item.ShortcutKeyDisplayString = shortcut;
            return item;
        }

        /// <summary>按右键落点调整选区，并记下这次要显示单元格菜单还是列标题菜单。</summary>
        /// <param name="sender">预览表。</param>
        /// <param name="e">鼠标事件。只处理右键。</param>
        /// <remarks>调整选区必须发生在菜单打开之前。滚动条和列宽拖动不弹出菜单。</remarks>
        private void PreviewMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;

            _previewMenuFromMouse = true;
            _previewMenuHeader = false;
            _suppressPreviewMenu = false;
            DataGridView.HitTestInfo hit = _preview.HitTest(e.X, e.Y);
            switch (hit.Type)
            {
                case DataGridViewHitTestType.Cell:
                    SelectPreviewCellIfNeeded(hit.RowIndex, hit.ColumnIndex);
                    break;
                case DataGridViewHitTestType.RowHeader:
                    SelectPreviewRowIfNeeded(hit.RowIndex);
                    break;
                case DataGridViewHitTestType.ColumnHeader:
                    _previewMenuHeader = true;
                    _previewHeaderColumn = hit.ColumnIndex;
                    SelectPreviewColumn(hit.ColumnIndex);
                    break;
                case DataGridViewHitTestType.None:
                case DataGridViewHitTestType.TopLeftHeader:
                    break;
                default:
                    _suppressPreviewMenu = true;
                    break;
            }
        }

        /// <summary>预览表有焦点时响应复制和全选。SQL 编辑器不经过这里。</summary>
        /// <param name="sender">预览表。</param>
        /// <param name="e">按键。处理过后不再交给表格自己的剪贴板。</param>
        private void PreviewGridKeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.Alt) return;
            if (e.KeyCode == Keys.A && !e.Shift)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectEntirePreview();
                return;
            }

            if (e.KeyCode != Keys.C) return;
            e.Handled = true;
            e.SuppressKeyPress = true;
            CopyPreview(e.Shift);
        }

        /// <summary>打开菜单前按点击位置留下四项或只留「复制标题」，并按有没有数据行决定能否点。</summary>
        /// <param name="sender">预览菜单。</param>
        /// <param name="e">取消时这次不弹出。</param>
        /// <remarks>键盘弹出菜单时没有右键落点，按单元格菜单处理。</remarks>
        private void PreviewMenuOpening(object sender, CancelEventArgs e)
        {
            bool fromMouse = _previewMenuFromMouse;
            _previewMenuFromMouse = false;
            if (fromMouse && _suppressPreviewMenu)
            {
                e.Cancel = true;
                return;
            }

            bool header = fromMouse && _previewMenuHeader;
            _previewCopyItem.Visible = !header;
            _previewCopyHeadersItem.Visible = !header;
            _previewSelectAllItem.Visible = !header;
            _previewExportItem.Visible = !header;
            _previewCopyTitleItem.Visible = header;

            bool hasRows = _preview.Rows.Count > 0;
            _previewCopyItem.Enabled = hasRows;
            _previewCopyHeadersItem.Enabled = hasRows;
            _previewSelectAllItem.Enabled = hasRows;
            QueryTab tab = ActiveTab;
            _previewExportItem.Enabled = tab != null && tab.HasPreview && tab.Preview != null && tab.Preview.Columns.Count > 0;
            _previewCopyTitleItem.Enabled = _previewHeaderColumn >= 0 && _previewHeaderColumn < _preview.Columns.Count;
        }

        /// <summary>点在未选中的格子上时，改为只选中该格。已经在选区里则保留选区。</summary>
        /// <param name="rowIndex">行号，从 0 开始。</param>
        /// <param name="columnIndex">列号，从 0 开始。</param>
        private void SelectPreviewCellIfNeeded(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || columnIndex < 0 || rowIndex >= _preview.Rows.Count || columnIndex >= _preview.Columns.Count) return;
            DataGridViewCell cell = _preview.Rows[rowIndex].Cells[columnIndex];
            if (cell.Selected) return;
            _preview.ClearSelection();
            _preview.CurrentCell = cell;
            cell.Selected = true;
        }

        /// <summary>点在行号上时，若该行没有整行进入选区，改为只选中这一行。</summary>
        /// <param name="rowIndex">行号，从 0 开始。</param>
        private void SelectPreviewRowIfNeeded(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _preview.Rows.Count) return;
            DataGridViewRow row = _preview.Rows[rowIndex];
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.Selected) continue;
                _preview.ClearSelection();
                row.Selected = true;
                if (row.Cells.Count > 0)
                    _preview.CurrentCell = row.Cells[0];
                return;
            }
        }

        /// <summary>点列标题时选中整列，供随后复制标题对应这一列。</summary>
        /// <param name="columnIndex">列号，从 0 开始。</param>
        private void SelectPreviewColumn(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= _preview.Columns.Count || _preview.Rows.Count == 0) return;
            _preview.ClearSelection();
            _preview.CurrentCell = _preview.Rows[0].Cells[columnIndex];
            foreach (DataGridViewRow row in _preview.Rows)
                row.Cells[columnIndex].Selected = true;
        }

        /// <summary>选中全部数据行。没有数据行时不做。</summary>
        private void SelectEntirePreview()
        {
            if (_preview.Rows.Count == 0) return;
            _preview.SelectAll();
        }

        /// <summary>把选区或整张预览放进剪贴板。带标题时在数据前加上字段编码和类型短名。</summary>
        /// <param name="withHeaders">为 true 时加两行表头。Ctrl+Shift+C 走这条。</param>
        /// <remarks>必须在界面线程调用。剪贴板被占用时弹出错误详情。没有数据行时直接返回。</remarks>
        private void CopyPreview(bool withHeaders)
        {
            if (_preview.Rows.Count == 0) return;
            try
            {
                CopyPreviewTable(withHeaders);
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        /// <summary>按当前选区或整张预览组制表符文本和 HTML。带标题时先写字段编码和类型短名。</summary>
        /// <param name="withHeaders">为 true 时在数据前加两行表头。</param>
        /// <remarks>
        /// 不改选区。不使用表格自带的剪贴板内容：那份 HTML 没有字符集，Excel 会把中文读成乱码。
        /// 单元格里的换行写成 br，贴进 Excel 时仍留在同一格。
        /// </remarks>
        private void CopyPreviewTable(bool withHeaders)
        {
            List<int> rowIndexes;
            List<int> columnIndexes;
            bool entire;
            CollectCopyAxes(out rowIndexes, out columnIndexes, out entire);
            if (columnIndexes.Count == 0) return;

            StringBuilder text = new StringBuilder();
            StringBuilder html = new StringBuilder();
            html.Append("<table>");
            string[] names = new string[columnIndexes.Count];
            string[] types = new string[columnIndexes.Count];
            for (int index = 0; index < columnIndexes.Count; index++)
            {
                DataGridViewColumn column = _preview.Columns[columnIndexes[index]];
                ResultColumn definition = column.Tag as ResultColumn;
                names[index] = definition == null ? (column.HeaderText ?? string.Empty) : (definition.Name ?? string.Empty);
                types[index] = FormatPreviewTypeName(definition);
            }

            if (withHeaders)
            {
                AppendClipboardRow(text, html, names, true);
                AppendClipboardRow(text, html, types, false);
            }

            string[] values = new string[columnIndexes.Count];
            bool firstDataRow = !withHeaders;
            foreach (int rowIndex in rowIndexes)
            {
                DataGridViewRow row = _preview.Rows[rowIndex];
                for (int index = 0; index < columnIndexes.Count; index++)
                {
                    DataGridViewCell cell = row.Cells[columnIndexes[index]];
                    values[index] = entire || cell.Selected ? PreviewCellText(cell) : string.Empty;
                }

                AppendClipboardRow(text, html, values, firstDataRow);
                firstDataRow = false;
            }

            html.Append("</table>");
            SetPreviewClipboard(text.ToString(), html.ToString());
            _status.Text = entire ? "已复制全部预览。" : "已复制选区。";
        }

        /// <summary>复制右键那一列的字段编码和类型短名。没有类型时只留字段编码。</summary>
        /// <remarks>必须在界面线程调用。剪贴板被占用时弹出错误详情。</remarks>
        private void CopyPreviewTitle()
        {
            if (_previewHeaderColumn < 0 || _previewHeaderColumn >= _preview.Columns.Count) return;
            try
            {
                DataGridViewColumn column = _preview.Columns[_previewHeaderColumn];
                ResultColumn definition = column.Tag as ResultColumn;
                string name = definition == null ? (column.HeaderText ?? string.Empty) : (definition.Name ?? string.Empty);
                string typeName = FormatPreviewTypeName(definition);
                StringBuilder text = new StringBuilder();
                StringBuilder html = new StringBuilder();
                html.Append("<table>");
                AppendClipboardRow(text, html, new[] { name }, true);
                if (!string.IsNullOrEmpty(typeName))
                    AppendClipboardRow(text, html, new[] { typeName }, false);
                html.Append("</table>");
                SetPreviewClipboard(text.ToString(), html.ToString());
                _status.Text = "已复制标题。";
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        /// <summary>确定这次复制覆盖哪些行和列。没有选区时覆盖全部数据行，列按从左到右的显示顺序。</summary>
        /// <param name="rowIndexes">要复制的行号，从上到下。</param>
        /// <param name="columnIndexes">要复制的列号，从左到右。</param>
        /// <param name="entire">没有选区、因此复制整张预览时为 true。</param>
        private void CollectCopyAxes(out List<int> rowIndexes, out List<int> columnIndexes, out bool entire)
        {
            entire = _preview.SelectedCells.Count == 0;
            rowIndexes = new List<int>();
            columnIndexes = new List<int>();
            bool[] rowHit = entire ? null : new bool[_preview.Rows.Count];
            bool[] columnHit = entire ? null : new bool[_preview.Columns.Count];
            if (!entire)
            {
                foreach (DataGridViewCell cell in _preview.SelectedCells)
                {
                    if (cell.RowIndex >= 0 && cell.RowIndex < rowHit.Length)
                        rowHit[cell.RowIndex] = true;
                    if (cell.ColumnIndex >= 0 && cell.ColumnIndex < columnHit.Length)
                        columnHit[cell.ColumnIndex] = true;
                }
            }

            if (entire)
            {
                for (int row = 0; row < _preview.Rows.Count; row++)
                    rowIndexes.Add(row);
            }
            else
            {
                for (int row = 0; row < rowHit.Length; row++)
                {
                    if (rowHit[row])
                        rowIndexes.Add(row);
                }
            }

            for (DataGridViewColumn column = _preview.Columns.GetFirstColumn(DataGridViewElementStates.Visible);
                column != null;
                column = _preview.Columns.GetNextColumn(column, DataGridViewElementStates.Visible, DataGridViewElementStates.None))
            {
                if (entire || columnHit[column.Index])
                    columnIndexes.Add(column.Index);
            }
        }

        /// <summary>把当前预览另存为 xlsx。先让用户选路径，再写入临时工作簿。</summary>
        /// <remarks>
        /// 不重新查询，也不修改当前工作簿。用户取消对话框时不导出。
        /// 路径指向当前工作簿时只提示，不覆盖。失败弹出错误详情，不提供「问 AI」。
        /// 成功后弹出与状态栏相同的一句，必须点确定。取消不弹这层框。
        /// </remarks>
        private void ExportPreview()
        {
            QueryTab tab = ActiveTab;
            if (tab == null || !tab.HasPreview || tab.Preview == null || tab.Preview.Columns.Count == 0) return;

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "导出为 Excel";
                dialog.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                dialog.DefaultExt = "xlsx";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                dialog.FileName = PreviewExportFileName(tab.Name);
                string directory = CurrentWorkbookDirectory();
                if (!string.IsNullOrEmpty(directory))
                    dialog.InitialDirectory = directory;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    _status.Text = "已取消导出。";
                    return;
                }

                string path = dialog.FileName;
                if (IsCurrentWorkbookPath(path))
                {
                    MessageBox.Show(this, "不能把预览导出到当前已打开的工作簿。", "导出为 Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                BufferedQueryResult result = tab.Preview;
                string sheetName = PreviewExportSheetName(tab.Name);
                try
                {
                    long rows = _operationRunner.Run(
                        this,
                        "导出 Excel",
                        true,
                        (progress, token) => Task.FromResult(WritePreviewFile(path, sheetName, result, progress, token)));
                    int previewLimit = _settings.Get(CoreSettings.MaxPreviewRows);
                    ShowOutcome(
                        "导出为 Excel",
                        result.IsTruncated
                            ? string.Format(
                                "已导出到“{0}”，共 {1:N0} 行，已达上限 {2:N0} 行并截断。可在设置中修改“最大预览行数”。",
                                path,
                                rows,
                                previewLimit)
                            : string.Format("已导出到“{0}”，共 {1:N0} 行。", path, rows),
                        result.IsTruncated);
                }
                catch (OperationCanceledException)
                {
                    _status.Text = "导出已取消。";
                }
                catch (Exception exception)
                {
                    ShowError(exception);
                }
            }
        }

        /// <summary>把预览写入一个新建的工作簿并另存为 xlsx，然后关掉它。</summary>
        /// <param name="path">用户选择的 xlsx 路径。</param>
        /// <param name="sheetName">已经整理过的工作表名。</param>
        /// <param name="result">当前页已经查出的预览。调用方保证有列。</param>
        /// <param name="progress">进度窗。写入器按块报告。</param>
        /// <param name="cancellationToken">用户取消时在保存前停下来。已创建的临时簿仍会关掉。</param>
        /// <returns>写入的数据行数，不含标题。</returns>
        /// <remarks>
        /// 必须在 Excel 界面线程调用。不释放宿主 Application，也不释放用户原来的工作簿。
        /// 不把临时簿的窗口藏起来。藏起来再另存会写出打不开的文件。重画保持关闭，避免这张表闪一下。
        /// 另存格式用整数 51，也就是 xlsx。传枚举时宿主会忽略，扩展名虽是 xlsx，内容仍是旧的 xls。
        /// 关掉临时簿之后核对文件头。不是有效的 xlsx 就删掉并抛出错误。取消或失败时也关掉临时簿。
        /// </remarks>
        private long WritePreviewFile(
            string path,
            string sheetName,
            BufferedQueryResult result,
            IProgress<OperationProgress> progress,
            System.Threading.CancellationToken cancellationToken)
        {
            ExcelInterop.Application application = _workbook.Application;
            ExcelInterop.Workbook original = null;
            ExcelInterop.Workbook temp = null;
            bool restoreScreen = false;
            bool previousScreen = true;
            bool restoreAlerts = false;
            bool previousAlerts = true;
            try
            {
                try
                {
                    original = application.ActiveWorkbook;
                }
                catch (COMException)
                {
                    // 读不到活动工作簿时仍继续导出，结束时不切换回去。
                }
                catch (InvalidComObjectException)
                {
                }

                try
                {
                    previousScreen = application.ScreenUpdating;
                    application.ScreenUpdating = false;
                    restoreScreen = true;
                }
                catch (COMException)
                {
                    // 关不掉重画时仍继续。新建的工作簿可能会闪一下。
                }
                catch (InvalidComObjectException)
                {
                }

                try
                {
                    previousAlerts = application.DisplayAlerts;
                    application.DisplayAlerts = false;
                    restoreAlerts = true;
                }
                catch (COMException)
                {
                    // 关不掉提示时不删除多余工作表，避免确认框卡在进度窗下面。
                }
                catch (InvalidComObjectException)
                {
                }

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Preparing,
                    TargetSheetName = sheetName,
                    Message = "正在创建临时工作簿…",
                    IsIndeterminate = true
                });
                temp = application.Workbooks.Add();
                if (restoreAlerts)
                    KeepOnlyFirstSheet(temp);
                RenameFirstSheet(temp, sheetName);
                long written = _writer.WriteResult(
                    temp,
                    sheetName,
                    result,
                    progress,
                    cancellationToken,
                    Guid.NewGuid().ToString("N"));
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Writing,
                    TargetSheetName = sheetName,
                    Message = "预览数据已写入，正在保存 xlsx 文件…",
                    Percent = 100,
                    IsIndeterminate = false
                });
                temp.SaveAs(
                    Filename: path,
                    FileFormat: 51,
                    AccessMode: ExcelInterop.XlSaveAsAccessMode.xlNoChange);
                try
                {
                    temp.Saved = true;
                }
                catch (COMException)
                {
                    // 标不成已保存时仍关闭。Close(false) 不会把已写好的文件再改掉。
                }
                catch (InvalidComObjectException)
                {
                }

                CloseWorkbook(temp);
                temp = null;
                EnsureXlsxFile(path);
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Completed,
                    TargetSheetName = sheetName,
                    Message = "xlsx 文件已保存并通过格式校验。",
                    Percent = 100,
                    IsIndeterminate = false
                });
                return written;
            }
            finally
            {
                if (temp != null)
                    CloseWorkbook(temp);
                ActivateWorkbook(original);
                if (restoreAlerts)
                {
                    try
                    {
                        application.DisplayAlerts = previousAlerts;
                    }
                    catch (COMException)
                    {
                        // 恢复提示失败不能盖住原来的导出异常。
                    }
                    catch (InvalidComObjectException)
                    {
                    }
                }

                if (restoreScreen)
                {
                    try
                    {
                        application.ScreenUpdating = previousScreen;
                    }
                    catch (COMException)
                    {
                        // 恢复重画失败不能盖住原来的导出异常。
                    }
                    catch (InvalidComObjectException)
                    {
                    }
                }

                if (original != null && !ReferenceEquals(original, _workbook))
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(original);
                    }
                    catch (InvalidComObjectException)
                    {
                        // 引用已经释放时不再处理。
                    }
                }
            }
        }

        /// <summary>确认另存结果是非空的 xlsx。xlsx 是 zip，文件头必须是 PK。</summary>
        /// <param name="path">刚关掉的临时簿写下的路径。</param>
        /// <remarks>不是有效文件时删掉它再抛出，避免状态栏把坏文件报成导出成功。</remarks>
        /// <exception cref="InvalidOperationException">文件为空、读不到，或文件头不是 PK。</exception>
        private static void EnsureXlsxFile(string path)
        {
            byte[] header = new byte[2];
            int read = 0;
            bool readable = false;
            try
            {
                using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    readable = true;
                    if (stream.Length > 0)
                        read = stream.Read(header, 0, 2);
                }
            }
            catch (IOException)
            {
                // 读不到就按坏文件处理，下面删掉并抛出。
            }
            catch (UnauthorizedAccessException)
            {
            }

            if (readable && read == 2 && header[0] == (byte)'P' && header[1] == (byte)'K')
                return;

            DeletePreviewExport(path);
            throw new InvalidOperationException("导出的文件不是有效的 xlsx。");
        }

        /// <summary>删掉核对失败的导出文件。</summary>
        /// <param name="path">要删除的路径。</param>
        /// <remarks>文件正被占用时删不掉，仍由调用方抛出原来的错误。</remarks>
        private static void DeletePreviewExport(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (IOException)
            {
                // 删不掉坏文件时仍抛出原来的错误。
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        /// <summary>只留第一张工作表，避免新建工作簿按 Excel 选项带出多余空表。</summary>
        /// <param name="workbook">临时工作簿。调用前须已关掉提示，否则删除可能弹出确认。</param>
        private static void KeepOnlyFirstSheet(ExcelInterop.Workbook workbook)
        {
            while (workbook.Worksheets.Count > 1)
            {
                ExcelInterop.Worksheet extra = workbook.Worksheets[workbook.Worksheets.Count] as ExcelInterop.Worksheet;
                if (extra == null) break;
                try
                {
                    extra.Delete();
                }
                finally
                {
                    Marshal.FinalReleaseComObject(extra);
                }
            }
        }

        /// <summary>把第一张工作表改成目标名，使写入器复用它而不是再追加一张。</summary>
        /// <param name="workbook">临时工作簿。</param>
        /// <param name="sheetName">合法的工作表名。</param>
        private static void RenameFirstSheet(ExcelInterop.Workbook workbook, string sheetName)
        {
            ExcelInterop.Worksheet first = workbook.Worksheets[1] as ExcelInterop.Worksheet;
            if (first == null)
                throw new InvalidOperationException("新建的工作簿没有工作表。");
            try
            {
                if (!string.Equals(first.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    first.Name = sheetName;
            }
            finally
            {
                Marshal.FinalReleaseComObject(first);
            }
        }

        /// <summary>关掉临时工作簿且不保存，并释放 COM 引用。</summary>
        /// <param name="workbook">要关掉的工作簿。调用后不得再使用。</param>
        private static void CloseWorkbook(ExcelInterop.Workbook workbook)
        {
            if (workbook == null) return;
            try
            {
                workbook.Close(false);
            }
            catch (COMException)
            {
                // 已经关掉或宿主拒绝时仍释放引用，避免临时簿留在进程里。
            }
            catch (InvalidComObjectException)
            {
            }
            finally
            {
                Marshal.FinalReleaseComObject(workbook);
            }
        }

        /// <summary>回到导出前的活动工作簿。</summary>
        /// <param name="workbook">导出前的活动工作簿。为空或已关掉时不做。不释放。</param>
        private static void ActivateWorkbook(ExcelInterop.Workbook workbook)
        {
            if (workbook == null) return;
            try
            {
                workbook.Activate();
            }
            catch (COMException)
            {
                // 原来的工作簿已经关掉时不打断导出结果。
            }
            catch (InvalidComObjectException)
            {
            }
        }

        /// <summary>读取当前工作簿所在目录，用作保存对话框的起始位置。</summary>
        /// <returns>已保存工作簿的目录。未保存或读不到时为空。</returns>
        private string CurrentWorkbookDirectory()
        {
            try
            {
                string fullName = _workbook.FullName;
                if (string.IsNullOrWhiteSpace(fullName) || !Path.IsPathRooted(fullName)) return null;
                return Path.GetDirectoryName(fullName);
            }
            catch (COMException)
            {
                // 读不到路径时用对话框自己的目录，不打断导出。
                return null;
            }
            catch (InvalidComObjectException)
            {
                return null;
            }
        }

        /// <summary>判断路径是不是当前正在编辑的工作簿，避免另存覆盖它。</summary>
        /// <param name="path">用户选择的文件路径。</param>
        /// <returns>指向当前工作簿文件时为 true。未保存的工作簿没有路径，返回 false。</returns>
        private bool IsCurrentWorkbookPath(string path)
        {
            try
            {
                string fullName = _workbook.FullName;
                if (string.IsNullOrWhiteSpace(fullName) || !Path.IsPathRooted(fullName)) return false;
                return string.Equals(Path.GetFullPath(fullName), Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase);
            }
            catch (COMException)
            {
                return false;
            }
            catch (InvalidComObjectException)
            {
                return false;
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (NotSupportedException)
            {
                return false;
            }
            catch (IOException)
            {
                return false;
            }
        }

        /// <summary>把页签名称整理成可用的文件名。去掉系统不允许的字符。</summary>
        /// <param name="name">页签名称，不含未保存标记。</param>
        /// <returns>不含扩展名的文件名。</returns>
        private static string PreviewExportFileName(string name)
        {
            string value = string.IsNullOrWhiteSpace(name) ? "预览" : name.Trim();
            foreach (char invalid in Path.GetInvalidFileNameChars())
                value = value.Replace(invalid.ToString(), string.Empty);
            value = value.Trim();
            if (value.Length == 0) return "预览";
            if (value.Length > 180) value = value.Substring(0, 180).Trim();
            return value.Length == 0 ? "预览" : value;
        }

        /// <summary>把页签名称整理成工作表名。去掉 Excel 不允许的字符，最长 31 个字符。</summary>
        /// <param name="name">页签名称。</param>
        /// <returns>可直接用作工作表名。空名称或 SQL 时返回「预览」。</returns>
        private static string PreviewExportSheetName(string name)
        {
            string value = name ?? string.Empty;
            char[] invalid = { ':', '\\', '/', '?', '*', '[', ']' };
            for (int index = 0; index < invalid.Length; index++)
                value = value.Replace(invalid[index].ToString(), string.Empty);
            value = value.Trim().Trim('\'');
            if (value.Length > 31)
                value = value.Substring(0, 31).Trim().Trim('\'');
            if (value.Length == 0 || string.Equals(value, SqlSheetTaskReader.SqlSheetName, StringComparison.OrdinalIgnoreCase))
                return "预览";
            return value;
        }

        /// <summary>读取格子里用户看见的文字。空值是空串。</summary>
        /// <param name="cell">预览表中的格子。</param>
        /// <returns>显示文字。</returns>
        private static string PreviewCellText(DataGridViewCell cell)
        {
            object value = cell.FormattedValue;
            if (value == null || value is DBNull) return string.Empty;
            return Convert.ToString(value) ?? string.Empty;
        }

        /// <summary>向制表符文本和 HTML 表各追加一行。</summary>
        /// <param name="text">制表符文本。列之间是制表符，行之间是换行。</param>
        /// <param name="html">HTML 表身，不含 html 标记。</param>
        /// <param name="cells">这一行的各格文字。</param>
        /// <param name="first">为 true 时文本前面不加换行。</param>
        private static void AppendClipboardRow(StringBuilder text, StringBuilder html, string[] cells, bool first)
        {
            if (!first)
                text.Append("\r\n");
            html.Append("<tr>");
            for (int index = 0; index < cells.Length; index++)
            {
                if (index > 0)
                    text.Append('\t');
                string value = cells[index] ?? string.Empty;
                text.Append(EscapeTsv(value));
                html.Append("<td>");
                html.Append(EncodeHtmlCell(value));
                html.Append("</td>");
            }

            html.Append("</tr>");
        }

        /// <summary>把含制表符或换行的格子用引号包起来，供纯文本粘贴。</summary>
        /// <param name="value">格子文字。</param>
        /// <returns>可放进制表符行的文字。</returns>
        private static string EscapeTsv(string value)
        {
            if (value.IndexOfAny(new[] { '\t', '\r', '\n', '"' }) < 0)
                return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        /// <summary>把格子文字转成 HTML 单元格内容。换行变成 br，贴进 Excel 时仍在同一格。</summary>
        /// <param name="value">格子文字。</param>
        /// <returns>已编码的 HTML。</returns>
        private static string EncodeHtmlCell(string value)
        {
            return WebUtility.HtmlEncode(value).Replace("\r\n", "<br>").Replace("\n", "<br>").Replace("\r", "<br>");
        }

        /// <summary>同时放入 Unicode 文本和 UTF-8 的 HTML。Excel 优先用 HTML，才能按列分开并保留类型行。</summary>
        /// <param name="text">制表符文本。给不读 HTML 的程序用。</param>
        /// <param name="htmlFragment">table 片段，不含 CF_HTML 头。</param>
        /// <remarks>
        /// 必须在界面线程调用。剪贴板被占用时由调用方弹出错误详情。
        /// HTML 必须以 UTF-8 字节放入，并写明字符集。按字符串放入时会按系统 ANSI 编码，中文贴进 Excel 就是乱码。
        /// </remarks>
        private static void SetPreviewClipboard(string text, string htmlFragment)
        {
            DataObject data = new DataObject();
            data.SetData(DataFormats.UnicodeText, text);
            byte[] htmlBytes = new UTF8Encoding(false).GetBytes(BuildHtmlClipboard(htmlFragment));
            using (MemoryStream html = new MemoryStream(htmlBytes))
            {
                data.SetData(DataFormats.Html, false, html);
                Clipboard.SetDataObject(data, true);
            }
        }

        /// <summary>按 CF_HTML 拼出剪贴板需要的描述头。偏移按 UTF-8 字节计算。</summary>
        /// <param name="fragment">StartFragment 与 EndFragment 之间的 HTML。</param>
        /// <returns>带描述头的完整 HTML 剪贴板文本。正文声明字符集为 UTF-8。</returns>
        private static string BuildHtmlClipboard(string fragment)
        {
            const string header =
                "Version:0.9\r\nStartHTML:{0:0000000000}\r\nEndHTML:{1:0000000000}\r\nStartFragment:{2:0000000000}\r\nEndFragment:{3:0000000000}\r\n";
            string start = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html;charset=utf-8\"></head><body>\r\n<!--StartFragment-->";
            string end = "<!--EndFragment-->\r\n</body></html>";
            string prefix = string.Format(header, 0, 0, 0, 0);
            Encoding utf8 = new UTF8Encoding(false);
            int startHtml = utf8.GetByteCount(prefix);
            int startFragment = startHtml + utf8.GetByteCount(start);
            int endFragment = startFragment + utf8.GetByteCount(fragment ?? string.Empty);
            int endHtml = endFragment + utf8.GetByteCount(end);
            prefix = string.Format(header, startHtml, endHtml, startFragment, endFragment);
            return prefix + start + (fragment ?? string.Empty) + end;
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
            _editorSurface.MouseUp -= EmptyEditorMouseUp;
            _emptyQueryButton.MouseUp -= EmptyEditorMouseUp;
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
            _preview.MouseDown -= PreviewMouseDown;
            _preview.KeyDown -= PreviewGridKeyDown;
            _previewMenu.Opening -= PreviewMenuOpening;
            _queryProfileMenu.Dispose();
            _queryTabMenu.Dispose();
            _previewMenu.Dispose();
            foreach (Image icon in _actionIcons)
                icon.Dispose();
            _actionIcons.Clear();
            _queryTabs.ImageList = null;
            _queryTabIcons.Dispose();
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

        /// <summary>把完成说明写入状态栏，并弹出必须点确定的提示。</summary>
        /// <param name="title">提示框标题。</param>
        /// <param name="message">与状态栏相同的一句说明。</param>
        /// <param name="truncated">结果被截断时用警告图标，否则用信息图标。</param>
        /// <remarks>在界面线程上以模态方式显示。调用方只在正常结束时调用。取消和失败不走这里。</remarks>
        private void ShowOutcome(string title, string message, bool truncated)
        {
            _status.Text = message;
            MessageBox.Show(
                this,
                message,
                title,
                MessageBoxButtons.OK,
                truncated ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        private void ShowError(Exception exception, string sqlForAi = null)
        {
            _status.Text = exception.Message;
            bool ask = ExceptionDetailForm.Show(this, exception, sqlForAi != null);
            if (ask)
                OpenAiErrorChat(sqlForAi, exception.Message);
        }

        private void UpdateConnectionMenuState()
        {
            bool hasSelection = _connectionTree.HasSelectedConnection;
            _editConnectionMenuItem.Enabled = hasSelection;
            _deleteConnectionMenuItem.Enabled = hasSelection;
            _setCurrentConnectionMenuItem.Enabled = hasSelection;
        }

        /// <summary>创建只显示图标的操作按钮。文字写在悬浮说明里。图像在窗体关闭时释放。</summary>
        /// <param name="iconFileName">Resources 目录下的图标文件名。</param>
        /// <param name="tip">概述换行详述的悬浮说明。</param>
        /// <param name="iconSize">图标边长。说明行的预览、写入和保存用 28。</param>
        /// <returns>无文字的图标按钮。</returns>
        private Button IconAction(string iconFileName, string tip, int iconSize)
        {
            Image icon = DatabaseTreeImageCatalog.LoadActionIcon(iconFileName, iconSize);
            _actionIcons.Add(icon);
            Button button = new Button
            {
                Text = string.Empty,
                Image = icon,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(4, 2, 4, 2),
                Margin = new Padding(0, 0, 3, 0),
                UseVisualStyleBackColor = true
            };
            _profileTip.SetToolTip(button, tip);
            return button;
        }

        private static Panel SectionPanel(string title, Control content)
        {
            Panel panel = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Control };
            panel.Controls.Add(content);
            panel.Controls.Add(SectionHeader(title));
            return LightBorder.Wrap(panel);
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

            editor.Visible = false;
            _editorSurface.Controls.Add(editor);
            TabPage page = new TabPage
            {
                Tag = tab,
                Padding = Padding.Empty,
                UseVisualStyleBackColor = true,
                ImageIndex = 0
            };
            tab.Page = page;
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

        /// <summary>从页签条移除并释放编辑器。最后一个页签关掉后显示新建按钮。</summary>
        /// <param name="tab">要移除的页签。编辑器不在页里，这里单独释放。</param>
        private void RemoveTab(QueryTab tab)
        {
            if (tab == null || tab.Page == null) return;
            SqlEditorControl editor = tab.Editor;
            editor.SqlTextChanged -= DocumentSqlTextChanged;
            editor.ExecuteSelectionRequested -= DocumentExecuteSelectionRequested;
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

            if (editor.Parent != null)
                editor.Parent.Controls.Remove(editor);
            editor.Dispose();
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
            ShowOnly(tab);
            PlaceNewQueryButton();
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
            RefreshAiReference();
        }

        /// <summary>只显示当前页的编辑器。没有页签时改显示中央的新建按钮。</summary>
        /// <param name="tab">要显示的页。为空表示编辑区是空的。</param>
        private void ShowOnly(QueryTab tab)
        {
            foreach (TabPage page in _queryTabs.TabPages)
            {
                QueryTab item = page.Tag as QueryTab;
                if (item == null || item.Editor == null || item.Editor.IsDisposed) continue;
                item.Editor.Visible = item == tab;
            }

            _emptyQueryButton.Visible = tab == null;
            if (tab == null)
            {
                _emptyQueryButton.BringToFront();
                CenterEmptyQueryButton();
                return;
            }

            if (tab.Editor.Parent != _editorSurface)
                _editorSurface.Controls.Add(tab.Editor);
            tab.Editor.Dock = DockStyle.Fill;
            tab.Editor.BringToFront();
        }

        /// <summary>把「请新建查询」放在编辑区正中。按钮本身没有边框。</summary>
        private void CenterEmptyQueryButton()
        {
            if (_emptyQueryButton == null || _editorSurface == null) return;
            Size size = _emptyQueryButton.PreferredSize;
            if (_emptyQueryButton.Width > 0) size.Width = _emptyQueryButton.Width;
            if (_emptyQueryButton.Height > 0) size.Height = _emptyQueryButton.Height;
            _emptyQueryButton.Location = new Point(
                Math.Max(0, (_editorSurface.ClientSize.Width - size.Width) / 2),
                Math.Max(0, (_editorSurface.ClientSize.Height - size.Height) / 2));
        }

        /// <summary>藏起页签控件，在编辑区显示「请新建查询」，并折叠预览。页签条和新建图标仍留着。</summary>
        private void ShowEmptyEditor()
        {
            _queryTabs.Visible = false;
            ShowOnly(null);
            PlaceNewQueryButton();
            ClearPreviewGrid();
            _resultsSplit.Panel2Collapsed = true;
            SyncProfileSelection(null);
            _status.Text = "请新建查询。";
            RefreshAiReference();
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

        /// <summary>把方案名、未保存标记和关闭空位写进页签文字。系统按这段文字计算宽度，标记另画。</summary>
        /// <param name="tab">要更新标题的页签。</param>
        private void RefreshTabTitle(QueryTab tab)
        {
            if (tab == null || tab.Page == null || tab.Page.IsDisposed) return;
            string caption = CloseMark.Caption(tab.TabTitle, _queryTabs.Font);
            if (string.Equals(tab.Page.Text, caption, StringComparison.Ordinal)) return;
            tab.Page.Text = caption;
            PlaceNewQueryButton();
        }

        /// <summary>把新建按钮放在最后一个页签右侧。页签放不下时停在条的右端，没有页签时停在左端。句柄或条宽还没准备好时不改宽度，避免重开窗体后把页签挤没。条高按页签实际高度留出整行标题。</summary>
        private void PlaceNewQueryButton()
        {
            if (_placingQueryButton || _queryTabBar == null || _queryTabHost == null || _addQueryButton == null) return;
            int bar = _queryTabBar.ClientSize.Width;
            if (bar <= 0) return;
            int button = Math.Max(28, _addQueryButton.Width);
            _placingQueryButton = true;
            try
            {
                if (!_queryTabs.IsHandleCreated)
                    return;
                if (!_queryTabs.Visible || _queryTabs.TabCount == 0)
                {
                    _queryTabHost.Width = button;
                    return;
                }

                if (_queryTabHost.Width != bar)
                    _queryTabHost.Width = bar;
                _queryTabHost.PerformLayout();
                int right = 0;
                int tabHeight = 0;
                for (int index = 0; index < _queryTabs.TabCount; index++)
                {
                    Rectangle rect = _queryTabs.GetTabRect(index);
                    if (rect.Right > right) right = rect.Right;
                    if (rect.Height > tabHeight) tabHeight = rect.Height;
                }

                int textHeight = TextRenderer.MeasureText("页签", _queryTabs.Font).Height + 8;
                int barHeight = Math.Max(28, Math.Max(tabHeight, textHeight));
                if (_queryTabBar.Height != barHeight)
                    _queryTabBar.Height = barHeight;

                const int slack = 12;
                int natural = right + slack + button;
                bool fits = natural < bar;
                int desired = fits ? Math.Max(button, natural) : bar;
                if (_queryTabHost.Width != desired)
                    _queryTabHost.Width = desired;
                _queryTabHost.PerformLayout();
                SetTabScrollButtons(fits);
            }
            finally
            {
                _placingQueryButton = false;
            }
        }

        /// <summary>页签放得下时藏起左右箭头，避免箭头盖住标题。放不下时留着，供左右滚动。</summary>
        /// <param name="hide">为 true 时隐藏箭头。</param>
        private void SetTabScrollButtons(bool hide)
        {
            if (_queryTabs == null || !_queryTabs.IsHandleCreated) return;
            IntPtr upDown = FindWindowEx(_queryTabs.Handle, IntPtr.Zero, "msctls_updown32", null);
            if (upDown == IntPtr.Zero) return;
            ShowWindow(upDown, hide ? 0 : 5);
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string className, string window);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr window, int command);

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
                RefreshAiReference();
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
            if (!CloseMark.Hit(_queryTabs.GetTabRect(index), e.Location)) return;
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
            Control source = sender as Control ?? _editorSurface;
            _queryTabMenu.Show(source, e.Location);
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

        /// <summary>写入 SQL 之后的美化结果。失败时错误详情已经弹出，调用方不再提示写入成功。</summary>
        private enum SavedSqlBeautify
        {
            /// <summary>未开启保存后美化，或没有可美化的列。</summary>
            Skipped,

            /// <summary>整张 SQL 页已美化。</summary>
            Done,

            /// <summary>用户取消了美化。结果和 SQL 已经写入。</summary>
            Cancelled,

            /// <summary>美化抛错。错误详情已显示。</summary>
            Failed
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

            /// <summary>页签条上的那一页。编辑器放在页签条下面，不放在这一页里。移除页签后置空，避免再改已释放的标题。</summary>
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

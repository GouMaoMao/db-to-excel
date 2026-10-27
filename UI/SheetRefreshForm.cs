using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CheckBoxState = System.Windows.Forms.VisualStyles.CheckBoxState;
using DB2Sheet.Contracts;
using DB2Sheet.Excel;
using DB2Sheet.Models;
using DB2Sheet.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.UI
{
    /// <summary>扫描当前工作簿的 SQL 页，勾选任务后批量刷新目标工作表。</summary>
    /// <remarks>
    /// 左栏与 SQL 查询窗体共用连接树和当前连接。右栏列出任务。
    /// 特殊参数按键分列，改完后写回 SQL 页首行，不改 SQL 正文。
    /// 点击摘要跳到 SQL 命令首行，点击目标表名打开该表。任务行右键可复制该行结果说明。
    /// 单击勾选框切换该行并记为连续勾选的起点。按住 Shift 点击时，把起点那次的勾选结果写到起点与该行之间，且不跳转。
    /// 「格式」菜单美化 SQL 页，或打开参数设置的「sql页格式」。美化走进度窗，按列推进，可取消。
    /// 窗体尺寸沿用会话记忆；左右比例只在用户拖动分割条时写入。
    /// 第一次显示保持透明，任务在这时读完。重名和读取错误等窗体显示后再提示。
    /// </remarks>
    public sealed class SheetRefreshForm : AppForm
    {
        private const int LeftMinWidth = 160;
        private const int RightMinWidth = 480;
        private static readonly Color CheckedRowBackColor = Color.FromArgb(232, 242, 252);
        private static readonly Color HeaderBackColor = Color.FromArgb(217, 217, 217);

        private readonly IOperationRunner _operationRunner;
        private readonly ISettingsStore _settings;
        private readonly ISqlSheetTaskReader _taskReader;
        private readonly IBatchRefreshService _refreshService;
        private readonly ExcelInterop.Workbook _workbook;
        private readonly DatabaseConnectionTree _connectionTree;
        private readonly SplitContainer _workspace;
        private readonly DataGridView _tasks;
        private readonly Label _hint;
        private readonly Label _status;
        private readonly Font _linkFont;
        private readonly Font _hintFont;
        private readonly List<Image> _actionIcons = new List<Image>();
        private readonly List<string> _optionKeys = new List<string>();
        private readonly int _checkColumn;
        private readonly int _targetColumn;
        private readonly int _summaryColumn;
        private readonly int _statusColumn;
        private readonly int _optionColumnStart;
        private readonly List<TaskRow> _rows = new List<TaskRow>();
        private readonly ContextMenuStrip _taskMenu;
        private readonly ToolStripMenuItem _copyResultItem;
        private readonly Action _openSqlFormatSettings;
        private TaskRow _contextRow;
        private bool _applyingSplit;
        private bool _userDraggingSplit;
        /// <summary>激活后重扫的抑制层数。对话框和批量写入可以嵌套，减到 0 才允许重扫。</summary>
        private int _activateSuppressDepth;
        private bool _scanning;
        private bool _binding;

        /// <summary>上一次单击勾选框的行。按住 Shift 时从这里延续到所点的行。重新绑表后作废。</summary>
        private int _rangeAnchor = -1;

        /// <summary>起点那次点击之后的勾选结果。Shift 范围整段都写成这个值。</summary>
        private bool _rangeAnchorChecked;

        /// <summary>这次按下已经按住 Shift。点击处理完之前用来拦住跳转，并识别勾选框是否先自己翻转。</summary>
        private bool _shiftClick;

        /// <summary>按住 Shift 时已经自己翻转过的勾选框所在行。只跳过这一行，避免再翻一次。</summary>
        private int _shiftAdjustedRow = -1;

        /// <summary>Shift 点击结束参数格的编辑时不把内容写回 SQL 页。</summary>
        private bool _discardEdit;

        /// <summary>首次显示还没画出来。这段时间的激活不读 SQL 页。</summary>
        private bool _opening;

        /// <summary>首次读取正在把重名和错误留到窗体显示之后。</summary>
        private bool _deferScanNotices;

        /// <summary>首次读取时延后的重名警告和错误详情，窗体显示后再按原顺序弹出。</summary>
        private readonly List<Action> _deferredScanNotices = new List<Action>();

        /// <summary>显示后紧接着的那次激活不再重扫，避免同一次打开读两遍 SQL 页。若没有这次激活，下一条界面消息会清掉标记。</summary>
        private bool _skipActivateScan;
        private string _scanSummary = "正在读取 SQL 页。";

        /// <summary>创建 Sheet 批量刷新窗体。</summary>
        /// <param name="connections">连接方案仓储。</param>
        /// <param name="providers">数据源提供程序注册表。</param>
        /// <param name="executionService">连接测试和查询执行服务。</param>
        /// <param name="operationRunner">批次进度窗体运行器。</param>
        /// <param name="settings">批次模式、并发数、分栏比例和窗体尺寸。</param>
        /// <param name="taskReader">SQL 页任务读取、首行写回和跳转。</param>
        /// <param name="refreshService">批量刷新编排服务。</param>
        /// <param name="workbook">当前操作的 Excel 工作簿。</param>
        /// <param name="jdbcEnvironment">本机 Java 路径，传给连接树里的编辑窗。</param>
        /// <param name="openJdbcEnvironment">打开 JDBC 环境窗体；为空时连接编辑不显示该入口。</param>
        /// <param name="openSqlFormatSettings">打开参数设置并停在「sql页格式」。为空时「格式设置」无动作。</param>
        public SheetRefreshForm(
            IConnectionProfileRepository connections,
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IOperationRunner operationRunner,
            ISettingsStore settings,
            ISqlSheetTaskReader taskReader,
            IBatchRefreshService refreshService,
            ExcelInterop.Workbook workbook,
            IJdbcEnvironmentStore jdbcEnvironment,
            Action openJdbcEnvironment = null,
            Action openSqlFormatSettings = null)
        {
            if (executionService == null) throw new ArgumentNullException(nameof(executionService));
            if (jdbcEnvironment == null) throw new ArgumentNullException(nameof(jdbcEnvironment));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _taskReader = taskReader ?? throw new ArgumentNullException(nameof(taskReader));
            _refreshService = refreshService ?? throw new ArgumentNullException(nameof(refreshService));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _openSqlFormatSettings = openSqlFormatSettings;

            InitializeWindow(
                AppPresentation.WindowTitle("Sheet 批量刷新"),
                1100,
                680,
                new Size(900, 520),
                _settings,
                "SheetRefresh");

            _connectionTree = new DatabaseConnectionTree(
                connections, providers, executionService, _operationRunner, _settings, jdbcEnvironment, openJdbcEnvironment);
            _linkFont = new Font(AppPresentation.DefaultFontName, AppPresentation.DefaultFontSize, FontStyle.Underline);
            _hintFont = new Font(AppPresentation.DefaultFontName, 8.25f, FontStyle.Regular);
            _hint = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = SystemColors.GrayText,
                Font = _hintFont,
                TextAlign = ContentAlignment.MiddleLeft
            };
            _status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Padding = new Padding(8, 6, 8, 0),
                ForeColor = SystemColors.GrayText,
                Text = _scanSummary
            };
            Button refreshButton = ButtonOf("批量刷新", "icon_writeexcel.png");
            refreshButton.Margin = new Padding(0, 4, 0, 0);

            _tasks = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                EnableHeadersVisualStyles = false,
                ShowCellToolTips = true,
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = HeaderBackColor,
                    ForeColor = SystemColors.ControlText,
                    SelectionBackColor = HeaderBackColor
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.White,
                    SelectionBackColor = SystemColors.Highlight,
                    SelectionForeColor = SystemColors.HighlightText
                }
            };
            _checkColumn = _tasks.Columns.Add(new DataGridViewCheckBoxColumn
            {
                HeaderText = string.Empty,
                Width = 36,
                ToolTipText = "全选或全不选"
            });
            _targetColumn = _tasks.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "目标表", Width = 180, ReadOnly = true });
            _summaryColumn = _tasks.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "SQL 摘要", Width = 280, ReadOnly = true });
            _statusColumn = _tasks.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", Width = 48, ReadOnly = true });
            _tasks.Columns[_summaryColumn].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _optionColumnStart = _tasks.Columns.Count;
            AddOptionColumn(SqlSheetHeader.StartCellKey);
            AddOptionColumn(SqlSheetHeader.ClearExtraColumnsKey);
            foreach (DataGridViewColumn column in _tasks.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            _copyResultItem = new ToolStripMenuItem("复制结果信息");
            _taskMenu = new ContextMenuStrip();
            _taskMenu.Items.Add(_copyResultItem);
            _tasks.ContextMenuStrip = _taskMenu;

            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 2, 0, 0)
            };
            actions.Controls.Add(refreshButton);
            Panel hintBar = new Panel { Dock = DockStyle.Top, Height = 72, Padding = new Padding(8, 4, 8, 4) };
            hintBar.Controls.Add(_hint);
            hintBar.Controls.Add(actions);

            Panel right = new Panel { Dock = DockStyle.Fill };
            right.Controls.Add(_tasks);
            right.Controls.Add(_status);
            right.Controls.Add(hintBar);

            _workspace = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6
            };
            _workspace.Panel1.Controls.Add(SectionPanel("数据库连接", _connectionTree));
            _workspace.Panel2.Controls.Add(SectionPanel("刷新任务", right));
            Controls.Add(_workspace);
            Controls.Add(BuildFormatMenu());

            refreshButton.Click += (sender, args) => RefreshTasks();
            _connectionTree.NoticeChanged += (sender, args) => _status.Text = _connectionTree.Notice;
            _tasks.CurrentCellDirtyStateChanged += TasksCurrentCellDirtyStateChanged;
            _tasks.CellValueChanged += TasksCellValueChanged;
            _tasks.CellEndEdit += TasksCellEndEdit;
            _tasks.CellBeginEdit += TasksCellBeginEdit;
            _tasks.CellPainting += TasksCellPainting;
            _tasks.ColumnHeaderMouseClick += TasksColumnHeaderMouseClick;
            _tasks.CellClick += TasksCellClick;
            _tasks.CellMouseDown += TasksCellMouseDown;
            _tasks.CellMouseMove += TasksCellMouseMove;
            _taskMenu.Opening += TaskMenuOpening;
            _copyResultItem.Click += CopyResultMessage;
            _tasks.MouseLeave += (sender, args) => _tasks.Cursor = Cursors.Default;
            _tasks.SelectionChanged += TasksSelectionChanged;
            FormClosed += SheetRefreshFormFormClosed;
            Activated += SheetRefreshFormActivated;
            _workspace.SizeChanged += WorkspaceSizeChanged;
            _workspace.SplitterMoving += (sender, args) => _userDraggingSplit = true;
            _workspace.SplitterMoved += (sender, args) => RememberSplit();
            UpdateModeLabel();
        }

        /// <summary>首次显示在最终布局完成前保持完全透明。</summary>
        protected override bool DeferFirstShow
        {
            get { return true; }
        }

        /// <summary>透明显示之前挡住激活，避免和准备阶段各读一遍 SQL 页。</summary>
        protected override void OnFirstShowOpening()
        {
            _opening = true;
        }

        /// <summary>窗体已经画出来之后，再弹出首次读取留下的重名警告和错误详情。</summary>
        /// <remarks>提示会进入自己的消息循环。打开标记保持到提示结束，避免对话框关掉时重扫。</remarks>
        protected override void OnFirstShowRevealed()
        {
            try
            {
                ShowDeferredFirstScanNotices();
            }
            finally
            {
                _opening = false;
                _skipActivateScan = true;
                if (IsHandleCreated && !IsDisposed)
                    BeginInvoke(new Action(ReleaseSkippedActivateScan));
            }
        }

        /// <summary>
        /// 窗口已经按最终宽度显示，但仍然完全透明。
        /// 在这里摆分栏并读完 SQL 页任务。显示之后再改分割条或逐行填表，会把上一帧留在屏幕上。
        /// </summary>
        protected override void PrepareFirstShow()
        {
            ApplySplit();
            _deferScanNotices = true;
            try
            {
                ScanTasks();
            }
            finally
            {
                _deferScanNotices = false;
            }
        }

        private void SheetRefreshFormActivated(object sender, EventArgs e)
        {
            if (_opening) return;
            if (_skipActivateScan)
            {
                _skipActivateScan = false;
                return;
            }
            if (_activateSuppressDepth > 0 || _scanning || IsDisposed) return;
            if (_tasks.IsCurrentCellInEditMode) return;
            UpdateModeLabel();
            ScanTasks();
        }

        private void ScanTasks()
        {
            if (_scanning) return;
            _scanning = true;
            try
            {
                ScanTasksCore();
            }
            finally
            {
                _scanning = false;
            }
        }

        private void ScanTasksCore()
        {
            Dictionary<int, bool> checks = new Dictionary<int, bool>();
            foreach (TaskRow row in _rows)
                checks[row.Task.SourceColumn] = row.Checked;

            try
            {
                IReadOnlyList<RefreshTaskDefinition> tasks = _taskReader.ReadTasks(_workbook);
                NormalizeHeaders(tasks);
                _rows.Clear();
                foreach (RefreshTaskDefinition task in tasks)
                {
                    bool isChecked = checks.TryGetValue(task.SourceColumn, out bool saved) && saved;
                    _rows.Add(new TaskRow(task, isChecked));
                }
                ApplyRowStates();
                BindRows();
                _scanSummary = _rows.Count == 0
                    ? "SQL 页中未发现有效任务。"
                    : string.Format("发现 {0} 个刷新任务。", _rows.Count);
                if (_tasks.SelectedRows.Count == 0) _status.Text = _scanSummary;
                ShowDuplicateWarning();
            }
            catch (Exception exception)
            {
                _rows.Clear();
                BindRows();
                _scanSummary = exception.Message;
                _status.Text = exception.Message;
                ShowOwnedError(exception);
            }
        }

        /// <summary>把解析时规范过的特殊参数写回首行，避免下次再读到旧写法。</summary>
        private void NormalizeHeaders(IReadOnlyList<RefreshTaskDefinition> tasks)
        {
            foreach (RefreshTaskDefinition task in tasks)
            {
                if (!task.OptionsNormalized || !string.IsNullOrEmpty(task.OptionsError)) continue;
                try
                {
                    _taskReader.WriteColumnHeader(_workbook, task.SourceColumn, HeaderText(task));
                }
                catch (Exception exception)
                {
                    ShowOwnedError(exception);
                    return;
                }
            }
        }

        /// <summary>弹出错误详情时先挡住激活事件，避免对话框关闭后立刻重扫并打断当前编辑。</summary>
        /// <remarks>首次读取时窗体还是透明的，只记下异常，等窗体显示后再弹。</remarks>
        private void ShowOwnedError(Exception exception)
        {
            if (_deferScanNotices)
            {
                Exception captured = exception;
                _deferredScanNotices.Add(() => ShowOwnedError(captured));
                return;
            }

            PushActivateSuppress();
            try
            {
                ExceptionDetailForm.Show(this, exception);
            }
            finally
            {
                PopActivateSuppress();
            }
        }

        /// <summary>增加一层激活重扫抑制。与 <see cref="PopActivateSuppress"/> 成对使用。</summary>
        private void PushActivateSuppress()
        {
            _activateSuppressDepth++;
        }

        /// <summary>去掉一层激活重扫抑制。层数已经为 0 时不再减少。</summary>
        private void PopActivateSuppress()
        {
            if (_activateSuppressDepth > 0) _activateSuppressDepth--;
        }

        /// <summary>清掉“显示后跳过一次激活重扫”的标记。窗体已经释放时什么也不做。</summary>
        private void ReleaseSkippedActivateScan()
        {
            if (IsDisposed) return;
            _skipActivateScan = false;
        }

        /// <summary>按首次读取时的顺序弹出延后的重名警告和错误详情。</summary>
        private void ShowDeferredFirstScanNotices()
        {
            if (_deferredScanNotices.Count == 0) return;
            Action[] notices = _deferredScanNotices.ToArray();
            _deferredScanNotices.Clear();
            foreach (Action notice in notices)
                notice();
        }

        /// <summary>目标表重名时提示。首次读取先记下，等窗体显示后再弹。</summary>
        private void ShowDuplicateWarning()
        {
            string message = SqlSheetHeader.DescribeDuplicateTargets(_rows.Select(item => item.Task).ToList());
            if (string.IsNullOrEmpty(message)) return;
            if (_deferScanNotices)
            {
                _deferredScanNotices.Add(ShowDuplicateWarning);
                return;
            }

            PushActivateSuppress();
            try
            {
                MessageBox.Show(this, message, "批量刷新", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                PopActivateSuppress();
            }
        }

        private void ApplyRowStates()
        {
            Dictionary<string, string> duplicateMessages = DuplicateMessages();
            foreach (TaskRow row in _rows)
            {
                row.DuplicateMessage = duplicateMessages.TryGetValue(row.Task.TargetSheetName ?? string.Empty, out string message)
                    ? message
                    : string.Empty;
                ApplyCheckState(row);
            }
        }

        private Dictionary<string, string> DuplicateMessages()
        {
            Dictionary<string, string> messages = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string text = SqlSheetHeader.DescribeDuplicateTargets(_rows.Select(item => item.Task).ToList());
            if (string.IsNullOrEmpty(text)) return messages;
            foreach (string line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                foreach (TaskRow row in _rows)
                {
                    if (line.IndexOf("“" + row.Task.TargetSheetName + "”", StringComparison.OrdinalIgnoreCase) >= 0)
                        messages[row.Task.TargetSheetName] = line;
                }
            }
            return messages;
        }

        private static void ApplyCheckState(TaskRow row)
        {
            if (!string.IsNullOrEmpty(row.DuplicateMessage))
            {
                row.State = TaskVisualState.Error;
                row.Message = row.DuplicateMessage;
                return;
            }
            if (!string.IsNullOrEmpty(row.Task.OptionsError))
            {
                row.State = TaskVisualState.Error;
                row.Message = row.Task.OptionsError;
                return;
            }
            if (!row.Checked)
            {
                row.State = TaskVisualState.None;
                row.Message = "未选择";
                return;
            }
            if (row.State == TaskVisualState.Success || row.State == TaskVisualState.Truncated) return;
            row.State = TaskVisualState.Pending;
            row.Message = "待刷新";
        }

        private void BindRows()
        {
            _rangeAnchor = -1;
            _binding = true;
            try
            {
                _tasks.Rows.Clear();
                EnsureOptionColumns();
                foreach (TaskRow row in _rows)
                {
                    int index = _tasks.Rows.Add();
                    DataGridViewRow gridRow = _tasks.Rows[index];
                    gridRow.Tag = row;
                    gridRow.Cells[_checkColumn].Value = row.Checked;
                    DecorateRow(gridRow, row);
                }
            }
            finally
            {
                _binding = false;
            }
            InvalidateHeaderCheck();
            _tasks.InvalidateColumn(_statusColumn);
        }

        /// <summary>保证已知参数两列始终存在，其他键按第一次出现的顺序各占一列。</summary>
        private void EnsureOptionColumns()
        {
            List<string> keys = new List<string>
            {
                SqlSheetHeader.StartCellKey,
                SqlSheetHeader.ClearExtraColumnsKey
            };
            foreach (TaskRow row in _rows)
            {
                foreach (SqlSheetOption option in SqlSheetHeader.SplitOptions(row.Task.OptionsText))
                {
                    if (option.Key.Length == 0 || ContainsKey(keys, option.Key)) continue;
                    keys.Add(option.Key);
                }
            }

            if (SameOptionKeys(keys)) return;
            while (_tasks.Columns.Count > _optionColumnStart)
                _tasks.Columns.RemoveAt(_tasks.Columns.Count - 1);
            _optionKeys.Clear();
            foreach (string key in keys)
                AddOptionColumn(key);
        }

        private void AddOptionColumn(string key)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn
            {
                HeaderText = key,
                Width = 120,
                ReadOnly = false,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            column.HeaderCell.ToolTipText = OptionTooltip(key);
            _tasks.Columns.Add(column);
            _optionKeys.Add(key);
        }

        private bool SameOptionKeys(IReadOnlyList<string> keys)
        {
            if (keys.Count != _optionKeys.Count) return false;
            for (int i = 0; i < keys.Count; i++)
            {
                if (!string.Equals(keys[i], _optionKeys[i], StringComparison.Ordinal)) return false;
            }
            return true;
        }

        private static bool ContainsKey(IReadOnlyList<string> keys, string key)
        {
            foreach (string item in keys)
            {
                if (string.Equals(item, key, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        private static string OptionTooltip(string key)
        {
            if (key == SqlSheetHeader.StartCellKey)
                return "结果从表的这个单元格开始贴，含标题行。默认 A1。";
            if (key == SqlSheetHeader.ClearExtraColumnsKey)
                return "是否清掉结果没有覆盖到的右侧旧列。默认是。";
            return "会保留并写回，当前刷新不使用这项。";
        }

        private static string OptionDisplay(RefreshTaskDefinition task, string key)
        {
            foreach (SqlSheetOption option in SqlSheetHeader.SplitOptions(task.OptionsText))
            {
                if (!string.Equals(option.Key, key, StringComparison.Ordinal)) continue;
                return option.Value ?? string.Empty;
            }

            if (key == SqlSheetHeader.StartCellKey) return "A1";
            if (key == SqlSheetHeader.ClearExtraColumnsKey) return "是";
            return string.Empty;
        }

        /// <summary>按勾选任务批量刷新。没有当前连接时先说明如何设置，不开始刷新。</summary>
        /// <remarks>没有连接方案时指向连接管理新建；有方案但未设当前连接时指向左侧列表右键。</remarks>
        private void RefreshTasks()
        {
            ConnectionProfileSnapshot connection = _connectionTree.ExecutionConnection();
            if (connection == null)
            {
                MessageBox.Show(
                    this,
                    _connectionTree.HasConnectionProfiles
                        ? AppPresentation.SetCurrentConnectionPrompt
                        : AppPresentation.NoConnectionPrompt,
                    "批量刷新",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string duplicateMessage = SqlSheetHeader.DescribeDuplicateTargets(_rows.Select(item => item.Task).ToList());
            if (!string.IsNullOrEmpty(duplicateMessage))
            {
                MessageBox.Show(this, duplicateMessage, "批量刷新", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<TaskRow> selected = _rows.Where(item => item.Checked).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "请选择要刷新的任务。", "批量刷新", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            List<RefreshTaskDefinition> executable = selected
                .Where(item => string.IsNullOrEmpty(item.Task.OptionsError))
                .Select(item => item.Task)
                .ToList();
            string optionErrors = DescribeOptionErrors(selected);
            if (executable.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "已勾选的任务参数有误，不能刷新。" + Environment.NewLine + Environment.NewLine + optionErrors,
                    "批量刷新",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!ConfirmRefresh(executable.Count, optionErrors)) return;

            // 写入会激活目标表。抑制到结果和结束说明都结束，避免内层对话框提前放开重扫。
            PushActivateSuppress();
            try
            {
                BatchRefreshResult result = _operationRunner.Run(
                    this,
                    "批量刷新 Sheet",
                    true,
                    (progress, token) => _refreshService.RefreshAsync(
                        _workbook,
                        connection,
                        executable,
                        _settings.Get(CoreSettings.BatchMode),
                        _settings.Get(CoreSettings.MaxParallelism),
                        _settings.Get(CoreSettings.MaxExportRows),
                        _settings.Get(CoreSettings.ResultBlockSize),
                        _settings.Get(CoreSettings.QueryTimeoutSeconds),
                        Guid.NewGuid().ToString("N"),
                        progress,
                        token));
                ApplyResults(result);
            }
            catch (OperationCanceledException)
            {
                _status.Text = "批量刷新已取消。";
            }
            catch (Exception exception)
            {
                _status.Text = exception.Message;
                ShowOwnedError(exception);
            }
            finally
            {
                PopActivateSuppress();
            }
        }

        /// <summary>列出已勾选且参数有误的任务。没有这类任务时为空。</summary>
        /// <param name="selected">当前勾选的任务行。</param>
        /// <returns>每个错误一行。句式与重名提示一致。没有错误时为空。</returns>
        private static string DescribeOptionErrors(IReadOnlyList<TaskRow> selected)
        {
            StringBuilder message = new StringBuilder();
            foreach (TaskRow row in selected)
            {
                string error = row.Task.OptionsError;
                if (string.IsNullOrEmpty(error)) continue;
                if (message.Length > 0) message.AppendLine();
                message.Append("目标表“");
                message.Append(row.Task.TargetSheetName);
                message.Append("”：");
                message.Append(error);
                if (!error.EndsWith("。", StringComparison.Ordinal)) message.Append("。");
            }

            return message.ToString();
        }

        /// <summary>执行前让用户确认。取消则不开始刷新。</summary>
        /// <param name="executableCount">参数有效、实际会执行的任务数。</param>
        /// <param name="optionErrors">已勾选但参数有误、将跳过的说明。没有时为空。</param>
        /// <returns>用户选择「是」时为 true。</returns>
        private bool ConfirmRefresh(int executableCount, string optionErrors)
        {
            string message = "确定刷新已勾选的 " + executableCount.ToString() + " 个任务吗？";
            if (!string.IsNullOrEmpty(optionErrors))
                message += Environment.NewLine + optionErrors;

            PushActivateSuppress();
            try
            {
                return MessageBox.Show(
                    this,
                    message,
                    "批量刷新",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.Yes;
            }
            finally
            {
                PopActivateSuppress();
            }
        }

        private void ApplyResults(BatchRefreshResult result)
        {
            Dictionary<string, RefreshTaskResult> byId = result.Tasks.ToDictionary(item => item.Task.Id, StringComparer.OrdinalIgnoreCase);
            foreach (TaskRow row in _rows)
            {
                if (!row.Checked || !string.IsNullOrEmpty(row.Task.OptionsError)) continue;
                if (!byId.TryGetValue(row.Task.Id, out RefreshTaskResult taskResult)) continue;
                if (taskResult.Succeeded && taskResult.IsTruncated)
                {
                    row.State = TaskVisualState.Truncated;
                    row.Message = taskResult.Message;
                }
                else if (taskResult.Succeeded)
                {
                    row.State = TaskVisualState.Success;
                    row.Message = taskResult.Message;
                }
                else
                {
                    row.State = TaskVisualState.Error;
                    row.Message = taskResult.Message;
                }
            }

            int succeeded = result.Tasks.Count(item => item.Succeeded && !item.IsTruncated);
            int truncated = result.Tasks.Count(item => item.Succeeded && item.IsTruncated);
            int failed = result.Tasks.Count(item => !item.Succeeded && !item.Skipped);
            _scanSummary = string.Format("刷新完成：成功 {0}，失败 {1}，截断 {2}。", succeeded, failed, truncated);
            _status.Text = _scanSummary;
            foreach (DataGridViewRow gridRow in _tasks.Rows)
            {
                TaskRow row = gridRow.Tag as TaskRow;
                if (row != null) DecorateRow(gridRow, row);
            }
            _tasks.InvalidateColumn(_statusColumn);
            if (_settings.Get(CoreSettings.BeautifyAfterRefresh))
                BeautifySqlSheet(false);
            _status.Text = _scanSummary;
            ShowRefreshSummary(result, succeeded, failed, truncated);
        }

        /// <summary>刷新正常结束后列出失败和截断，图例单独说明目标表标题行底色。取消或异常不走这里。</summary>
        private void ShowRefreshSummary(BatchRefreshResult result, int succeeded, int failed, int truncated)
        {
            string summary = string.Format("刷新完成：成功 {0}，失败 {1}，截断 {2}。", succeeded, failed, truncated);
            StringBuilder details = new StringBuilder();
            AppendDetailSection(details, "× [失败]", result.Tasks, item => !item.Succeeded && !item.Skipped);
            AppendDetailSection(details, "△ [截断]", result.Tasks, item => item.Succeeded && item.IsTruncated);

            PushActivateSuppress();
            try
            {
                RefreshSummaryForm.Show(
                    this,
                    summary,
                    details.ToString(),
                    _settings.Get(CoreSettings.ResultSuccessColor),
                    _settings.Get(CoreSettings.ResultTruncatedColor));
            }
            finally
            {
                PopActivateSuppress();
            }
        }

        /// <summary>把一类结果接进明细列表。每条是标记、表名和说明，写在同一行。没有条目时不写，也不留空行。</summary>
        /// <param name="details">正在组装的明细。失败写完后再写截断，中间不插标题。</param>
        /// <param name="marker">条目前缀，例如「× [失败]」或「△ [截断]」。</param>
        /// <param name="tasks">本次刷新的全部任务结果。</param>
        /// <param name="include">为 true 时该任务写入列表。</param>
        private static void AppendDetailSection(
            StringBuilder details,
            string marker,
            IReadOnlyList<RefreshTaskResult> tasks,
            Func<RefreshTaskResult, bool> include)
        {
            foreach (RefreshTaskResult item in tasks)
            {
                if (!include(item)) continue;
                if (details.Length > 0)
                    details.AppendLine();

                details.Append(marker);
                details.Append(" 「");
                details.Append(item.Task.TargetSheetName);
                details.Append("」");
                string message = SingleLine(item.Message);
                if (message.Length == 0) continue;
                details.Append("：");
                details.Append(message);
            }
        }

        /// <summary>把说明收成一行，避免条目被说明里的换行拆开。</summary>
        /// <param name="message">该条的说明。空白时返回空字符串。</param>
        /// <returns>去掉首尾空白，并把换行换成空格后的说明。</returns>
        private static string SingleLine(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return string.Empty;
            return message.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Trim();
        }

        /// <summary>表头复选框：未全选时全选，已全选时全不选。没有任务时不改变勾选。</summary>
        /// <remarks>在绑定期间写入勾选值，并刷新当前单元格，避免它把点击前的未勾写回去。</remarks>
        private void ToggleSelection()
        {
            if (_rows.Count == 0 || _tasks.IsDisposed) return;
            bool select = !_rows.All(item => item.Checked);
            _tasks.EndEdit();
            _binding = true;
            try
            {
                foreach (DataGridViewRow gridRow in _tasks.Rows)
                    WriteChecked(gridRow, select);
            }
            finally
            {
                _binding = false;
            }

            _tasks.RefreshEdit();
            InvalidateHeaderCheck();
            _status.Text = _scanSummary;
        }

        /// <summary>行勾选变化后重画表头复选框。</summary>
        private void InvalidateHeaderCheck()
        {
            if (_tasks.IsDisposed) return;
            _tasks.InvalidateCell(_checkColumn, -1);
        }

        private void UpdateModeLabel()
        {
            BatchExecutionMode mode = _settings.Get(CoreSettings.BatchMode);
            string execution = mode == BatchExecutionMode.Parallel
                ? "执行方式：并行，最大并发 " + _settings.Get(CoreSettings.MaxParallelism).ToString() + "（可修改）。"
                : "执行方式：串行（可修改）。";
            _hint.Text = "按住 Shift 点击，可一次勾上或去掉连续多行。"
                + execution
                + "最大导出 Excel 行：" + _settings.Get(CoreSettings.MaxExportRows).ToString() + "（可修改）。"
                + "查询超时：" + _settings.Get(CoreSettings.QueryTimeoutSeconds).ToString() + " 秒（可修改）。";
        }

        private void TasksCurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (_tasks.IsCurrentCellDirty && _tasks.CurrentCell is DataGridViewCheckBoxCell)
                _tasks.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void TasksCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_binding || e.RowIndex < 0 || e.ColumnIndex != _checkColumn) return;
            TaskRow row = _tasks.Rows[e.RowIndex].Tag as TaskRow;
            if (row == null) return;
            object checkValue = _tasks.Rows[e.RowIndex].Cells[_checkColumn].Value;
            row.Checked = checkValue is bool && (bool)checkValue;
            ApplyCheckState(row);
            InvalidateHeaderCheck();
            DecorateRow(_tasks.Rows[e.RowIndex], row);
            _tasks.InvalidateRow(e.RowIndex);
            if (_shiftClick)
                _shiftAdjustedRow = e.RowIndex;
            else
            {
                _rangeAnchor = e.RowIndex;
                _rangeAnchorChecked = row.Checked;
            }
            TasksSelectionChanged(this, EventArgs.Empty);
        }

        private void TasksCellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (_shiftClick)
            {
                e.Cancel = true;
                return;
            }
            if (e.ColumnIndex == _checkColumn || e.ColumnIndex >= _optionColumnStart) return;
            e.Cancel = true;
        }

        private void TasksCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (_binding || _discardEdit || e.RowIndex < 0 || e.ColumnIndex < _optionColumnStart) return;
            DataGridViewRow gridRow = _tasks.Rows[e.RowIndex];
            TaskRow row = gridRow.Tag as TaskRow;
            if (row == null) return;
            string edited = CollectOptions(gridRow);
            SqlSheetHeader header = SqlSheetHeader.ParseOptions(row.Task.TargetSheetName, edited);
            try
            {
                _taskReader.WriteColumnHeader(_workbook, row.Task.SourceColumn, header.HeaderText);
            }
            catch (Exception exception)
            {
                DecorateRow(gridRow, row);
                _status.Text = exception.Message;
                ShowOwnedError(exception);
                return;
            }

            row.Task = ReplaceOptions(row.Task, header);
            ApplyRowStates();
            foreach (DataGridViewRow itemRow in _tasks.Rows)
            {
                TaskRow item = itemRow.Tag as TaskRow;
                if (item != null) DecorateRow(itemRow, item);
            }
            _tasks.InvalidateColumn(_statusColumn);
            TasksSelectionChanged(this, EventArgs.Empty);
        }

        /// <summary>右键先选中所点的任务行。左键按下时记下是否按住 Shift，点击结束后再决定跳转还是连续勾选。</summary>
        private void TasksCellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                DataGridViewCell cell = _tasks.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (_tasks.CurrentCell != cell)
                    _tasks.CurrentCell = cell;
                return;
            }

            if (e.Button != MouseButtons.Left || e.RowIndex < 0 || e.ColumnIndex < 0) return;
            _shiftClick = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
        }

        /// <summary>
        /// 把起点那次的勾选结果写到起点与所点行之间。没有起点时只切换所点的这一行，并把它记成起点。
        /// </summary>
        /// <param name="rowIndex">这次 Shift 点击的任务行。</param>
        /// <remarks>
        /// 延后到这次点击处理完再写。勾选框有时会先自己翻转，那种情况不再翻第二次。
        /// 起点只在单击勾选框时移动，方便按住 Shift 继续往更远的行扩。
        /// </remarks>
        private void ApplyShiftRange(int rowIndex)
        {
            if (!_tasks.IsDisposed && _tasks.IsCurrentCellInEditMode)
            {
                if (_tasks.CurrentCell is DataGridViewCheckBoxCell)
                    _tasks.EndEdit();
                else
                {
                    _discardEdit = true;
                    try
                    {
                        _tasks.CancelEdit();
                    }
                    finally
                    {
                        _discardEdit = false;
                    }
                }
            }

            bool alreadyToggled = _shiftAdjustedRow == rowIndex;
            _shiftAdjustedRow = -1;
            _shiftClick = false;
            if (_tasks.IsDisposed || rowIndex < 0 || rowIndex >= _tasks.Rows.Count) return;

            if (_rangeAnchor < 0 || _rangeAnchor >= _tasks.Rows.Count)
            {
                TaskRow row = _tasks.Rows[rowIndex].Tag as TaskRow;
                if (row == null) return;
                if (!alreadyToggled)
                {
                    _binding = true;
                    try
                    {
                        WriteChecked(_tasks.Rows[rowIndex], !row.Checked);
                    }
                    finally
                    {
                        _binding = false;
                    }
                    _tasks.RefreshEdit();
                    InvalidateHeaderCheck();
                    _tasks.InvalidateColumn(_statusColumn);
                }
                _rangeAnchor = rowIndex;
                _rangeAnchorChecked = row.Checked;
                TasksSelectionChanged(this, EventArgs.Empty);
                return;
            }

            int start = Math.Min(_rangeAnchor, rowIndex);
            int end = Math.Max(_rangeAnchor, rowIndex);
            bool select = _rangeAnchorChecked;
            _binding = true;
            try
            {
                for (int i = start; i <= end; i++)
                    WriteChecked(_tasks.Rows[i], select);
            }
            finally
            {
                _binding = false;
            }

            _tasks.RefreshEdit();
            InvalidateHeaderCheck();
            _tasks.InvalidateColumn(_statusColumn);
            TasksSelectionChanged(this, EventArgs.Empty);
        }

        /// <summary>把一行的勾选结果写进任务和勾选框，并刷新该行颜色。</summary>
        /// <param name="gridRow">任务列表中的一行。</param>
        /// <param name="select">这一行要变成的勾选结果。</param>
        /// <remarks>调用方在绑定期间调用，避免勾选框把这次写入再报一遍。</remarks>
        private void WriteChecked(DataGridViewRow gridRow, bool select)
        {
            TaskRow row = gridRow.Tag as TaskRow;
            if (row == null) return;
            row.Checked = select;
            ApplyCheckState(row);
            gridRow.Cells[_checkColumn].Value = select;
            DecorateRow(gridRow, row);
            if (gridRow.DataGridView != null)
                gridRow.DataGridView.InvalidateRow(gridRow.Index);
        }

        /// <summary>只有点在任务单元格上才弹出菜单。说明为空时不能复制。</summary>
        private void TaskMenuOpening(object sender, CancelEventArgs e)
        {
            Point point = _tasks.PointToClient(Cursor.Position);
            DataGridView.HitTestInfo hit = _tasks.HitTest(point.X, point.Y);
            _contextRow = hit.RowIndex < 0 ? null : _tasks.Rows[hit.RowIndex].Tag as TaskRow;
            if (_contextRow == null)
            {
                e.Cancel = true;
                return;
            }

            _copyResultItem.Enabled = !string.IsNullOrEmpty(_contextRow.Message);
        }

        /// <summary>把右键所在任务行的结果说明写入剪贴板。</summary>
        /// <remarks>剪贴板偶发占用时不打断操作。</remarks>
        private void CopyResultMessage(object sender, EventArgs e)
        {
            TaskRow row = _contextRow;
            if (row == null || string.IsNullOrEmpty(row.Message)) return;
            try
            {
                Clipboard.SetText(row.Message);
            }
            catch (Exception)
            {
                // 剪贴板偶发占用；复制失败不打断操作。
            }
        }

        private void TasksCellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (_shiftClick)
            {
                int rowIndex = e.RowIndex;
                BeginInvoke(new Action(() => ApplyShiftRange(rowIndex)));
                return;
            }

            TaskRow row = _tasks.Rows[e.RowIndex].Tag as TaskRow;
            if (row == null) return;
            if (e.ColumnIndex == _targetColumn) JumpToTarget(row);
            else if (e.ColumnIndex == _summaryColumn) JumpToSql(row);
        }

        private void TasksCellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            bool link = e.RowIndex >= 0 && (e.ColumnIndex == _targetColumn || e.ColumnIndex == _summaryColumn);
            Cursor next = link ? Cursors.Hand : Cursors.Default;
            if (_tasks.Cursor != next) _tasks.Cursor = next;
        }

        /// <summary>激活 SQL 页并选中该任务命令开始的单元格。</summary>
        private void JumpToSql(TaskRow row)
        {
            int excelRow = row.Task.SqlCommandRow < 2 ? 2 : row.Task.SqlCommandRow;
            try
            {
                _taskReader.ActivateSqlCell(_workbook, row.Task.SourceColumn, excelRow);
            }
            catch (Exception exception)
            {
                ShowOwnedError(exception);
            }
        }

        /// <summary>激活目标工作表。表还不存在时只提示，不创建。</summary>
        private void JumpToTarget(TaskRow row)
        {
            try
            {
                if (_taskReader.TryActivateWorksheet(_workbook, row.Task.TargetSheetName)) return;
                PushActivateSuppress();
                try
                {
                    MessageBox.Show(
                        this,
                        "目标表“" + row.Task.TargetSheetName + "”还不存在，刷新时会创建。",
                        "批量刷新",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                finally
                {
                    PopActivateSuppress();
                }
            }
            catch (Exception exception)
            {
                ShowOwnedError(exception);
            }
        }

        /// <summary>点击勾选列表头时全选或全不选。</summary>
        /// <remarks>延后到这次点击处理完再改勾选。否则当前行的复选框会在全选之后被拨回去。</remarks>
        private void TasksColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex != _checkColumn || e.Button != MouseButtons.Left) return;
            BeginInvoke(new Action(ToggleSelection));
        }

        private void TasksCellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                if (e.ColumnIndex == _checkColumn) PaintHeaderCheck(e);
                return;
            }
            if (e.ColumnIndex != _statusColumn) return;
            e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
            TaskRow row = _tasks.Rows[e.RowIndex].Tag as TaskRow;
            Color? color = StatusColor(row);
            if (color.HasValue)
            {
                int size = Math.Min(14, Math.Max(8, e.CellBounds.Height - 10));
                Rectangle box = new Rectangle(
                    e.CellBounds.X + (e.CellBounds.Width - size) / 2,
                    e.CellBounds.Y + (e.CellBounds.Height - size) / 2,
                    size,
                    size);
                SmoothingMode smoothing = e.Graphics.SmoothingMode;
                using (Region previousClip = e.Graphics.Clip)
                {
                    e.Graphics.SetClip(e.CellBounds, CombineMode.Intersect);
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    using (SolidBrush brush = new SolidBrush(color.Value))
                    {
                        e.Graphics.FillEllipse(brush, box);
                    }
                    e.Graphics.SmoothingMode = smoothing;
                    e.Graphics.Clip = previousClip;
                }
            }
            e.Handled = true;
        }

        /// <summary>在勾选列表头绘制全选框。全部勾选为勾选，只勾一部分为半选。</summary>
        private void PaintHeaderCheck(DataGridViewCellPaintingEventArgs e)
        {
            e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
            int checkedCount = 0;
            foreach (TaskRow row in _rows)
            {
                if (row.Checked) checkedCount++;
            }

            bool all = _rows.Count > 0 && checkedCount == _rows.Count;
            bool partial = checkedCount > 0 && !all;
            if (Application.RenderWithVisualStyles)
            {
                CheckBoxState state = all
                    ? CheckBoxState.CheckedNormal
                    : partial ? CheckBoxState.MixedNormal : CheckBoxState.UncheckedNormal;
                Size glyph = CheckBoxRenderer.GetGlyphSize(e.Graphics, state);
                Point origin = new Point(
                    e.CellBounds.X + (e.CellBounds.Width - glyph.Width) / 2,
                    e.CellBounds.Y + (e.CellBounds.Height - glyph.Height) / 2);
                CheckBoxRenderer.DrawCheckBox(e.Graphics, origin, state);
            }
            else
            {
                ButtonState state = all
                    ? ButtonState.Checked
                    : partial ? ButtonState.Checked | ButtonState.Inactive : ButtonState.Normal;
                const int size = 13;
                Rectangle box = new Rectangle(
                    e.CellBounds.X + (e.CellBounds.Width - size) / 2,
                    e.CellBounds.Y + (e.CellBounds.Height - size) / 2,
                    size,
                    size);
                ControlPaint.DrawCheckBox(e.Graphics, box, state);
            }

            e.Handled = true;
        }

        private void DecorateRow(DataGridViewRow gridRow, TaskRow row)
        {
            gridRow.DefaultCellStyle.BackColor = row.Checked ? CheckedRowBackColor : Color.White;
            gridRow.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            gridRow.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;

            DataGridViewCell target = gridRow.Cells[_targetColumn];
            target.Value = row.Task.TargetSheetName;
            target.ToolTipText = "转到目标表";
            ApplyLinkStyle(target);

            DataGridViewCell summary = gridRow.Cells[_summaryColumn];
            bool hasSummary = !string.IsNullOrEmpty(row.Task.Summary);
            summary.Value = hasSummary ? row.Task.Summary : "无";
            summary.ToolTipText = hasSummary ? row.Task.Summary : "转到 SQL 命令";
            ApplyLinkStyle(summary);

            gridRow.Cells[_statusColumn].ToolTipText = row.Message ?? string.Empty;
            for (int i = 0; i < _optionKeys.Count; i++)
                gridRow.Cells[_optionColumnStart + i].Value = OptionDisplay(row.Task, _optionKeys[i]);
        }

        private void ApplyLinkStyle(DataGridViewCell cell)
        {
            cell.Style.ForeColor = Color.Blue;
            cell.Style.Font = _linkFont;
            cell.Style.SelectionForeColor = SystemColors.HighlightText;
        }

        private void TasksSelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow selected = _tasks.CurrentRow;
            if (selected == null || !selected.Selected)
            {
                _status.Text = _scanSummary;
                return;
            }

            TaskRow row = selected.Tag as TaskRow;
            _status.Text = row == null || string.IsNullOrEmpty(row.Message) ? _scanSummary : row.Message;
        }

        private void WorkspaceSizeChanged(object sender, EventArgs e)
        {
            _userDraggingSplit = false;
            _applyingSplit = true;
            try
            {
                ApplySplit();
            }
            finally
            {
                _applyingSplit = false;
            }
        }

        private void ApplySplit()
        {
            ApplySplitterPanelMinSizes(_workspace, LeftMinWidth, RightMinWidth);
            int available = _workspace.Width - _workspace.SplitterWidth;
            if (available <= 0) return;
            int ratio = _settings.Get(CoreSettings.RefreshSplitRatio);
            int distance = ratio > 0 ? available * ratio / 1000 : available / 5;
            TrySetSplitterDistance(_workspace, distance);
        }

        private void RememberSplit()
        {
            if (!_userDraggingSplit) return;
            _userDraggingSplit = false;
            if (_applyingSplit || _workspace.IsDisposed) return;
            int available = _workspace.Width - _workspace.SplitterWidth;
            if (available <= 0) return;
            int ratio = (int)Math.Round(_workspace.SplitterDistance * 1000d / available);
            ratio = Math.Max(1, Math.Min(999, ratio));
            if (ratio == _settings.Get(CoreSettings.RefreshSplitRatio)) return;
            _settings.Set(CoreSettings.RefreshSplitRatio, ratio);
        }

        private static void ApplySplitterPanelMinSizes(SplitContainer split, int panel1MinSize, int panel2MinSize)
        {
            if (split == null || split.IsDisposed) return;
            int available = split.Width - split.SplitterWidth;
            if (available < panel1MinSize + panel2MinSize) return;
            if (split.Panel1MinSize != panel1MinSize) split.Panel1MinSize = panel1MinSize;
            if (split.Panel2MinSize != panel2MinSize) split.Panel2MinSize = panel2MinSize;
        }

        private static void TrySetSplitterDistance(SplitContainer split, int preferred)
        {
            if (split == null || split.IsDisposed) return;
            int minimum = split.Panel1MinSize;
            int maximum = split.Width - split.Panel2MinSize - split.SplitterWidth;
            if (maximum < minimum) return;
            int distance = Math.Max(minimum, Math.Min(maximum, preferred));
            if (split.SplitterDistance != distance) split.SplitterDistance = distance;
        }

        private string CollectOptions(DataGridViewRow gridRow)
        {
            List<string> parts = new List<string>();
            for (int i = 0; i < _optionKeys.Count; i++)
            {
                string key = _optionKeys[i];
                string value = (Convert.ToString(gridRow.Cells[_optionColumnStart + i].Value) ?? string.Empty).Trim();
                if (key == SqlSheetHeader.StartCellKey)
                {
                    if (value.Length == 0 || string.Equals(value, "A1", StringComparison.OrdinalIgnoreCase)) continue;
                    parts.Add(key + "=" + value);
                    continue;
                }

                if (key == SqlSheetHeader.ClearExtraColumnsKey)
                {
                    if (value.Length == 0 || value == "是") continue;
                    parts.Add(key + "=" + value);
                    continue;
                }

                if (value.Length > 0) parts.Add(key + "=" + value);
            }

            return string.Join(";", parts);
        }

        private static string HeaderText(RefreshTaskDefinition task)
        {
            return string.IsNullOrEmpty(task.OptionsText)
                ? task.TargetSheetName
                : task.TargetSheetName + "//" + task.OptionsText;
        }

        private static RefreshTaskDefinition ReplaceOptions(RefreshTaskDefinition task, SqlSheetHeader header)
        {
            return new RefreshTaskDefinition(
                task.Id,
                task.TargetSheetName,
                task.QueryText,
                task.SourceColumn,
                header.OptionsText,
                header.StartCell,
                header.StartRow,
                header.StartColumn,
                header.ClearExtraColumns,
                header.Error,
                task.Summary,
                false,
                task.SqlCommandRow);
        }

        private void SheetRefreshFormFormClosed(object sender, FormClosedEventArgs e)
        {
            foreach (Image icon in _actionIcons)
                icon.Dispose();
            _actionIcons.Clear();
            _linkFont.Dispose();
            _hintFont.Dispose();
            _taskMenu.Dispose();
        }

        /// <summary>按当前设置给出该行状态圆点的颜色。未勾选时不着色。</summary>
        /// <param name="row">任务行。为空时不着色。</param>
        /// <returns>圆点颜色。不着色时为空。</returns>
        private Color? StatusColor(TaskRow row)
        {
            if (row == null) return null;
            if (row.State == TaskVisualState.Error) return _settings.Get(CoreSettings.ResultErrorColor);
            if (!row.Checked) return null;
            if (row.State == TaskVisualState.Pending) return _settings.Get(CoreSettings.ResultPendingColor);
            if (row.State == TaskVisualState.Success) return _settings.Get(CoreSettings.ResultSuccessColor);
            if (row.State == TaskVisualState.Truncated) return _settings.Get(CoreSettings.ResultTruncatedColor);
            return null;
        }

        /// <summary>格式菜单：美化整张 SQL 页，或打开参数设置里的格式页。</summary>
        /// <returns>停靠在窗体顶部的菜单。调用方加入控件集合。</returns>
        private MenuStrip BuildFormatMenu()
        {
            ToolStripMenuItem formatMenu = new ToolStripMenuItem("格式");
            formatMenu.DropDownItems.Add(new ToolStripMenuItem("美化格式", null, (sender, args) => BeautifySqlSheet(true)));
            formatMenu.DropDownItems.Add(new ToolStripMenuItem("格式设置", null, (sender, args) => OpenSqlFormatSettings()));

            MenuStrip menu = new MenuStrip { Dock = DockStyle.Top };
            menu.Items.Add(formatMenu);
            MainMenuStrip = menu;
            return menu;
        }

        /// <summary>打开参数设置并停在「sql页格式」。期间不因窗体重新激活而重扫任务。</summary>
        private void OpenSqlFormatSettings()
        {
            if (_openSqlFormatSettings == null) return;
            PushActivateSuppress();
            try
            {
                _openSqlFormatSettings();
            }
            finally
            {
                PopActivateSuppress();
            }
        }

        /// <summary>用进度窗美化当前工作簿的 SQL 页。没有可美化的列时不打开进度窗。</summary>
        /// <param name="acknowledge">为 true 时，成功或没有可美化的列弹出确定框。刷新结束后的自动美化传 false，结果仍由刷新结束说明承担。</param>
        /// <remarks>取消只改状态栏。失败弹出错误详情。确定框在界面线程上模态显示，并留在激活抑制之内，避免被 Excel 挡住。</remarks>
        private void BeautifySqlSheet(bool acknowledge)
        {
            int count;
            try
            {
                count = _taskReader.CountStyleColumns(_workbook);
            }
            catch (Exception exception)
            {
                _status.Text = exception.Message;
                ShowOwnedError(exception);
                return;
            }

            if (count == 0)
            {
                const string message = "没有可美化的列。";
                _status.Text = message;
                if (acknowledge)
                {
                    MessageBox.Show(this, message, "美化格式", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                return;
            }

            PushActivateSuppress();
            try
            {
                int styled = _operationRunner.Run(
                    this,
                    "美化格式",
                    true,
                    (progress, token) => _taskReader.FormatSqlSheet(_workbook, progress, token));
                string message = string.Format("已美化 {0} 列。", styled);
                _status.Text = message;
                if (acknowledge)
                {
                    MessageBox.Show(this, message, "美化格式", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                _status.Text = "美化已取消。";
            }
            catch (Exception exception)
            {
                _status.Text = exception.Message;
                ShowOwnedError(exception);
            }
            finally
            {
                PopActivateSuppress();
            }
        }

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
                UseVisualStyleBackColor = true
            };
        }

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

        /// <summary>列表中的一行任务，以及当次刷新的颜色状态。</summary>
        private sealed class TaskRow
        {
            public TaskRow(RefreshTaskDefinition task, bool isChecked)
            {
                Task = task;
                Checked = isChecked;
            }

            public RefreshTaskDefinition Task { get; set; }
            public bool Checked { get; set; }
            public TaskVisualState State { get; set; }
            public string Message { get; set; }
            public string DuplicateMessage { get; set; }
        }

        /// <summary>任务行的颜色状态。未勾选时状态列不着色，行底色仍表示是否勾选。</summary>
        private enum TaskVisualState
        {
            None,
            Pending,
            Success,
            Truncated,
            Error
        }
    }
}

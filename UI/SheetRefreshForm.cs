using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace DB2Sheet.UI
{
    /// <summary>扫描当前工作簿的 SQL Sheet，并批量刷新对应目标工作表。</summary>
    /// <remarks>
    /// 窗体读取 SQL Sheet 作为任务输入，使用选中的共享连接执行查询，并将结果写回同一工作簿。
    /// </remarks>
    public sealed class SheetRefreshForm : AppForm
    {
        private readonly IConnectionProfileRepository _connections;
        private readonly IProviderRegistry _providers;
        private readonly IDataSourceExecutionService _executionService;
        private readonly IOperationRunner _operationRunner;
        private readonly ISettingsStore _settings;
        private readonly ISqlSheetTaskReader _taskReader;
        private readonly IBatchRefreshService _refreshService;
        private readonly ExcelInterop.Workbook _workbook;
        private readonly ConnectionProfileSelector _connectionSelector;
        private readonly ListView _tasks;
        private readonly Label _status;
        private IReadOnlyList<RefreshTaskDefinition> _definitions = new List<RefreshTaskDefinition>().AsReadOnly();

        /// <summary>创建 Sheet 批量刷新窗体。</summary>
        /// <param name="connections">连接方案仓储。</param>
        /// <param name="providers">数据源提供程序注册表。</param>
        /// <param name="executionService">连接测试和查询执行服务。</param>
        /// <param name="operationRunner">批次进度窗体运行器。</param>
        /// <param name="settings">批次模式、并发数、行数和超时设置。</param>
        /// <param name="taskReader">SQL Sheet 任务读取器。</param>
        /// <param name="refreshService">批量刷新编排服务。</param>
        /// <param name="workbook">当前操作的 Excel 工作簿。</param>
        public SheetRefreshForm(
            IConnectionProfileRepository connections,
            IProviderRegistry providers,
            IDataSourceExecutionService executionService,
            IOperationRunner operationRunner,
            ISettingsStore settings,
            ISqlSheetTaskReader taskReader,
            IBatchRefreshService refreshService,
            ExcelInterop.Workbook workbook)
        {
            _connections = connections ?? throw new ArgumentNullException(nameof(connections));
            _providers = providers ?? throw new ArgumentNullException(nameof(providers));
            _executionService = executionService ?? throw new ArgumentNullException(nameof(executionService));
            _operationRunner = operationRunner ?? throw new ArgumentNullException(nameof(operationRunner));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _taskReader = taskReader ?? throw new ArgumentNullException(nameof(taskReader));
            _refreshService = refreshService ?? throw new ArgumentNullException(nameof(refreshService));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));

            Text = AppPresentation.WindowTitle("Sheet 批量刷新");
            Width = 900;
            Height = 620;
            MinimumSize = new Size(700, 460);
            StartPosition = FormStartPosition.CenterScreen;

            _connectionSelector = new ConnectionProfileSelector(_connections) { Dock = DockStyle.Fill };
            _tasks = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false
            };
            _tasks.Columns.Add("源列", 80);
            _tasks.Columns.Add("目标 Sheet", 200);
            _tasks.Columns.Add("SQL 摘要", 360);
            _tasks.Columns.Add("状态", 180);
            _status = new Label { AutoSize = true, ForeColor = SystemColors.GrayText };

            Button scanButton = new Button { Text = "重新扫描 SQL Sheet", AutoSize = true };
            Button refreshButton = new Button { Text = "刷新全部", AutoSize = true };
            Button closeButton = new Button { Text = "关闭", AutoSize = true };

            TableLayoutPanel selector = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 2
            };
            selector.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            selector.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            selector.Controls.Add(new Label { Text = "共享连接", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            selector.Controls.Add(_connectionSelector, 1, 0);

            FlowLayoutPanel actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            actions.Controls.Add(scanButton);
            actions.Controls.Add(refreshButton);
            actions.Controls.Add(closeButton);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 4
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(selector, 0, 0);
            root.Controls.Add(actions, 0, 1);
            root.Controls.Add(_tasks, 0, 2);
            root.Controls.Add(_status, 0, 3);
            Controls.Add(root);

            scanButton.Click += (sender, args) => ScanTasks();
            refreshButton.Click += (sender, args) => RefreshTasks();
            closeButton.Click += (sender, args) => Close();
            _connectionSelector.ManageRequested += ManageConnections;
            ScanTasks();
        }

        private void ScanTasks()
        {
            try
            {
                _definitions = _taskReader.ReadTasks(_workbook);
                _tasks.Items.Clear();
                foreach (RefreshTaskDefinition task in _definitions)
                {
                    ListViewItem item = new ListViewItem(ColumnName(task.SourceColumn)) { Tag = task };
                    item.SubItems.Add(task.TargetSheetName);
                    item.SubItems.Add(Summarize(task.QueryText));
                    item.SubItems.Add("待刷新");
                    _tasks.Items.Add(item);
                }
                _status.Text = _definitions.Count == 0
                    ? "SQL Sheet 中未发现有效任务。"
                    : string.Format("发现 {0} 个刷新任务。", _definitions.Count);
            }
            catch (Exception exception)
            {
                _definitions = new List<RefreshTaskDefinition>().AsReadOnly();
                _tasks.Items.Clear();
                ShowError(exception);
            }
        }

        private void RefreshTasks()
        {
            ConnectionProfileSnapshot connection = _connectionSelector.SelectedProfile;
            if (connection == null)
            {
                MessageBox.Show(this, "请选择连接方案。", "批量刷新", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (_definitions.Count == 0)
            {
                MessageBox.Show(this, "没有可执行的刷新任务。", "批量刷新", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                BatchRefreshResult result = _operationRunner.Run(
                    this,
                    "批量刷新 Sheet",
                    true,
                    (progress, token) => _refreshService.RefreshAsync(
                        _workbook,
                        connection,
                        _definitions,
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
                ShowError(exception);
            }
        }

        private void ApplyResults(BatchRefreshResult result)
        {
            Dictionary<string, RefreshTaskResult> byId = result.Tasks.ToDictionary(item => item.Task.Id, StringComparer.OrdinalIgnoreCase);
            foreach (ListViewItem item in _tasks.Items)
            {
                RefreshTaskDefinition task = item.Tag as RefreshTaskDefinition;
                if (task == null || !byId.TryGetValue(task.Id, out RefreshTaskResult taskResult)) continue;
                item.SubItems[3].Text = taskResult.Message;
                item.ForeColor = taskResult.Succeeded ? Color.DarkGreen : taskResult.Skipped ? Color.DarkGoldenrod : Color.DarkRed;
            }
            int succeeded = result.Tasks.Count(item => item.Succeeded);
            int failed = result.Tasks.Count(item => !item.Succeeded && !item.Skipped);
            int skipped = result.Tasks.Count(item => item.Skipped);
            _status.Text = string.Format("刷新完成：成功 {0}，失败 {1}，跳过 {2}。", succeeded, failed, skipped);
        }

        private void ManageConnections(object sender, EventArgs e)
        {
            using (ConnectionProfilesForm form = new ConnectionProfilesForm(
                _connections, _providers, _executionService, _operationRunner, _connectionSelector.SelectedProfileId))
            {
                form.ShowDialog(this);
                if (!string.IsNullOrWhiteSpace(form.SelectedProfileId))
                    _connectionSelector.SelectedProfileId = form.SelectedProfileId;
            }
        }

        private void ShowError(Exception exception)
        {
            _status.Text = exception.Message;
            MessageBox.Show(this, exception.Message, AppPresentation.DisplayName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static string Summarize(string query)
        {
            string singleLine = (query ?? string.Empty).Replace('\r', ' ').Replace('\n', ' ').Trim();
            return singleLine.Length <= 100 ? singleLine : singleLine.Substring(0, 100) + "…";
        }

        private static string ColumnName(int column)
        {
            string name = string.Empty;
            while (column > 0)
            {
                column--;
                name = (char)('A' + column % 26) + name;
                column /= 26;
            }
            return name;
        }
    }
}

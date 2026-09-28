using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
        private readonly ConnectionProfileSelector _connectionSelector;
        private readonly ComboBox _queryProfiles;
        private readonly TextBox _queryName;
        private readonly TextBox _targetSheet;
        private readonly TextBox _sql;
        private readonly DataGridView _preview;
        private readonly Label _status;
        private string _currentQueryId;
        private DateTime _currentCreatedUtc;

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
            Width = 1100;
            Height = 760;
            MinimumSize = new Size(820, 600);
            StartPosition = FormStartPosition.CenterScreen;

            _connectionSelector = new ConnectionProfileSelector(_connections) { Dock = DockStyle.Fill };
            _queryProfiles = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                DisplayMember = nameof(QueryItem.DisplayName)
            };
            _queryName = new TextBox { Dock = DockStyle.Fill };
            _targetSheet = new TextBox { Dock = DockStyle.Fill, Text = "查询结果" };
            _sql = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                AcceptsReturn = true,
                AcceptsTab = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font(AppPresentation.CodeFontName, AppPresentation.CodeFontSize)
            };
            _preview = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            };
            _status = new Label { AutoSize = true, ForeColor = SystemColors.GrayText };

            Button newButton = ButtonOf("新建");
            Button saveButton = ButtonOf("保存方案");
            Button deleteButton = ButtonOf("删除方案");
            Button previewButton = ButtonOf("预览");
            Button exportButton = ButtonOf("写入 Sheet");
            Button closeButton = ButtonOf("关闭");

            FlowLayoutPanel toolbar = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            toolbar.Controls.Add(newButton);
            toolbar.Controls.Add(saveButton);
            toolbar.Controls.Add(deleteButton);
            toolbar.Controls.Add(previewButton);
            toolbar.Controls.Add(exportButton);
            toolbar.Controls.Add(closeButton);

            TableLayoutPanel metadata = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 4
            };
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            metadata.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            metadata.Controls.Add(LabelOf("查询方案"), 0, 0);
            metadata.Controls.Add(_queryProfiles, 1, 0);
            metadata.Controls.Add(LabelOf("共享连接"), 2, 0);
            metadata.Controls.Add(_connectionSelector, 3, 0);
            metadata.Controls.Add(LabelOf("方案名称"), 0, 1);
            metadata.Controls.Add(_queryName, 1, 1);
            metadata.Controls.Add(LabelOf("目标 Sheet"), 2, 1);
            metadata.Controls.Add(_targetSheet, 3, 1);

            SplitContainer workspace = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 280
            };
            workspace.Panel1.Controls.Add(_sql);
            workspace.Panel1.Controls.Add(new Label { Text = "SQL（仅允许只读查询）", Dock = DockStyle.Top, AutoSize = true });
            workspace.Panel2.Controls.Add(_preview);
            workspace.Panel2.Controls.Add(new Label { Text = "结果预览", Dock = DockStyle.Top, AutoSize = true });

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
            root.Controls.Add(toolbar, 0, 0);
            root.Controls.Add(metadata, 0, 1);
            root.Controls.Add(workspace, 0, 2);
            root.Controls.Add(_status, 0, 3);
            Controls.Add(root);

            newButton.Click += (sender, args) => NewQuery();
            saveButton.Click += (sender, args) => SaveQuery();
            deleteButton.Click += (sender, args) => DeleteQuery();
            previewButton.Click += (sender, args) => PreviewQuery();
            exportButton.Click += (sender, args) => ExportQuery();
            closeButton.Click += (sender, args) => Close();
            _queryProfiles.SelectedIndexChanged += QueryProfilesSelectedIndexChanged;
            _connectionSelector.ManageRequested += ManageConnections;

            ReloadQueries();
        }

        private void ReloadQueries(string selectedId = null)
        {
            List<QueryItem> items = _queries.GetAll().Select(profile => new QueryItem(profile)).ToList();
            _queryProfiles.DataSource = null;
            _queryProfiles.DataSource = items;
            _queryProfiles.DisplayMember = nameof(QueryItem.DisplayName);
            if (!string.IsNullOrWhiteSpace(selectedId))
            {
                _queryProfiles.SelectedItem = items.FirstOrDefault(item =>
                    string.Equals(item.Profile.Id, selectedId, StringComparison.OrdinalIgnoreCase));
            }
            if (_queryProfiles.SelectedIndex < 0) NewQuery();
        }

        private void NewQuery()
        {
            _queryProfiles.SelectedIndex = -1;
            _currentQueryId = Guid.NewGuid().ToString("N");
            _currentCreatedUtc = DateTime.UtcNow;
            _queryName.Clear();
            _sql.Clear();
            _targetSheet.Text = "查询结果";
            _preview.Columns.Clear();
            _status.Text = "新查询方案。";
            _queryName.Focus();
        }

        private void QueryProfilesSelectedIndexChanged(object sender, EventArgs e)
        {
            QueryProfile profile = (_queryProfiles.SelectedItem as QueryItem)?.Profile;
            if (profile == null) return;
            _currentQueryId = profile.Id;
            _currentCreatedUtc = profile.CreatedUtc;
            _queryName.Text = profile.Name;
            _sql.Text = profile.QueryText;
            _targetSheet.Text = profile.TargetSheetName;
            _connectionSelector.SelectedProfileId = profile.ConnectionProfileId;
            _status.Text = "已加载查询方案。";
        }

        private void SaveQuery()
        {
            if (!ValidateInputs(false, out ConnectionProfileSnapshot connection)) return;
            QueryProfile profile = new QueryProfile
            {
                Id = _currentQueryId ?? Guid.NewGuid().ToString("N"),
                Name = _queryName.Text.Trim(),
                ProviderId = connection.ProviderId,
                ConnectionProfileId = connection.Id,
                QueryText = _sql.Text,
                TargetSheetName = _targetSheet.Text.Trim(),
                CreatedUtc = _currentCreatedUtc == default(DateTime) ? DateTime.UtcNow : _currentCreatedUtc,
                UpdatedUtc = DateTime.UtcNow
            };
            _queries.Save(profile);
            _currentQueryId = profile.Id;
            _currentCreatedUtc = profile.CreatedUtc;
            ReloadQueries(profile.Id);
            _status.Text = "查询方案已保存。";
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

        private void PreviewQuery()
        {
            if (!ValidateInputs(true, out ConnectionProfileSnapshot connection)) return;
            try
            {
                DataSourceRequest request = CreateRequest(connection, ExecutionPurpose.Preview, _settings.Get(CoreSettings.MaxPreviewRows));
                BufferedQueryResult result = _operationRunner.Run(
                    this, "查询预览", true,
                    (progress, token) => _queryBuffer.ExecuteAsync(request, Guid.NewGuid().ToString("N"), progress, token));
                BindPreview(result);
                _status.Text = string.Format("预览完成，共 {0:N0} 行{1}。", result.RowCount, result.IsTruncated ? "（已截断）" : string.Empty);
            }
            catch (OperationCanceledException)
            {
                _status.Text = "查询预览已取消。";
            }
            catch (Exception exception)
            {
                ShowError(exception);
            }
        }

        private void ExportQuery()
        {
            if (!ValidateInputs(true, out ConnectionProfileSnapshot connection)) return;
            try
            {
                DataSourceRequest request = CreateRequest(connection, ExecutionPurpose.Export, _settings.Get(CoreSettings.MaxExportRows));
                string target = _targetSheet.Text.Trim();
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

        private DataSourceRequest CreateRequest(ConnectionProfileSnapshot connection, ExecutionPurpose purpose, int rowLimit)
        {
            return new DataSourceRequest(
                connection,
                _sql.Text,
                purpose,
                rowLimit,
                _settings.Get(CoreSettings.ResultBlockSize),
                _settings.Get(CoreSettings.QueryTimeoutSeconds));
        }

        private void BindPreview(BufferedQueryResult result)
        {
            _preview.SuspendLayout();
            try
            {
                _preview.Columns.Clear();
                foreach (ResultColumn column in result.Columns)
                {
                    _preview.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = column.Name, Name = Guid.NewGuid().ToString("N") });
                }
                foreach (object[,] block in result.Blocks)
                {
                    int rows = block.GetLength(0);
                    int columns = block.GetLength(1);
                    for (int row = 0; row < rows; row++)
                    {
                        object[] values = new object[columns];
                        for (int column = 0; column < columns; column++) values[column] = block[row, column];
                        _preview.Rows.Add(values);
                    }
                }
            }
            finally
            {
                _preview.ResumeLayout();
            }
        }

        private bool ValidateInputs(bool requireTarget, out ConnectionProfileSnapshot connection)
        {
            connection = _connectionSelector.SelectedProfile;
            string message = null;
            if (connection == null) message = "请选择连接方案。";
            else if (string.IsNullOrWhiteSpace(_queryName.Text)) message = "请输入查询方案名称。";
            else if (string.IsNullOrWhiteSpace(_sql.Text)) message = "请输入 SQL。";
            else if (requireTarget && string.IsNullOrWhiteSpace(_targetSheet.Text)) message = "请输入目标 Sheet 名称。";
            if (message == null) return true;
            MessageBox.Show(this, message, "SQL 查询", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
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

        private static Button ButtonOf(string text) => new Button { Text = text, AutoSize = true };
        private static Label LabelOf(string text) => new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Left };

        /// <summary>为查询方案下拉框包装模型和显示名称。</summary>
        private sealed class QueryItem
        {
            public QueryItem(QueryProfile profile) { Profile = profile; }
            public QueryProfile Profile { get; }
            public string DisplayName => Profile.Name;
        }
    }
}

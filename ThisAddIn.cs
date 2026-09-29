using System;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Excel;
using DB2Sheet.Infrastructure;
using DB2Sheet.Providers;
using DB2Sheet.Services;
using DB2Sheet.Storage;
using DB2Sheet.UI;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using OfficeCore = Microsoft.Office.Core;

namespace DB2Sheet
{
    /// <summary>Excel VSTO 加载项入口，负责生命周期、依赖组装、Ribbon 创建和顶层窗体管理。</summary>
    /// <remarks>
    /// 服务在首次启动或 Ribbon 操作前延迟初始化。SQL 查询和批量刷新窗体保持单实例；连接管理和设置使用模态临时窗体。
    /// </remarks>
    public partial class ThisAddIn
    {
        private ApplicationPaths _paths;
        private ILogger _logger;
        private IProviderRegistry _providers;
        private ISettingsRegistry _settingsRegistry;
        private ISettingsStore _settings;
        private IConnectionProfileRepository _connections;
        private IQueryProfileRepository _queries;
        private IDataSourceExecutionService _executionService;
        private IOperationRunner _operationRunner;
        private IQueryBufferService _queryBuffer;
        private IExcelResultWriter _resultWriter;
        private ISqlSheetTaskReader _taskReader;
        private IBatchRefreshService _batchRefresh;
        private QueryEditorForm _queryEditorForm;
        private SheetRefreshForm _sheetRefreshForm;

        /// <summary>Excel 加载插件时初始化服务。</summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            EnsureServices();
        }

        /// <summary>Excel 卸载插件时关闭仍然打开的非模态窗体。</summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            CloseForm(_queryEditorForm);
            CloseForm(_sheetRefreshForm);
            _queryEditorForm = null;
            _sheetRefreshForm = null;
        }

        /// <summary>创建 Office 请求的自定义 Ribbon 扩展对象。</summary>
        /// <returns>绑定到本加载项 UI 入口的 Ribbon 实例。</returns>
        protected override OfficeCore.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new DB2SheetRibbon(OpenQueryEditor, OpenSheetRefresh, OpenConnections, OpenSettings);
        }

        /// <summary>按依赖顺序创建路径、日志、提供程序、仓储和业务服务。</summary>
        /// <remarks>通过提供程序字段判断是否已初始化，因此重复调用不会重复注册服务。</remarks>
        private void EnsureServices()
        {
            if (_providers != null) return;

            _paths = new ApplicationPaths();
            _logger = new FileLogger(_paths);
            ProviderRegistry providers = new ProviderRegistry();
            providers.Register(new SqlServerProvider());
            providers.Register(new MySqlProvider());
            providers.Register(new PostgreSqlProvider());
            providers.Register(new SqliteProvider());
            _providers = providers;

            _settingsRegistry = SettingsRegistry.CreateDefault();
            _settings = new JsonSettingsStore(_paths, _settingsRegistry);
            _connections = new ConnectionProfileRepository(_paths);
            _queries = new QueryProfileRepository(_paths);
            _executionService = new DataSourceExecutionService(_providers, _logger);
            _operationRunner = new OperationRunner(_logger);
            _queryBuffer = new QueryBufferService(_executionService);
            _resultWriter = new ExcelResultWriter();
            _taskReader = new SqlSheetTaskReader();
            _batchRefresh = new BatchRefreshService(_queryBuffer, _resultWriter, _logger);
        }

        /// <summary>打开或激活 SQL 查询单实例窗体。</summary>
        private void OpenQueryEditor()
        {
            ExecuteUiAction(() =>
            {
                ExcelInterop.Workbook workbook = RequireActiveWorkbook();
                if (_queryEditorForm == null || _queryEditorForm.IsDisposed)
                {
                    _queryEditorForm = new QueryEditorForm(
                        _queries, _connections, _providers, _executionService, _queryBuffer,
                        _resultWriter, _operationRunner, _settings, workbook);
                    _queryEditorForm.FormClosed += (sender, args) => _queryEditorForm = null;
                    _queryEditorForm.Show();
                }
                else
                {
                    ActivateForm(_queryEditorForm);
                }
            });
        }

        /// <summary>打开或激活批量刷新单实例窗体。</summary>
        private void OpenSheetRefresh()
        {
            ExecuteUiAction(() =>
            {
                ExcelInterop.Workbook workbook = RequireActiveWorkbook();
                if (_sheetRefreshForm == null || _sheetRefreshForm.IsDisposed)
                {
                    _sheetRefreshForm = new SheetRefreshForm(
                        _connections, _providers, _executionService, _operationRunner, _settings,
                        _taskReader, _batchRefresh, workbook);
                    _sheetRefreshForm.FormClosed += (sender, args) => _sheetRefreshForm = null;
                    _sheetRefreshForm.Show();
                }
                else
                {
                    ActivateForm(_sheetRefreshForm);
                }
            });
        }

        /// <summary>以模态方式打开连接方案管理窗体。</summary>
        private void OpenConnections()
        {
            ExecuteUiAction(() =>
            {
                using (ConnectionProfilesForm form = new ConnectionProfilesForm(
                    _connections, _providers, _executionService, _operationRunner, _settings))
                {
                    form.ShowDialog();
                }
            });
        }

        /// <summary>以模态方式打开设置窗体。</summary>
        private void OpenSettings()
        {
            ExecuteUiAction(() =>
            {
                using (SettingsForm form = new SettingsForm(_settingsRegistry, _settings))
                {
                    form.ShowDialog();
                }
            });
        }

        /// <summary>统一执行 Ribbon UI 操作并记录、显示未处理异常。</summary>
        /// <param name="action">要在 Excel UI 线程执行的操作。</param>
        private void ExecuteUiAction(Action action)
        {
            try
            {
                EnsureServices();
                action();
            }
            catch (Exception exception)
            {
                _logger?.Write(LogSeverity.Error, "Ribbon 命令执行失败。", exception: exception);
                MessageBox.Show(exception.Message, AppPresentation.DisplayName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>获取当前活动工作簿。</summary>
        /// <returns>Excel 当前活动工作簿 COM 对象。</returns>
        /// <exception cref="InvalidOperationException">当前没有打开或创建工作簿。</exception>
        private ExcelInterop.Workbook RequireActiveWorkbook()
        {
            return Application.ActiveWorkbook ??
                throw new InvalidOperationException("请先打开或创建一个 Excel 工作簿。");
        }

        /// <summary>恢复、显示并激活已存在的非模态窗体。</summary>
        /// <param name="form">要激活的窗体。</param>
        private static void ActivateForm(Form form)
        {
            if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
            form.Show();
            form.Activate();
        }

        /// <summary>安全关闭尚未释放的窗体。</summary>
        /// <param name="form">可能为空或已释放的窗体。</param>
        private static void CloseForm(Form form)
        {
            if (form != null && !form.IsDisposed) form.Close();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

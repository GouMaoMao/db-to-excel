# DB2Sheet 架构说明

## 技术基础

- .NET Framework 4.7.2
- Excel VSTO Add-in
- Windows Forms
- Microsoft Office Interop Excel
- ADO.NET 数据库 Provider
- JSON 本地配置

## 分层结构

### Contracts

定义稳定边界，业务窗体和具体实现通过接口协作。

关键接口：

- `IDataSourceProvider`：通用数据源能力、连接参数、验证、测试和流式执行。
- `IDatabaseQueryProvider`：ADO.NET 数据库专用扩展。
- `IDataResultStream`：按块读取结果、取消和截断状态。
- `IDataSourceExecutionService`：统一启动查询与连接测试。
- `IOperationRunner`：显示独立进度窗体并管理取消令牌。
- `ISettingsRegistry` / `ISettingsStore`：动态设置定义与持久化。
- `IConnectionProfileRepository` / `IQueryProfileRepository`：连接和查询方案仓储。
- `ISqlSheetTaskReader`：从工作簿读取批量任务。
- `IExcelResultWriter`：将缓冲结果写入目标 Sheet。
- `IQueryBufferService`：消费流式结果并生成可供 UI 或 Excel 使用的缓冲结果。
- `IBatchRefreshService`：批量查询和串行 Excel 写入协调。

### Models

包含 Provider 能力、动态参数、连接快照、查询请求、查询方案、设置定义、进度、结果块和批量结果模型。

重要原则：

- 执行时使用不可变的 `ConnectionProfileSnapshot`，避免查询期间连接配置变化。
- 设置项由 `SettingDefinition<T>` 注册，不使用固定设置对象封死字段。
- `IDataResultStream` 和 `ResultBlock` 隔离具体数据库 Reader。

### Providers

- `DatabaseProviderBase`：请求校验、连接打开、只读会话、命令执行和流式 Reader 创建。
- `DatabaseResultStream`：顺序读取、分块、行数限制、截断判断、取消和资源释放。
- `SqlServerProvider`
- `MySqlProvider`
- `PostgreSqlProvider`
- `SqliteProvider`
- `DuckDbProvider`：本地 DuckDB 文件。打开时 `access_mode=READ_ONLY`。64 位 `duckdb.dll` 在构建时复制到插件旁边，随安装包发布。允许 `SELECT ... FROM read_parquet` 等只读表函数。
- `JdbcProvider`：通用 JDBC。Java 运行时由用户在「JDBC 环境」中检测或指定；厂商驱动 jar 写在连接方案中。转接程序嵌在插件内，首次使用时释放到本地数据目录，再由 Excel 进程外的 `java.exe` 启动。连接编辑填写完整 JDBC URL、驱动 jar、驱动类和凭据。

未来非数据库数据源应直接实现 `IDataSourceProvider`，不必继承 `DatabaseProviderBase`。JDBC 同样不继承 `DatabaseProviderBase`，因为它不是 ADO.NET 连接。

### Services

- `ProviderRegistry`：注册和解析 Provider。
- `SettingsRegistry`：注册核心及未来扩展设置。
- `SqlReadOnlyValidator`：去除注释和字面量后执行只读规则校验。
- `DataSourceExecutionService`：统一调用 Provider。查询日志写入 SQL 正文和哈希；文件日志对密码类键脱敏。
- `OperationRunner`：协调进度窗体、任务和取消。
- `QueryBufferService`：将流式结果消费为多个二维数组块。
- `BatchRefreshService`：串行或受限并行读取，并在 Excel UI 上下文串行写入。

### Storage

本地数据默认保存于：

`%LocalAppData%\DB2Sheet`

文件和目录包括：

- `connections.json`
- `settings.json`
- `jdbc-environment.json`
- `queries\`：每个查询方案一个文件。旧版 `queries.json` 只在首次加载时迁入该目录。
- `Logs\db2sheet-yyyyMMdd.log`

仓储采用临时文件和替换方式写入；损坏的 JSON 会被重命名备份。

当前连接密码仍保存在本地连接 JSON 中。生产发布前应评估 Windows DPAPI 或凭据管理器保护。

### Excel

- `SqlSheetTaskReader`：读取固定名称 `SQL` 的工作表，交给 `SqlSheetTaskParser` 解析。
- `SqlSheetHeader`：解析每列第 1 行的目标表名，以及 `//` 后的参数。当前会执行的键是「起始单元格」和「清空多余列」。
- `ExcelResultWriter`：查找或创建目标 Sheet，写一行标题和数据块。交互查询清空目标表已用区域并从 A1 写。批量刷新按起始单元格锚定，右侧多余列仅在参数要求时清空。

每列一个任务。第 2 行起的非空单元格拼成 SQL，前置注释留在 SQL 里并作为列表摘要。目标表名忽略大小写后必须唯一，否则整批不写。列协议的细则见 `.github/copilot-instructions.md`。

Excel COM 访问必须串行并位于 Excel UI 上下文。不要在后台线程直接操作 Workbook、Worksheet 或 Range。

### UI

- `OperationProgressForm`：通用进度和取消窗口。
- `DatabaseConnectionTree`：SQL 查询和批量刷新共用的连接树。
- `ConnectionProfileEditForm`：由 Provider 参数元数据动态生成字段。
- `ConnectionProfilesForm`：连接方案增删改。
- `SettingsForm`：由设置注册表动态生成编辑界面。
- `QueryEditorForm`：交互式 SQL 查询、方案复用、预览和导出。
- `SheetRefreshForm`：SQL Sheet 任务展示和批量刷新。
- `JdbcEnvironmentForm`：检测或指定本机 Java（搜索常见安装位置）。厂商驱动在连接编辑中选择。
- `DB2SheetRibbon`：Ribbon XML 和 Office 回调转发。查询组有「SQL 查询」「批量刷新」。管理组有「连接管理」，以及「设置」下拉里的「参数设置」「JDBC 环境」。关于组显示版本。窗体和按钮的细则见 `.github/copilot-instructions.md`。

## 应用组装

`ThisAddIn` 在启动时组装单例服务：

1. `ApplicationPaths` 和 `FileLogger`
2. SQL Server、MySQL、PostgreSQL、SQLite、DuckDB 和 JDBC 提供程序，以及 `ProviderRegistry`
3. 设置注册表和 JSON 设置存储
4. 连接与查询仓储
5. 执行服务、操作协调器和查询缓冲服务
6. SQL Sheet 读取器、Excel 写入器和批量刷新服务

功能区在 Excel 请求 Ribbon 时创建。业务窗体在用户点击对应按钮时创建。SQL 查询和批量刷新保持单实例；连接管理、设置和 JDBC 环境用模态窗体。

## 查询执行流程

1. UI 选择共享连接并生成 `DataSourceRequest`。
2. Provider 校验连接和只读 SQL。
3. ADO.NET 提供程序打开连接并配置只读会话，用 `SequentialAccess` 执行 Reader，由 `DatabaseResultStream` 按块读取。JDBC 不走 ADO.NET，由进程外的 `java.exe` 执行，结果经 `JdbcResultStream` 按块读取。
4. 预览经 `QueryBufferService` 缓冲后绑定到 DataGridView，或由 Excel 写入器分块写入 Sheet。
5. 进度通过 `IProgress<OperationProgress>` 上报；取消令牌触发命令取消。

## 批量刷新流程

1. `SqlSheetTaskReader` 扫描 `SQL` 工作表的每个已用列。
2. 第 1 行是目标表名和 `//` 参数，第 2 行起的非空单元格拼成 SQL。
3. `SheetRefreshForm` 统一选择一个连接并读取批量设置。连接不写在 SQL 页里。
4. 目标表名忽略大小写后有重复时，刷新开始前抛出异常，不写入任何表。
5. 串行模式逐任务查询和写入。
6. 并行模式仅并行数据库读取，且受最大并发数限制。
7. Excel 写入始终回到启动刷新时的 UI 同步上下文并串行执行。每张目标表写入前的跳转和重画见 `.github/copilot-instructions.md` 的「批量刷新」。单个任务失败不终止其余任务；用户取消会终止整批。

## 安全边界

客户端只读校验不能代替数据库权限。推荐同时使用：

- 数据库只读账号。
- 最小权限和网络访问控制。
- Provider 连接只读意图或会话只读设置。
- 客户端 SQL 词法校验。
- 查询日志写入 SQL 正文和哈希。文件日志对密码、令牌、连接字符串一类键脱敏，不把这条脱敏当成可以主动记录密码。

## 扩展原则

新增数据源时：

1. 优先实现 `IDataSourceProvider`。
2. 通过 `ConnectionParameters` 描述 UI 字段，不在窗体中写死参数。
3. 返回 `IDataResultStream`，不要让 UI 依赖具体响应格式。
4. 在 `ThisAddIn` 的 Provider 注册处注册实现。
5. 增加 Provider 特定测试和只读策略说明。

新增 JDBC 数据源时不要把 JDK 或厂商驱动打进安装包。用户通过「JDBC 环境」检测或指定 `java.exe`，并在连接方案中选择驱动 jar 与驱动类。转接程序嵌在 DLL 内，首次使用时释放到 `%LocalAppData%\DB2Sheet\db2sheet-jdbc-bridge.jar`。

新增设置时：

1. 创建新的 `SettingDefinition<T>`。
2. 在设置注册表中注册。
3. 使用 `ISettingsStore` 读取。
4. 不向 `SettingsForm` 添加特定字段；界面应继续由元数据生成。

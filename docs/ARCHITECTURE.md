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
- `DataSourceExecutionService`：统一调用 Provider 并写入脱敏日志。
- `OperationRunner`：协调进度窗体、任务和取消。
- `QueryBufferService`：将流式结果消费为多个二维数组块。
- `BatchRefreshService`：串行或受限并行读取，并在 Excel UI 上下文串行写入。

### Storage

本地数据默认保存于：

`%LocalAppData%\DB2Sheet`

文件包括：

- `connections.json`
- `queries.json`
- `settings.json`
- `Logs\db2sheet-yyyyMMdd.log`

仓储采用临时文件和替换方式写入；损坏的 JSON 会被重命名备份。

当前连接密码仍保存在本地连接 JSON 中。生产发布前应评估 Windows DPAPI 或凭据管理器保护。

### Excel

- `SqlSheetTaskReader`：读取固定名称 `SQL` 的工作表。
- `ExcelResultWriter`：查找或创建目标 Sheet、清空旧内容、写标题和数据块。

Excel COM 访问必须串行并位于 Excel UI 上下文。不要在后台线程直接操作 Workbook、Worksheet 或 Range。

### UI

- `OperationProgressForm`：通用进度和取消窗口。
- `ConnectionProfileSelector`：两个业务窗体复用的共享连接选择器。
- `ConnectionProfileEditForm`：由 Provider 参数元数据动态生成字段。
- `ConnectionProfilesForm`：连接方案增删改。
- `SettingsForm`：由设置注册表动态生成编辑界面。
- `QueryEditorForm`：交互式 SQL 查询、方案复用、预览和导出。
- `SheetRefreshForm`：SQL Sheet 任务展示和批量刷新。
- `JdbcEnvironmentForm`：检测或指定本机 Java（搜索常见安装位置）。厂商驱动在连接编辑中选择。
- `DB2SheetRibbon`：Ribbon XML 和 Office 回调转发。

## 应用组装

`ThisAddIn` 在启动时组装单例服务：

1. `ApplicationPaths` 和 `FileLogger`
2. SQL Server、MySQL、PostgreSQL、SQLite、DuckDB 和 JDBC 提供程序，以及 `ProviderRegistry`
3. 设置注册表和 JSON 设置存储
4. 连接与查询仓储
5. 执行服务、操作协调器和查询缓冲服务
6. SQL Sheet 读取器、Excel 写入器和批量刷新服务
7. Ribbon 回调和业务窗体

## 查询执行流程

1. UI 选择共享连接并生成 `DataSourceRequest`。
2. Provider 校验连接和只读 SQL。
3. 打开连接并配置 Provider 特定只读会话。
4. 使用 `SequentialAccess` 执行 Reader。
5. `DatabaseResultStream` 按设置的块大小读取。
6. 预览绑定到 DataGridView，或由 Excel 写入器分块写入 Sheet。
7. 进度通过 `IProgress<OperationProgress>` 上报；取消令牌触发命令取消。

## 批量刷新流程

1. `SqlSheetTaskReader` 扫描 `SQL` 工作表的每个已用列。
2. 第 1 行作为目标 Sheet，第 2 行起拼接为 SQL。
3. `SheetRefreshForm` 统一选择连接并读取批量设置。
4. 串行模式逐任务查询和写入。
5. 并行模式仅并行数据库读取，且受最大并发数限制。
6. Excel 写入始终回到启动刷新时的 UI 同步上下文并串行执行。
7. 重复目标 Sheet 的后续任务跳过。

## 安全边界

客户端只读校验不能代替数据库权限。推荐同时使用：

- 数据库只读账号。
- 最小权限和网络访问控制。
- Provider 连接只读意图或会话只读设置。
- 客户端 SQL 词法校验。
- 日志脱敏，不记录完整凭据或完整 SQL。

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

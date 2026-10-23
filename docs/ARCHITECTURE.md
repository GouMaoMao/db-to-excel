# DB2Sheet 架构说明

## 技术基础

- .NET Framework 4.7.2
- Excel VSTO Add-in。宿主是 Excel 与 WPS 表格。
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
- `ai-models.json`：模型配置。密钥是当前 Windows 用户的 DPAPI 密文。测试通过后缓存该接口返回的型号 id。
- `ai-thinking-mappings.json`：用户可编辑的思考强度 API 参数映射，按协议、接口主机和可选型号规则匹配。首次运行时从程序集默认配置生成。
- `ai-skills\`：一份 Skill 一个文件，只作上下文。
- `ai-conversations\`：一条对话一个文件。除可见问答外，还存本轮 token 计数，以及查看请求详情用的分类原文和这次调用的地址、协议、模型、发送时间、HTTP 状态与耗时。
- `webview2\`：对话页 WebView2 的浏览器数据。
- `Logs\db2sheet-yyyyMMdd.log`
- `Logs\ai-audit-yyyyMMdd.log`：呼叫模型前写入的审计记录。写失败则不发送。删除对话不删除这些记录。

仓储采用临时文件和替换方式写入；损坏的 JSON 会被重命名备份。

连接密码明文写在 `connections.json`。模型密钥不沿用这种做法。保护方式见 `docs/DEPLOYMENT.md` 的「安全要求」。AI 对话的界面和发送边界见 `.github/copilot-instructions.md` 的「AI 助手」。发给模型的说明在嵌入的 `Prompts/*.md`，代码只填占位符。

### Excel

- `SqlSheetTaskReader`：读取固定名称 `SQL` 的工作表，交给 `SqlSheetTaskParser` 解析。
- `SqlSheetHeader`：解析每列第 1 行的目标表名，以及 `//` 后的参数。
- `ExcelResultWriter`：查找或创建目标 Sheet，写一行标题和数据块。

列协议、写入区域和格式保留的细则见 `.github/copilot-instructions.md`。

Excel COM 访问必须串行并位于 Excel UI 上下文。不要在后台线程直接操作 Workbook、Worksheet 或 Range。

### UI

- `OperationProgressForm`：通用进度和取消窗口。
- `DatabaseConnectionTree`：SQL 查询和批量刷新共用的连接树。
- `ConnectionProfileEditForm`：由 Provider 参数元数据动态生成字段。
- `ConnectionProfilesForm`：连接方案增删改。
- `SettingsForm`：由设置注册表动态生成编辑界面。
- `QueryEditorForm`：交互式 SQL 查询。编辑器是 AvalonEdit，语法高亮使用其自带的 T-SQL 定义；SQL 页正文美化使用同一份定义。
- `SheetRefreshForm`：SQL Sheet 任务展示和批量刷新。
- `JdbcEnvironmentForm`：检测或指定本机 Java（搜索常见安装位置）。厂商驱动在连接编辑中选择。
- `LocalDataForm`：只列出连接、查询、Skill、AI 思考强度映射等本机路径，不改保存位置。
- `DB2SheetRibbon`：Ribbon XML 和 Office 回调转发。查询组有「SQL 查询」「批量刷新」。管理组有「连接管理」，以及「设置」下拉里的「参数设置」、「AI 设置」（模型设置、Skill 设置）、「JDBC 环境」和「本地数据」。关于组显示版本。
- `AiModelSettingsForm` / `AiSkillSettingsForm`：模型与 Skill。模型列表和对话模型菜单按配置显示子型号；模型连接测试通过 `OperationRunner` 显示阶段进度。协议和发送边界见 `.github/copilot-instructions.md` 的「AI 助手」。
- `AiChatControl`：SQL 查询窗右侧的对话，以及报错时弹出的 `AiChatForm`。正文用 WebView2 显示。只提问，不执行 SQL。共用控件根据 Skill 勾选状态切换底部说明，不拦截发送；交互细则见 `.github/copilot-instructions.md` 的「AI 助手」。正文页只向宿主发 `pointerdown` 一种消息，宿主借此通过 `DropDownCloser` 收起本线程的下拉菜单和 @ 引用弹窗，因为点在 WebView2 里 WinForms 收不到。
- `AiThinkingMappingStore` / `AiChatClient`：从用户级 JSON 映射读取所选思考强度对应的请求参数，并将非消息参数摘要随对话保存供请求详情展示。自动档不依赖映射文件；手动档位的映射错误会阻止本次发送。具体行为见 `.github/copilot-instructions.md` 的「AI 助手」。

窗体和按钮的细则见 `.github/copilot-instructions.md`。

## 应用组装

`ThisAddIn` 在启动时组装单例服务：

1. `ApplicationPaths` 和 `FileLogger`
2. SQL Server、MySQL、PostgreSQL、SQLite、DuckDB 和 JDBC 提供程序，以及 `ProviderRegistry`
3. 设置注册表和 JSON 设置存储
4. 连接与查询仓储，以及 AI 的模型、Skill、对话存储和审计日志
5. 执行服务、操作协调器和查询缓冲服务
6. SQL Sheet 读取器、Excel 写入器和批量刷新服务

功能区在 Excel 请求 Ribbon 时创建。业务窗体在用户点击对应按钮时创建。SQL 查询和批量刷新保持单实例；连接管理、设置、模型设置、Skill 设置、JDBC 环境和本地数据用模态窗体。这两个单实例窗体第一次打开时，布局完成前不显示。打开方式见 `.github/copilot-instructions.md` 的「界面」。报错时的 AI 对话是查询窗拥有的非模态窗。

## 查询执行流程

1. UI 选择共享连接并生成 `DataSourceRequest`。
2. Provider 校验连接和只读 SQL。
3. ADO.NET 提供程序打开连接并配置只读会话，用 `SequentialAccess` 执行 Reader，由 `DatabaseResultStream` 按块读取。JDBC 不走 ADO.NET，由进程外的 `java.exe` 执行，结果经 `JdbcResultStream` 按块读取。
4. 预览经 `QueryBufferService` 缓冲后绑定到 DataGridView。预览右键可复制，或把已缓冲的预览存成 xlsx。写入 Sheet 和把 SQL 写入 SQL 页的细则见 `.github/copilot-instructions.md`。
5. 进度通过 `IProgress<OperationProgress>` 上报；取消令牌触发命令取消。

## 批量刷新流程

1. `SqlSheetTaskReader` 扫描 `SQL` 工作表的每个已用列，交给解析器得到任务。
2. `SheetRefreshForm` 统一选择一个连接并读取批量设置。
3. 串行模式逐任务查询和写入。并行模式仅并行数据库读取，且受最大并发数限制。
4. Excel 写入回到启动刷新时的 UI 同步上下文并串行执行。

列协议、重名、单任务失败和写入前跳转的细则见 `.github/copilot-instructions.md` 的「批量刷新」。

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

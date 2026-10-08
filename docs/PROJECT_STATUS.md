# DB2Sheet 项目状态

## 项目目标

DB2Sheet 是一个基于 .NET Framework 4.7.2 的 Excel VSTO 加载项。它面向只读数据查询场景，允许用户维护共享数据源连接、编写和复用 SQL、在 Windows Forms 中预览结果，并将结果分块写入 Excel Sheet。

设计目标包括：

- 分层、可读、可维护。
- 数据源 Provider 可持续扩展，不限定为 ODBC 或 ADO.NET；未来可加入 REST API 等数据源。
- 查询仅允许只读操作，并结合 SQL 校验、只读连接或会话、数据库只读账号进行纵深防护。
- 所有耗时操作支持进度展示和取消，避免 Excel 界面长时间无响应。
- 两个 Ribbon 业务入口分别对应交互式 SQL 查询和 SQL Sheet 批量刷新。
- 连接方案由两个业务窗体共享。
- 设置由注册表动态驱动，不能把未来参数限定在固定模型中。
- 不使用 VSTO ClickOnce；最终由 Advanced Installer 封装部署。

## 当前进度

功能主路径已在代码中实现：连接、只读查询、预览、写入 Excel、SQL 页批量刷新，以及 JDBC。

真实 Excel 集成测试尚未记为完成。解决方案引用 `DB2Sheet.Tests`，该目录不入库，当前工作区没有这份源码。本文不记录测试通过个数。

## 已完成

- 分层目录：`Contracts`、`Models`、`Providers`、`Services`、`Storage`、`Infrastructure`、`Excel`、`UI`。
- 通用数据源 Provider、动态参数、流式结果和可取消执行契约。
- SQL Server、MySQL、PostgreSQL、SQLite、DuckDB Provider。DuckDB 为本地只读文件，64 位本机库随插件输出发布。
- 通用 JDBC Provider。Java 运行时由「JDBC 环境」检测或指定；厂商驱动 jar 写在连接方案中。连接编辑填写完整 JDBC URL。
- Ribbon XML 动态加载，提供“SQL 查询”“批量刷新”“连接管理”，以及「设置」下拉（「参数设置」「JDBC 环境」）。
- 通用及 Provider 特定只读 SQL 校验。JDBC 允许查询前的 `SET` 语句。
- MySQL、PostgreSQL、SQLite 只读会话；SQL Server 使用只读连接意图；DuckDB 在打开连接时使用只读访问模式。
- ADO.NET 用 `DbDataReader` 流式读取；JDBC 用 `JdbcResultStream`。两者都按块读取，并做行数限制和截断检测。
- 独立通用进度窗体和统一操作执行协调器。
- 共享连接方案仓储及 JSON 持久化。
- 查询方案仓储，支持保存、覆盖、删除和复用。每个方案一个文件，放在 `queries\`；旧版 `queries.json` 在首次加载时迁移。
- 可扩展设置注册表、未知键保留、动态设置界面和批量保存。
- 滚动文件日志和关联 ID。查询日志写入 SQL 正文和哈希；文件日志对密码类键脱敏。
- 动态连接参数编辑、连接测试、连接管理和共享连接选择控件。
- `QueryEditorForm`：SQL 编辑、方案复用、结果预览、写入目标 Sheet。
- `SheetRefreshForm`：扫描 SQL Sheet、展示任务、统一连接选择、批量刷新。
- SQL 页协议的实现见 `Excel/SqlSheetTaskParser` 与 `Excel/SqlSheetHeader`。细则见 `.github/copilot-instructions.md`。
- Excel 写入：交互查询清空目标表已用区域并从 A1 写；批量刷新按起始单元格锚定，多余列仅在参数要求时清空。
- 批量串行或受限并行数据库读取；Excel COM 写入保持串行。
- 目标表名忽略大小写后重复时，整批失败，不写入任何表。单个任务失败不终止其余任务。
- `ThisAddIn` 在启动时组装服务。功能区在 Excel 请求 Ribbon 时创建，业务窗体在用户点击时创建。
- Office PIA 引用已通过 `UseOfficeInterop` 启用，Ribbon 相关代码在主项目中。
- Ribbon 关于组显示版本。

## 尚未完成

1. 在真实 Excel 调试宿主中验证 Ribbon：SQL 查询、批量刷新、连接管理、参数设置、JDBC 环境、版本。
2. 使用实际数据库验证连接测试、预览、导出和批量刷新。JDBC 需本机 Java 和连接方案中的驱动 jar。
3. 验证取消、超时、截断、空结果、大结果和并行刷新。
4. 验证工作簿关闭、切换活动工作簿和窗体长期打开等生命周期场景。
5. 根据运行结果修复 COM 释放、线程切换或 Provider 兼容问题。
6. 在批量调度与 Excel COM 写入边界解耦后，扩展串并行、失败隔离和取消自动化测试。
7. 部署上尚未决定的项见 `docs/DEPLOYMENT.md`：代码签名、Office 最低版本、x86/x64、安装范围，以及卸载时是否保留用户配置。连接密码仍明文保存在 `connections.json`。

## 关键业务规则

- 只允许只读查询；数据库侧仍应优先使用只读账号。
- `SQL` 工作表名称固定为 `SQL`。
- SQL 页每列一个任务。第 1 行是目标表名和 `//` 参数，第 2 行起的非空单元格拼成 SQL。细则见 `.github/copilot-instructions.md`。
- 批量刷新窗体统一选择一个共享连接，不在 SQL Sheet 中保存连接信息。
- 预览上限、导出上限、查询超时、结果块大小、串并行模式和最大并发数均来自动态设置。
- Ribbon 使用 Ribbon XML，不使用可视化 Ribbon Designer。
- 部署使用 Advanced Installer，不使用 ClickOnce。

## 继续开发时的建议顺序

1. 读取 `.github/copilot-instructions.md` 和本文件。
2. 阅读 `docs/ARCHITECTURE.md`，确认边界和既有设计决策。
3. 阅读 `docs/TESTING.md`，执行必须人工完成的 Excel 集成测试。
4. 根据集成测试结果评估批量调度与 Excel COM 写入边界解耦。
5. 修复集成测试发现的问题，并按约定正文「文档」一节同步更新受影响的文档。
6. 发布使用 `scripts/publish-release.ps1`。安装工程在本机，不入库。签名和位数等待决定项见 `docs/DEPLOYMENT.md`。

## 文档维护规则

- 进度描述跟当前代码走。完成一个阶段时，同步更新「当前进度」「已完成」和「尚未完成」。
- 行为细则以 `.github/copilot-instructions.md` 为准。本文件只记进度，不另写一版界面、SQL 页或批量刷新规则。
- 改行为时，按约定正文「文档」一节，同一改动里更新受影响的 `docs` 文件。

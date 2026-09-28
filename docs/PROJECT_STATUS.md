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

**总体进度：约 90%**

当前处于“核心纯逻辑自动化测试完成，进入真实 Excel 集成测试和稳定化”阶段。

最近一次验证：Debug / Release 构建成功；MSTest v2 自动化测试 41 个全部通过，0 个失败。

## 已完成

- 分层目录：`Contracts`、`Models`、`Providers`、`Services`、`Storage`、`Infrastructure`、`Excel`、`UI`。
- 通用数据源 Provider、动态参数、流式结果和可取消执行契约。
- SQL Server、MySQL、PostgreSQL、SQLite Provider。
- 通用及 Provider 特定只读 SQL 校验。
- MySQL、PostgreSQL、SQLite 只读会话；SQL Server 使用只读连接意图。
- `DbDataReader` 流式读取、结果分块、行数限制和截断检测。
- 独立通用进度窗体和统一操作执行协调器。
- 共享连接方案仓储及 JSON 持久化。
- 查询方案仓储，支持保存、覆盖、删除和复用。
- 可扩展设置注册表、未知键保留、动态设置界面和批量保存。
- 滚动文件日志、关联 ID 和敏感信息脱敏。
- 动态连接参数编辑、连接测试、连接管理和共享连接选择控件。
- `QueryEditorForm`：SQL 编辑、方案复用、结果预览、写入目标 Sheet。
- `SheetRefreshForm`：扫描 SQL Sheet、展示任务、统一连接选择、批量刷新。
- SQL Sheet 协议：每列一个任务，第 1 行为目标 Sheet 名，第 2 行起为 SQL。
- Excel 目标 Sheet 创建、旧内容清理、标题写入和结果分块写入。
- 批量串行或受限并行数据库读取；Excel COM 写入保持串行。
- 同一批次重复目标 Sheet 的后续任务跳过。
- Ribbon XML 动态加载，提供“SQL 查询”“批量刷新”“连接管理”“设置”四个入口。
- `ThisAddIn` 中的服务组装和业务窗体生命周期管理。
- Office PIA 引用已通过 `UseOfficeInterop` 启用，Ribbon 相关代码编译通过。
- 已建立 `DB2Sheet.Tests`（.NET Framework 4.7.2、MSTest v2）自动化测试项目。
- SQL 只读校验、核心设置注册与定义共 15 个测试通过。
- SQL Sheet 二维单元格解析已从 Excel COM 读取中提取，4 个纯逻辑测试通过。
- JSON 设置存储 8 个测试通过，覆盖持久化、批量校验、未知键保留、事件和损坏文件恢复。
- 连接方案与查询方案仓储共 14 个测试通过，覆盖持久化、复制隔离、排序、覆盖、删除、校验和损坏文件恢复。

## 尚未完成

1. 在真实 Excel 调试宿主中验证 Ribbon 加载及四个回调。
2. 使用实际数据库验证连接测试、预览、导出和批量刷新。
3. 验证取消、超时、截断、空结果、大结果和并行刷新。
4. 验证工作簿关闭、切换活动工作簿和窗体长期打开等生命周期场景。
5. 根据运行结果修复 COM 释放、线程切换或 Provider 兼容问题。
6. 在批量调度与 Excel COM 写入边界解耦后，扩展串并行、失败隔离和取消自动化测试。
7. 完善使用、安全、Provider 扩展和 Advanced Installer 部署文档。

## 关键业务规则

- 只允许只读查询；数据库侧仍应优先使用只读账号。
- `SQL` 工作表名称固定为 `SQL`。
- SQL Sheet 每列一个任务：第 1 行为目标 Sheet，第 2 行至最后使用行的非空内容按换行拼接为 SQL。
- 批量刷新窗体统一选择一个共享连接，不在 SQL Sheet 中保存连接信息。
- 预览上限、导出上限、查询超时、结果块大小、串并行模式和最大并发数均来自动态设置。
- Ribbon 使用 Ribbon XML，不使用可视化 Ribbon Designer。
- 部署使用 Advanced Installer，不使用 ClickOnce。

## 继续开发时的建议顺序

1. 读取 `.github/copilot-instructions.md` 和本文件。
2. 阅读 `docs/ARCHITECTURE.md`，确认边界和既有设计决策。
3. 阅读 `docs/TESTING.md`，执行必须人工完成的 Excel 集成测试。
4. 根据集成测试结果评估批量调度与 Excel COM 写入边界解耦。
5. 修复集成测试发现的问题，并同步更新本文件。
6. 最后完善 `docs/DEPLOYMENT.md` 并制作 Advanced Installer 安装包。

## 文档维护规则

- 本文件是项目长期进度和待办事项的权威来源。
- 每完成一个阶段，应同步更新“当前进度”“已完成”“尚未完成”和最近验证结果。
- `C:\Users\YQSL\.copilot\plans\plan-db2sheet-excel.md` 仅作为本机 Copilot 工作计划，不作为项目权威记录。

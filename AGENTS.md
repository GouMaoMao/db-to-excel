# 项目约定

本文件是所有编码 Agent 的入口和强制规则，供不同工具共用。功能细则只在 `docs/specs/`，同一条规则只写一处；各工具专用入口只引用本文件，不得复制约定正文。用户主目录里的旧计划和 Agent 记忆只作历史记录；发生冲突时，以本文件、`docs/specs/` 和当前代码为准。

## 读取顺序

- 开始工作前先读本文件，再读 `docs/HANDOFF.md`，并检查当前代码、`git status` 和相关差异，不得只依据交接记录推断工作区状态。
- 动手改文件前按「规格索引」读完命中的规格，命中多份就都读。规格与代码冲突时以代码为准，并在同一改动里改规格。
- 结束任务前更新 `docs/HANDOFF.md`，写明目标、完成情况、改动文件、验证结果、阻塞项和明确的下一步。
- 发布或打安装包时还必须读取 `.cursor/skills/publish-release/SKILL.md`；该文件虽位于工具专用目录，其发布规则适用于所有 Agent。
- 最终回复用一行列出本轮读过的规格，格式为「已读规格：`ui-common.md`、`batch-refresh.md`」。没有命中任何规格时这一行不写。

## 规格索引

| 触发条件 | 规格文件 |
| --- | --- |
| 窗体外观、通用控件、连接树显示；改动 `UI/` 下任何文件或新增 `AppForm` 派生窗体 | `docs/specs/ui-common.md` |
| SQL 查询窗、查询页签、查询方案、连接树选中状态；`UI/QueryEditorForm.cs`、`UI/SqlEditorControl.cs`、`Services/QuerySchemeNote.cs` | `docs/specs/sql-query-form.md` |
| 批量刷新、美化格式、状态圆点、刷新进度与结束说明；`UI/SheetRefreshForm.cs`、`Services/BatchRefreshService.cs`、`Excel/SqlSheetTaskReader.cs`、`Excel/SqlSheetTaskParser.cs` | `docs/specs/batch-refresh.md` |
| 预览、导出、写入 Sheet、Excel 结果写入；`Excel/ExcelResultWriter.cs`、`Services/QueryBufferService.cs`、`UI/QueryEditorForm.cs` 的预览与写入路径 | `docs/specs/preview-and-write.md` |
| AI 对话、模型设置、Skill 设置、思考强度、请求与审计；`UI/AiChat*.cs`、`UI/AiModelSettingsForm.cs`、`UI/AiSkillSettingsForm.cs`、`Services/Ai*.cs`、`Storage/Ai*.cs`、`Prompts/*.md` | `docs/specs/ai-assistant.md` |

## 技术边界

- .NET Framework 4.7.2，C# 7.3。不要使用更高版本的语法。
- 旧式 csproj。新增 UI 等源文件必须在 `DB2Sheet.csproj` 里显式 `Compile Include`。
- 发布不用 VSTO ClickOnce，由 Advanced Installer 封装。调试时 Excel 仍只复制 `DB2Sheet.dll.manifest` 里列出的文件。
- 用户可见名称、窗口标题、界面字体和代码字体只改 `UI/AppPresentation.cs`。程序集名、命名空间、Ribbon 控件 ID、持久化目录是独立技术标识。
- JDBC 是通用数据源。连接编辑填写完整 JDBC URL、驱动 jar、驱动类和凭据。不要把 JDK 或厂商驱动打进安装包。
- `db2sheet-jdbc-bridge.jar` 嵌在 `DB2Sheet.dll` 内，首次使用释放到 `%LocalAppData%\DB2Sheet`。VSTO 加载时插件 DLL 旁边没有这个文件。
- `duckdb.dll` 写入 VSTO 清单，按 `Assembly.CodeBase` 的插件原始目录加载。Excel 会把插件复制到临时目录，那个目录里没有这份原生库。
- 「JDBC 环境」只检测本机 Java、指定 java.exe，或打开 JDK 下载页。驱动 jar 写在连接方案里。插件不下载 JDK。
- 用户操作失败不得只写状态栏、吞掉异常或空 `catch` 后继续；须弹出错误详情窗（含异常类型、消息、堆栈与 InnerException，可用 `ExceptionDetailForm`）。仅清理、取消、资源释放等非用户主路径可故意吞异常，并在代码注释里写明原因。
- 改代码不要用打补丁叠在旧逻辑上。先从现有功能和架构看，能不能和已有需求合在一起，或直接改设计。先问清需求方要解决的是什么。若看起来只能打补丁，先反问需求方这个需求是否真有必要，得到确认后再做。
- 在保证功能实现的前提下，以精简代码为荣，以重复代码为耻。能并进已有逻辑就不要再写一份。判断精简和重复时不计注释；中文 XML 文档注释仍按「代码文档」章节编写。
- 调用会改动环境的命令时，命令体开头用 PowerShell 块注释写明用途和风险，格式为 `<# 用途｜风险 #>`，用途一句话说清这条命令要达成什么。只读命令不写，命令体不是 PowerShell 时不写。

## 操作说明

- 给用户跟着做的步骤说明（含计划）里，步骤用正文按顺序写成完整句子，一段一件事。原因、名词、限制和可以不做的事项不写进步骤句子，紧跟在该段后面，用 Markdown 引用块单独标出。代码注释按「代码文档」，提交说明按「提交说明」，界面文案按 `docs/specs/ui-common.md` 的悬浮说明规则。

## 代码文档

- 新增或修改的非生成代码必须同步编写中文 XML 文档注释。不要手改 `*.Designer.cs` 等生成文件。
- 每个类、接口、枚举和结构用 `<summary>` 说明职责和设计目的。
- 关键公开成员按需要使用 `<summary>`、`<param>`、`<returns>`、`<remarks>` 和 `<exception>`。
- 涉及文件、数据库、Excel COM、异步任务或 UI 时，写明副作用、线程要求、取消行为、资源所有权和释放责任。
- 注释解释设计意图，不要只复述语句。

## 提交说明

- 使用约定式提交。类型用英文小写，摘要用中文，冒号后空一格：`feat: 用一句话说明为什么改`。
- 类型按改动性质选择：`feat` 新功能，`fix` 缺陷修复，`docs` 文档，`style` 格式，`refactor` 重构，`perf` 性能，`test` 测试，`chore` 构建或杂项。
- 标题写原因。需要补充影响或限制时，另起一行写正文。
- 一次改动若覆盖多项需求或功能，能拆就拆成多次提交，每次只含一类改动，方便 review。不要把互不相关的功能、修复和杂项塞进同一个提交。

## 设置

- 键名以 `session.` 开头的项是会话状态，不出现在设置界面。
- 新增或删除注册项时，同步修改 `DB2Sheet.Tests/Services/SettingsRegistryTests.cs` 里的注册数量断言。
- 功能区「设置」里有「本地数据」。只列出连接、查询、参数、JDBC 环境、模型、AI 思考强度映射、Skill、对话、日志、对话页缓存和 JDBC 转接程序的本机路径，不改保存位置。可以复制路径，或在资源管理器中打开已经存在的文件和目录。尚未生成时说明还没有这个文件。

## 文档

- `docs` 根目录描述当前代码的结构和流程，`docs/specs/` 放功能规格。改查询、批量刷新、存储、Ribbon 或部署脚本的行为时，同一改动里更新受影响的文档或规格。
- 文档、规格与代码冲突时，以当前代码为准，并在同一改动里改文档或规格。过时描述不留在文档里当规格。
- 界面、SQL 页和批量刷新的细则只写在 `docs/specs/`，`docs` 根和其他文档用短句指向那里，不另写一版。
- 本文件只放强制规则和索引。超过约 15 KB 就把功能细则移回 `docs/specs/`，不再往本文件加功能描述。
- `DB2Sheet.Tests/` 不入库。文档不写无法在当前工作区核对的测试通过个数。

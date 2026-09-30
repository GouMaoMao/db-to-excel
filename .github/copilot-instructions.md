# 项目约定

本文件是仓库里唯一的 AI 约定正文。用户主目录 `~\.copilot\plans` 里的旧计划只作历史记录，与本文件冲突时以本文件和当前代码为准。

## 技术边界

- .NET Framework 4.7.2，C# 7.3。不要使用更高版本的语法。
- 旧式 csproj。新增 UI 等源文件必须在 `DB2Sheet.csproj` 里显式 `Compile Include`。
- 不要修改 `DB2Sheet1.csproj`。
- 不使用 VSTO ClickOnce。部署由 Advanced Installer 封装。
- 用户可见名称、窗口标题、界面字体和代码字体只改 `UI/AppPresentation.cs`。程序集名、命名空间、Ribbon 控件 ID、持久化目录是独立技术标识。
- JDBC 是通用数据源。连接编辑填写完整 JDBC URL、驱动 jar、驱动类和凭据。不要把 JDK 或厂商驱动打进安装包。
- `db2sheet-jdbc-bridge.jar` 嵌在 `DB2Sheet.dll` 内，首次使用释放到 `%LocalAppData%\DB2Sheet`。不要改回只在插件 DLL 旁边查找；Excel 经 VSTO 加载时那个目录里没有这个文件。
- 「JDBC 环境」只管本机 Java：点「检测 Java」搜索本机，或浏览指定 java.exe；也可打开 JDK 下载页自行安装。厂商驱动 jar 写在连接方案里，不要再做集中驱动文件夹。不要让用户去下载转接程序。不要在插件内一键下载 JDK。
- 用户操作失败不得只写状态栏、吞掉异常或空 `catch` 后继续；须弹出错误详情窗（含异常类型、消息、堆栈与 InnerException，可用 `ExceptionDetailForm`）。仅清理、取消、资源释放等非用户主路径可故意吞异常，并在代码注释里写明原因。

## 代码文档

- 新增或修改的非生成代码必须同步编写中文 XML 文档注释。不要手改 `*.Designer.cs` 等生成文件。
- 每个类、接口、枚举和结构用 `<summary>` 说明职责和设计目的。
- 关键公开成员按需要使用 `<summary>`、`<param>`、`<returns>`、`<remarks>` 和 `<exception>`。
- 涉及文件、数据库、Excel COM、异步任务或 UI 时，写明副作用、线程要求、取消行为、资源所有权和释放责任。
- 注释解释设计意图，不要只复述语句。

## 界面

- 所有应用窗体显示嵌入的 rocket.ico，不用 WinForms 默认图标。
- 界面字体用 Microsoft YaHei UI。进度日志用 Consolas。
- SQL 编辑字体用 Microsoft YaHei UI Light，使中英文同一字面高度；缺失时回退等线 Light，再回退微软雅黑 UI。不要改回 Cascadia 或 Consolas。
- 进度窗口在任务成功后自动关闭；只有出错时保留，供用户查看。
- 查询方案信息显示在左下角的查询方案区域，不要放进底部状态栏。
- 功能窗体相对工作区居中，使用 `FormStartPosition.Manual`，不要用 `CenterScreen`。
- 五个功能窗体记住宽高和最大化状态，键为 `session.formBounds`：SQL 查询、批量刷新、连接管理、设置、JDBC 环境。进度、文本输入、连接编辑这些小窗不记。

## SQL 查询窗体

- SQL 区标题为「SQL编辑器（仅限查询）」。左栏标题为「数据库连接」。
- 只有一个菜单，三项为连接、查询、编辑。
- 当前连接根节点显示 `名称 [当前]`。名称加粗，`[当前]` 用非粗体灰色与名称区分；未选中时名称行使用浅色高亮底 RGB(255,243,205)，图标保持彩色。用 `session.activeConnectionId` 记住，打开时直接显示，不要为了显示当前连接而重新测试。
- 没有结果时预览区保持折叠。展开后编辑区与预览区高度为 1:1。
- 左右默认 1:5，即可用宽度的 1/6。左栏最小宽度 140，右栏 420。
- 只有用户拖动分割条时，才把千分比写入 `session.workspaceSplitRatio`。缩放、最大化、还原只按已存比例重算，不要把这次自动调整写回设置。值为 0 表示尚未记住，仍用 1:5。

## 预览与写入

- 不要把 LIMIT 或 TOP 拼进用户 SQL。
- 仅预览时加会话行数上限，值为「最大预览行数 + 1」：MySQL 用 `SQL_SELECT_LIMIT`，SQL Server 用 `ROWCOUNT`。PostgreSQL 和 SQLite 没有对应会话设置，不要改写 SQL。
- 导出使用最大导出行数，不套预览用的会话行数上限。
- 达到上限后仍 `Cancel()` 以停掉服务器；释放读取器时吞掉因此产生的异常，不能把截断当成操作失败。
- 状态要写出已读行数和上限，并提示可在设置中修改「最大预览行数」。
- 预览表头两行：字段编码加粗，下一行是 CLR 类型短名、斜体、灰色。不要根据结果列去查字段注释。
- Excel 写入标题只占一行字段编码，底色为灰色 RGB(217,217,217)，编码行加粗。数据从第 2 行开始。标题和数据组成的区域使用灰色细网格线 RGB(128,128,128)。
- 连接树里表和视图的注释用灰色斜体，编码保持正文字体。表注释来自数据库目录。查不到或目录调用失败时只显示编码，不中断展开，也不弹错误窗。整批为空或目录失败时记一条文件日志：空注释用 Information，目录抛错用 Warning，不要记成 Error。Oracle 等驱动若 REMARKS 为空，由用户在 JDBC URL 中打开 `remarksReporting=true` 或对应 remarks 参数，不要在插件里拼厂商 SQL。

## 提交说明

- 使用约定式提交。类型用英文小写，摘要用中文，冒号后空一格：`feat: 用一句话说明为什么改`。
- 类型按改动性质选择：`feat` 新功能，`fix` 缺陷修复，`docs` 文档，`style` 格式，`refactor` 重构，`perf` 性能，`test` 测试，`chore` 构建或杂项。
- 标题写原因。需要补充影响或限制时，另起一行写正文。

## 设置

- 键名以 `session.` 开头的项是会话状态，不出现在设置界面。
- 新增或删除注册项时，同步修改 `DB2Sheet.Tests/Services/SettingsRegistryTests.cs` 里的注册数量断言。

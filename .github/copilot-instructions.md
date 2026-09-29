# 项目约定

本文件是仓库里唯一的 AI 约定正文。用户主目录 `~\.copilot\plans` 里的旧计划只作历史记录，与本文件冲突时以本文件和当前代码为准。

## 技术边界

- .NET Framework 4.7.2，C# 7.3。不要使用更高版本的语法。
- 旧式 csproj。新增 UI 等源文件必须在 `DB2Sheet.csproj` 里显式 `Compile Include`。
- 不要修改 `DB2Sheet1.csproj`。
- 不使用 VSTO ClickOnce。部署由 Advanced Installer 封装。
- 用户可见名称、窗口标题、界面字体和代码字体只改 `UI/AppPresentation.cs`。程序集名、命名空间、Ribbon 控件 ID、持久化目录是独立技术标识。

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
- 四个功能窗体记住宽高和最大化状态，键为 `session.formBounds`：SQL 查询、批量刷新、连接管理、设置。进度、文本输入、连接编辑这些小窗不记。

## SQL 查询窗体

- SQL 区标题为「SQL编辑器（仅限查询）」。左栏标题为「数据库连接」。
- 只有一个菜单，三项为连接、查询、编辑。
- 当前连接根节点显示 `名称 [当前]`。用 `session.activeConnectionId` 记住，打开时直接显示，不要为了显示当前连接而重新测试。
- 没有结果时预览区保持折叠。展开后编辑区与预览区高度为 1:1。
- 左右默认 1:5，即可用宽度的 1/6。左栏最小宽度 140，右栏 420。
- 只有用户拖动分割条时，才把千分比写入 `session.workspaceSplitRatio`。缩放、最大化、还原只按已存比例重算，不要把这次自动调整写回设置。值为 0 表示尚未记住，仍用 1:5。

## 预览与写入

- 不要把 LIMIT 或 TOP 拼进用户 SQL。
- 仅预览时加会话行数上限，值为「最大预览行数 + 1」：MySQL 用 `SQL_SELECT_LIMIT`，SQL Server 用 `ROWCOUNT`。PostgreSQL 和 SQLite 没有对应会话设置，不要改写 SQL。
- 导出使用最大导出行数，不套预览用的会话行数上限。
- 达到上限后仍 `Cancel()` 以停掉服务器；释放读取器时吞掉因此产生的异常，不能把截断当成操作失败。
- 状态要写出已读行数和上限，并提示可在设置中修改「最大预览行数」。
- 预览表头两行：字段名加粗，下一行是 CLR 类型短名、斜体、灰色。
- Excel 写入标题保持「字段名 + 换行 + 类型名」。不要为了和预览表头对齐去改 `Excel/ExcelResultWriter.cs`。

## 设置

- 键名以 `session.` 开头的项是会话状态，不出现在设置界面。
- 新增或删除注册项时，同步修改 `DB2Sheet.Tests/Services/SettingsRegistryTests.cs` 里的注册数量断言。

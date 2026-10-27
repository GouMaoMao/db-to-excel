# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- 修复「新建/编辑连接方案」窗体中按钮上方提示文本不自动换行的缺陷。
- 修复「JDBC 环境」窗体高度过低、长说明把「检测 Java／浏览／打开 JDK 下载页」挤出可视区域的问题。
- 全部功能窗体在小屏、高缩放下都能看全界面：窗口不超出屏幕工作区，底部按钮始终可见，放不下的内容可滚动。
- 完整计划见仓库根目录 `implementation_plan.md`。

## 已完成

- `FormSizeMemory` 新增 `ClampToWorkingArea`：首次显示、建句柄前把窗体收进光标所在屏幕的工作区，只缩小不放大；缩小后按「功能窗体相对工作区居中」的约定重新居中。
- `AppForm.SetVisibleCore` 在首次显示分支统一调用该收敛，普通显示和推迟显示两条路径都生效，所有 `AppForm` 派生窗体（含不记忆尺寸的小窗）一并受保护。
- `ConnectionProfileEditForm.ApplyHintWidths` 现在也给按钮上方的 `_status` 设最大宽度，缺 Java、测试失败等长句会折行而不是被截断。
- `JdbcEnvironmentForm` 布局改为两段：上方滚动区只放说明，下方固定区放「检测 Java」、路径行、「打开 JDK 下载页」和状态行；初始高度 520 提到 560。
- `AiModelSettingsForm` 右侧编辑表格包进自动滚动面板，根表格的状态行与操作行由固定像素高改为自适应高，操作栏改自适应高。
- `AiSkillSettingsForm` 操作行由固定 40 像素改为自适应高。
- `LocalDataForm` 引言按窗宽折行，窄窗不再横向裁掉半句。

## 本次涉及文件

- `UI/FormSizeMemory.cs`
- `UI/AppForm.cs`
- `UI/ConnectionProfileEditForm.cs`
- `UI/JdbcEnvironmentForm.cs`
- `UI/AiModelSettingsForm.cs`
- `UI/AiSkillSettingsForm.cs`
- `UI/LocalDataForm.cs`
- `implementation_plan.md`（新增，计划正文）
- `docs/HANDOFF.md`

## 验证结果

- `MSBuild DB2Sheet.csproj /t:Rebuild /p:Configuration=Debug` 通过，0 错误。
- 编译仅剩 1 个既有警告 CS2002（`UI/DatabaseTreeImageCatalog.cs` 在 csproj 中重复 Include），与本次改动无关。
- 未在 Excel 宿主中人工核对界面；以下手工项仍需使用者确认：连接方案编辑窗长提示折行、JDBC 环境窗按钮始终可见、模型设置与 Skill 设置在矮窗下的滚动与按钮完整、各窗在 150% 缩放下不超出工作区。
- 未运行 `DB2Sheet.Tests`：测试项目仍有三个既有编译错误（`AiTranscriptHtmlTests`、`AppPresentationTests`、`QueryProfileRepositoryTests`），本次未处理。

## 阻塞项与下一步

- 无阻塞。需要在 Windows 上分别以 100% 和 150% 缩放打开下列窗体做目视验收：连接管理、连接方案编辑、JDBC 环境、参数设置、模型设置、Skill 设置、本地数据、SQL 查询、批量刷新。
- 若目视发现某个窗体内容仍被裁（例如模型设置右侧字段在极矮窗口下的换行），按 `implementation_plan.md` 的同一思路补滚动或自适应行高，不要再写死更大的初始高度。
- `implementation_plan.md` 是否长期保留待定；确认完成后可删除或移入 `docs/`。

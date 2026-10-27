# 实施计划（Implementation Plan）

[概述]
本次改动目标是：① 修复「新建/编辑连接方案」窗体中，按钮上方提示文本不自动换行的 BUG；② 解决「JDBC 环境」窗体高度过低、把「检测 Java／浏览／打开 JDK 下载页」等按钮挤出可视区域的问题；③ 让其余功能窗体（连接管理、新建连接方案、参数设置、模型设置、Skill 设置、本地数据、AI 请求详情等）的高度也随分辨率自适应，使不同分辨率的用户都能看全界面。已确认的取舍是：**保持适中的窗口高度，放不下的内容用滚动条；窗口不得超出屏幕工作区，底部操作按钮始终可见**。

实现思路分两层：
- **统一兜底层**：在窗体基类 `AppForm` 首次显示、创建句柄之前，统一把窗口收进光标所在屏幕的工作区（只缩小、不放大），从而一次性覆盖所有 `AppForm` 派生的功能窗体，无论它是否使用 `FormSizeMemory` 记忆尺寸。
- **针对性修复层**：对内容会被裁掉或控件被挤出视口的窗体逐个修正——JDBC 环境窗（只让长说明滚动，操作按钮固定在下方）、模型设置窗（右侧编辑区加自动滚动、状态行/操作行改为自适应高度）、Skill 设置窗（操作行改自适应高度）、本地数据窗（说明文字按宽度换行）、连接方案编辑窗（状态提示换行）。

影响范围仅限 UI 层，不改动 JDBC／查询执行逻辑，不新增源文件，`DB2Sheet.csproj` 无需改动。

[类型]
无新增或修改的类型定义、接口、枚举、数据结构。

[文件]
修改文件（全部位于 `D:\Codes\DB2Sheet\UI`）：

- `FormSizeMemory.cs`：新增可复用的 `ClampToWorkingArea` 方法（把窗口收进工作区），已有 `Restore`／`Center` 逻辑保持不变。
- `AppForm.cs`：在 `SetVisibleCore` 首次显示分支中调用一次工作区收敛，作为全局兜底。
- `ConnectionProfileEditForm.cs`：BUG1 修复——在 `ApplyHintWidths()` 中为按钮上方的状态标签 `_status` 设置最大宽度，使其自动换行。
- `JdbcEnvironmentForm.cs`：重构 Java 分组布局——只有「说明文字」放进滚动区，`检测 Java`／路径行／`打开 JDK 下载页`／状态行放进固定在滚动区下方的面板，任何窗口高度下都可见；同时把初始高度由 520 提到 560。
- `AiModelSettingsForm.cs`：右侧编辑表格包进自动滚动面板；根表格的「状态行」与「操作行」由固定像素高改为自适应高（避免高 DPI 下按钮被裁）。
- `AiSkillSettingsForm.cs`：根表格的「操作行」由固定 40 像素改为自适应高。
- `LocalDataForm.cs`：说明文字按窗体宽度折行（原来是 `AutoSize = true` 单行，会被横向裁掉）。
- `docs/HANDOFF.md`：按 `AGENTS.md` 要求记录本次目标、改动文件、验证结果与后续事项。
- `implementation_plan.md`（仓库根目录）：本计划文档。

无需删除或移动文件，无需新增源文件。

[函数]
新增函数：
- `DB2Sheet.UI.FormSizeMemory.ClampToWorkingArea(Form form, int margin = 8)`（`UI/FormSizeMemory.cs`）
  - 用途：把窗体尺寸收到光标所在屏幕的工作区内，并且不超出。
  - 行为：`form` 为空或已释放、窗口处于最大化／最小化时直接返回；用 `CursorWorkingArea()` 得到工作区（不接触窗体、不创建句柄）；`Width`／`Height` 分别取 `Math.Min(当前值, 工作区 - margin)`；尺寸未变化时不赋值，避免多余布局。低于 `MinimumSize` 的值由 WinForms 自动抬到 `MinimumSize`。

修改函数：
- `DB2Sheet.UI.AppForm.SetVisibleCore(bool value)`（`UI/AppForm.cs`）
  - 在 `if (value && !_preparedFirstShow)` 分支内、`CreateHandle()` 之前调用 `FormSizeMemory.ClampToWorkingArea(this)`，使「推迟显示」与普通两条路径都以收敛后的尺寸创建句柄并跑布局，用户看不到尺寸跳变。
- `DB2Sheet.UI.ConnectionProfileEditForm.ApplyHintWidths()`（`UI/ConnectionProfileEditForm.cs`）
  - 增加 `_status.MaximumSize = new Size(width, 0);`（与 `_providerHint` 同一 `width`），并更新方法摘要说明「按钮上方状态提示」也会收窄换行。
- `DB2Sheet.UI.JdbcEnvironmentForm` 构造函数与相关私有方法（`UI/JdbcEnvironmentForm.cs`）
  - `javaLayout` 拆成两块：行 0（`Percent 100`）= `scrollHost`（`AutoScroll`，只装 `guide`）；行 1（`AutoSize`）= 新的 `controlsPanel`（`TableLayoutPanel`，装 `检测 Java` 行、路径行、`打开 JDK 下载页` 行、`_javaStatus`）。
  - `FitScrollContent`／`GuideSizeChanged` 调整为只按说明内容的高度决定滚动条，不再受按钮行影响。
  - 类级 `<remarks>` 更新为「说明滚动、按钮固定」的新行为。
- `DB2Sheet.UI.AiModelSettingsForm` 构造函数（`UI/AiModelSettingsForm.cs`）
  - 新增 `Panel editorScroll { Dock = Fill, AutoScroll = true }`，其内放 `editor`（`AutoSize = true, AutoSizeMode = GrowAndShrink, Dock = Top`）；`split.Panel2` 改为加入 `editorScroll`。
  - 根表格 `root.RowStyles` 中 `_status` 行（原 `Absolute(32)`）与 `actions` 行（原 `Absolute(36)`）改为 `SizeType.AutoSize`。
- `DB2Sheet.UI.AiSkillSettingsForm` 构造函数（`UI/AiSkillSettingsForm.cs`）
  - 根表格操作行由 `Absolute(40)` 改为 `SizeType.AutoSize`。
- `DB2Sheet.UI.LocalDataForm` 构造函数（`UI/LocalDataForm.cs`）
  - `intro` 改为按宽度折行：初始 `MaximumSize = new Size(Math.Max(120, ClientSize.Width - 24), 0)`，并在 `Resize` 时同步更新，避免长句横向被裁。

移除函数：无。

[类]
修改类：
- `DB2Sheet.UI.AppForm`（`UI/AppForm.cs`）：仅新增一次工作区收敛调用，公开/受保护成员不变。
- `DB2Sheet.UI.FormSizeMemory`（`UI/FormSizeMemory.cs`）：新增一个 `internal static` 方法。
- `DB2Sheet.UI.JdbcEnvironmentForm`／`AiModelSettingsForm`／`AiSkillSettingsForm`／`LocalDataForm`／`ConnectionProfileEditForm`：仅构造函数与私有布局辅助方法内的调整，不新增对外接口，不改变继承关系。

新增类／移除类：无。

[依赖]
无新增或变更依赖，不涉及 NuGet 包与项目引用。

[测试]
- 编译：`msbuild DB2Sheet.csproj`（VSTO 项目），确认无编译错误；按 `AGENTS.md` 为新增/改动的成员补中文 XML 注释。
- 手工验证（100% 与 150% 缩放，并模拟较小工作区）：
  1. 连接方案编辑窗：选择 JDBC 且缺少 Java、以及测试连接失败后，按钮上方的提示文字完整折行、不被横向裁切；`保存`／`取消` 始终可见。
  2. JDBC 环境窗：打开即可看到 `检测 Java`、路径框、`打开 JDK 下载页`、状态行；窗口调矮时只有说明区出现纵向滚动条。
  3. 连接管理／参数设置／模型设置／Skill 设置／本地数据／连接方案编辑：打开后窗口不超出屏幕，底部按钮可见；放不下的内容可滚动，不出现因标签不换行导致的横向滚动条。
  4. AI 请求详情等原本高度较大的窗体：在小屏上被收敛后内容仍可滚动查看。
  5. 进度窗、打开中提示等小窗表现不变（统一收敛不改变其尺寸）。
- 回归：确认 `FormSizeMemory` 记忆尺寸仍生效（用户拖动改尺寸后重启插件仍恢复，且不会被记忆值撑出屏幕）。
- 文档：更新 `docs/HANDOFF.md`，并检查 `docs/` 中是否有描述这些窗体尺寸、需要同步修改的段落。

[实施顺序]
1. `FormSizeMemory.ClampToWorkingArea` + `AppForm.SetVisibleCore` 接入（全局地基）。编译并抽查若干窗体。
2. `ConnectionProfileEditForm.ApplyHintWidths` 增加 `_status.MaximumSize`（BUG1）。
3. `JdbcEnvironmentForm` 构造布局重构：说明滚动、操作控件固定；初始高度提到 560。
4. `AiModelSettingsForm`：编辑区自动滚动 + 状态行/操作行自适应。
5. `AiSkillSettingsForm`：操作行自适应。
6. `LocalDataForm`：说明文字按宽度折行。
7. 全量编译 + 多分辨率/多 DPI 手工验证。
8. 更新 `docs/HANDOFF.md` 及受影响的 `docs/*`。

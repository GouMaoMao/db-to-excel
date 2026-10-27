# 设置类窗体统一改造计划

记录时间：2026-10-27。状态：已实现、编译通过、人工目视验收通过，并随 2.0.3 发布。

长期规范见仓库根目录 `AGENTS.md`，界面细则以该文件为准。本文只记本轮改造的背景、范围和步骤。

## 背景

SQL 查询和批量刷新之外，其余窗体都是设置、选择或提示性质。它们数量多、单个体量小，
搭建方式却是各写一套，导致两件事：改一次公共观感要翻五六个文件，看代码时同样的
逻辑要看三遍。

盘点结果（行数取记录时的实际文件）：

| 窗体 | 行数 | 性质 |
|---|---|---|
| `UI/QueryEditorForm.cs` | 3114 | 重，本轮不动内部逻辑 |
| `UI/SheetRefreshForm.cs` | 1599 | 重，本轮不动内部逻辑 |
| `UI/AiChatForm.cs` | 53 | 小，只换骨架 |
| `UI/OperationProgressForm.cs` | 503 | 中，只换骨架 |
| `UI/ConnectionProfileEditForm.cs` | 1120 | 中 |
| `UI/AiModelSettingsForm.cs` | 904 | 中 |
| `UI/AiSkillSettingsForm.cs` | 496 | 小 |
| `UI/JdbcEnvironmentForm.cs` | 477 | 中 |
| `UI/SettingsForm.cs` | 440 | 小 |
| `UI/ConnectionProfilesForm.cs` | 310 | 小 |
| `UI/LocalDataForm.cs` | 298 | 小 |
| `UI/RefreshSummaryForm.cs` | 220 | 小 |
| `UI/TextPromptDialog.cs` | 174 | 小 |
| `UI/ExceptionDetailForm.cs` | 162 | 小 |
| `UI/AiSentDetailForm.cs` | 156 | 小 |
| `UI/ExportSheetDialog.cs` | 151 | 已并入 `TextPromptDialog` |
| `UI/OpeningHintForm.cs` | 87 | 小 |

所有窗体都继承 `UI/AppForm.cs`，公共字体和图标已经统一。问题不在继承，在骨架和绘制。

### 重复点一：窗口骨架

`Text` + `Width/Height` + `MinimumSize` + `StartPosition` 加 `FormSizeMemory.Attach`
这五六行在 8 个窗体里一字不差（`AiChatForm`、`AiModelSettingsForm`、`AiSkillSettingsForm`、
`ConnectionProfilesForm`、`JdbcEnvironmentForm`、`QueryEditorForm`、`SettingsForm`、
`SheetRefreshForm`）。根表格
`TableLayoutPanel { Dock = Fill, Padding = 12, ColumnCount = 1 }` 连同 RowStyles
在 12 个文件里重复。底部按钮栏
`FlowLayoutPanel { AutoSize = true, FlowDirection = RightToLeft }` 在 13 个文件里重复，
而且 `Dock` 有 `Fill`、`Top`、`Bottom` 三种写法，各窗底部留白不一致就来自这里。

### 重复点二：提示文字折行有七处实现

- `AiModelSettingsForm.FitHelp` 与 `AiSkillSettingsForm.FitHelp` 逐字相同：用
  `TextRenderer.MeasureText` 量高，撑高所在行。
- `JdbcEnvironmentForm.GuideSizeChanged`：遍历标签设 `MaximumSize` 宽度。
- `LocalDataForm`：在 `Resize` 里重设 `MaximumSize`。
- `ConnectionProfileEditForm.ApplyHintWidths`：既设标签最大宽度，也设状态行。
- `SettingsForm.ApplyNoteWidths`：按输入列宽设置说明标签。
- `OperationProgressForm.ApplyWrappedLabelWidth`：按客户区宽度约束阶段标题、详情和日志。

行为并不一致：有的只裁宽度不撑高，有的撑高不裁宽度。

### 重复点三：「末项新增 + 分隔线 + 右键删除」的列表

`ModelListBox` 与 `SkillListBox` 的 `WndProc` 逐字相同，都为了在选中变化之前记住
这次是右键。`DrawEntry`、`DrawAddItem`、`AddModelItem` 与 `AddSkillItem`、
`RestorableIndex` 在两个文件里各一份，合计约 250 行纯复制。

### 重复点四：小对话框脚手架

`TextPromptDialog` 与 `ExportSheetDialog` 是同一段代码的两个副本：
`FixedDialog` + 12 像素内边距 + 提示标签 + 输入框 + 右对齐按钮栏 +
`Shown` 里全选并聚焦。

### 风格不统一

比重复更影响维护的几处：

- `ExportSheetDialog` 的标题是裸字符串，没走 `AppPresentation.WindowTitle`。
- `RefreshSummaryForm` 用 `Color.FromArgb(210, 214, 220)` 自造边框，仓库里已有
  `UI/LightBorder.cs`。
- `OpeningHintForm` 内联复刻了 `LightBorder.Wrap` 做的事。
- `LocalDataForm` 自己写了 `CenterOnWorkArea`，而 `FormSizeMemory.CursorWorkingArea`
  已经提供工作区矩形。

这些意味着改一次公共观感要翻五个文件。

## 目标

两层骨架，其余窗体只往里填内容：

- 功能窗骨架：11 个功能窗。其中 8 个记尺寸，`ConnectionProfileEditForm`、
  `LocalDataForm`、`RefreshSummaryForm` 按工作区或父窗居中；`OperationProgressForm`
  另换骨架。标题、初始尺寸、最小尺寸、居中或记忆、根表格、底部按钮栏，一次配好。
- 小对话框骨架：原有 5 个，合并后 4 个。固定对话框、无最大化、工作区或父窗居中、
  一个输入行加可选附加控件、确定与取消；`ExportSheetDialog` 并入 `TextPromptDialog`。

按 `AGENTS.md`「能并进已有逻辑就不要再写一份」，公共骨架并入已有的 `AppForm`，
不另造一层基类。需要独立文件的只有两类：自绘的带新增项列表，因为含一个 `ListBox`
子类；以及折行标签控件。

实际新增 2 个文件（`UI/AddItemList.cs`、`UI/WrapLabel.cs`），删除
`UI/ExportSheetDialog.cs`，扩写 `UI/AppForm.cs`。实现后净减约 190 行。

## 步骤

第一步（已完成），给 `AppForm` 增加 `InitializeWindow`、`CreateRootTable`、
`CreateActionBar` 三个受保护的工具方法，并把原有的 17 个 `AppForm` 派生窗体改为调用它们。
`ExportSheetDialog` 随后并入 `TextPromptDialog`，现存 16 个派生窗体。

> 记尺寸的 8 个窗体走带存储参数的重载；居中的功能窗和小对话框走显式居中重载。
> `QueryEditorForm`、`SheetRefreshForm`、`AiChatForm` 只换窗口骨架，不碰内部布局。

第二步（已完成），把 `ModelListBox`、`SkillListBox` 和两份绘制、末项、退选逻辑合成
`UI/AddItemList.cs`，模型设置与 Skill 设置改为使用它。

> `AddItemListBox` 统一处理右键状态、「新增」末项、分隔线和默认行高，并保存退选锚点；
> 右键落在「新增」上时由它恢复上一次可编辑项。非新增行通过 `EntryMeasure` 与 `EntryDraw`
> 交给两个窗体，模型设置保留子型号行高和一级缩进。右键不选中「新增」、分隔线只在上面
> 已有条目时出现的行为不变。

第三步（已完成），新增 `UI/WrapLabel.cs`，替掉七处折行实现，覆盖
`ConnectionProfileEditForm`、`JdbcEnvironmentForm`、`LocalDataForm`、
`AiModelSettingsForm`、`AiSkillSettingsForm`、`SettingsForm`、`OperationProgressForm`。

> 标签放在单列根表格时按根表格内容宽度折行，放在多列表格里时按所在列宽折行；
> 不再需要各窗自己的 `Resize` 钩子。`JdbcEnvironmentForm` 的滚动区与固定区分工不变。

第四步（已完成），删除 `ExportSheetDialog`，把「写入 Sheet」并入 `TextPromptDialog`，
只保留目标表名校验和「同时保存 SQL」这个附加勾选。

> 导出窗标题改为 `AppPresentation.WindowTitle("写入 Sheet")`。表名校验涉及
> `SqlSheetTaskReader` 的 SQL 页保留规则，移到调用方 `QueryEditorForm`，没有搬进通用对话框。
> `TextPromptDialog.TryShow` 保留普通输入重载，另有一个带附加勾选和校验的重载。

第五步（已完成），`RefreshSummaryForm`、`OpeningHintForm`、`LocalDataForm` 三处改用
`LightBorder` 与 `FormSizeMemory.CursorWorkingArea`，删掉各自的边框色常量和居中方法。

> 边框统一用 `SystemColors.ControlDark`，比原先的 RGB(210,214,220) 和 RGB(214,214,214) 深。
> `OpeningHintForm` 的无边框、不抢焦点特性保留，`ShowWithoutActivation` 和 `CreateParams`
> 没有改动。

第六步（编译已完成，人工目视待做），跑 `MSBuild DB2Sheet.csproj /t:Rebuild /p:Configuration=Debug`，
结果为 0 错误，仅剩既有的 CS2002 重复 Include 警告。全部窗体仍需人工检查边距、按钮位置、
折行和最小尺寸。

> `DB2Sheet.Tests` 仍有既有编译错误，本轮不动测试项目。新增的两个文件已写入
> `DB2Sheet.csproj` 的显式 `Compile Include`。

第七步（已完成），更新 `AGENTS.md`、`docs/HANDOFF.md` 与
`docs/ARCHITECTURE.md` 的 UI 一节，并已随 2.0.3 提交、推送和发布。

## 范围之外

- `QueryEditorForm` 与 `SheetRefreshForm` 的内部布局、分栏比例、页签行为不动，
  只把窗口骨架换成公共方法。
- `OperationProgressForm` 的两条进度条逻辑不动，只换骨架。
- 测试项目不修。
- 不引入新的第三方 UI 库，不改成 WPF，不做主题化。
- 不改 `AppPresentation` 里的用户可见文案，只让「写入 Sheet」的窗口标题改走
  `WindowTitle`。

## 风险与待办

- 折行统一后，某些窗体首次显示的高度可能比现在矮或高一两行。判断标准是底部按钮
  始终可见，不追求与改前像素一致。`WrapLabel` 已用临时窗体验证过单列和横幅两种
  布局，并已在真实 Excel 宿主里逐窗目视通过。
- 边框统一为 `SystemColors.ControlDark` 后，`RefreshSummaryForm` 和
  `OpeningHintForm` 的边框会比改前深；这是已确认的预期变化。
- `AppForm` 增加工具方法后，忘记迁移的窗体不会报错，只会在视觉上仍不统一。
  当前 16 个现存派生窗体都已调用 `InitializeWindow`；9 个功能窗和 3 个小对话框使用根表格和按钮栏。
- 共用对话框让多行输入能力对其他调用方可见。实现保留了普通输入和带附加勾选的两个
  `TryShow` 重载，业务校验留在调用方，没有把表名规则塞进通用对话框。
- 列表行高：`AddItemListBox` 的「新增」行高 36、普通行高 28、模型子型号 26，
  已用临时窗体验证，并已目视右键选中和分隔线。
- 人工目视已由用户验收通过，检查范围包括连接编辑窗的 Java 横幅、JDBC 环境窗的引导区、
  设置窗的框下说明、进度窗的阶段标题和详情、模型/Skill 列表的分隔线与缩进、写入 Sheet 的表名校验。

## 已确认的决定

1. 提交节奏：一次全改、一次提交。
2. 入口合并：删除 `ExportSheetDialog`，并入 `TextPromptDialog`。
3. 迁移范围：原有全部 17 个 `AppForm` 派生窗体（合并后现存 16 个）。
4. 边框颜色：统一用 `SystemColors.ControlDark`。
5. 折行范围：连同 `SettingsForm` 和 `OperationProgressForm` 一起替换。


# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- 设置类窗体统一改造：公共骨架、带「新增」末项的列表、折行标签、小对话框和边框/居中实现。
- 完整计划、已确认的决定和风险见 `docs/UI-UNIFICATION-PLAN.md`。
- 验收标准：16 个现存 `AppForm` 派生窗体都走公共骨架（原有 17 个，`ExportSheetDialog` 已并入）；模型设置与 Skill 设置不再各有一份列表绘制；长说明按可用宽度折行且底部按钮始终可见；编译通过。

## 已完成

- `UI/AppForm.cs` 新增 `InitializeWindow`（普通/记尺寸两个重载）、`CreateRootTable`（单列铺满）、`CreateActionBar`。
- 原有 17 个 `AppForm` 派生窗体全部改为调用 `InitializeWindow`；`ExportSheetDialog` 随后并入
  `TextPromptDialog`，现存 16 个派生窗体。9 个功能窗和 3 个小对话框使用公共根表格和按钮栏。
  `QueryEditorForm`、`SheetRefreshForm`、`AiChatForm` 只换窗口骨架，内部布局未动。
- 新增 `UI/AddItemList.cs`：统一右键状态、「新增」末项、分隔线、退选锚点和默认行高。
  模型设置与 Skill 设置改为使用 `AddItemListBox`，各自只保留非新增行的绘制/行高和编辑器逻辑。
- 新增 `UI/WrapLabel.cs`：按父表格列宽或父容器可用宽度折行并撑高。替换了模型设置、Skill 设置、
  JDBC 环境、本地数据、连接编辑、设置、进度窗共七处折行实现。
- 删除 `UI/ExportSheetDialog.cs`；`TextPromptDialog` 增加带附加勾选和校验的 `TryShow` 重载，
  普通输入继续走原重载。目标表名校验移到 `QueryEditorForm.ValidateSheetName`，标题改走
  `AppPresentation.WindowTitle("写入 Sheet")`。
- `RefreshSummaryForm`、`OpeningHintForm` 改用 `LightBorder`，边框统一为 `SystemColors.ControlDark`。
- `LocalDataForm` 改用 `FormSizeMemory.CenterOnWorkingArea`，删除自己的居中方法。
- `DB2Sheet.csproj` 增加两个新文件的 `Compile Include`，移除 `ExportSheetDialog.cs`。
- 对照当前代码检查 `AGENTS.md`，补充公共窗口骨架、`WrapLabel` 折行、带「新增」末项列表自绘、
  记尺寸窗体与小窗的居中口径，以及写入 Sheet 不记尺寸这几处约定。

## 本次涉及文件

- `UI/AppForm.cs`、`UI/FormSizeMemory.cs`
- `UI/AddItemList.cs`（新增）、`UI/WrapLabel.cs`（新增）
- `UI/AiChatForm.cs`、`UI/AiModelSettingsForm.cs`、`UI/AiSentDetailForm.cs`、`UI/AiSkillSettingsForm.cs`
- `UI/ConnectionProfileEditForm.cs`、`UI/ConnectionProfilesForm.cs`、`UI/ExceptionDetailForm.cs`
- `UI/JdbcEnvironmentForm.cs`、`UI/LocalDataForm.cs`、`UI/OpeningHintForm.cs`
- `UI/OperationProgressForm.cs`、`UI/QueryEditorForm.cs`、`UI/RefreshSummaryForm.cs`
- `UI/SettingsForm.cs`、`UI/SheetRefreshForm.cs`、`UI/TextPromptDialog.cs`
- `UI/ExportSheetDialog.cs`（删除）
- `DB2Sheet.csproj`、`Properties/AssemblyInfo.cs`
- `scripts/publish-release.ps1`
- `AGENTS.md`
- `docs/UI-UNIFICATION-PLAN.md`、`docs/HANDOFF.md`、`docs/ARCHITECTURE.md`

## 验证结果

- `MSBuild DB2Sheet.csproj /t:Rebuild /p:Configuration=Debug` 通过，0 错误。
- 仅剩既有警告 CS2002（`UI/DatabaseTreeImageCatalog.cs` 在 csproj 中重复 Include），与本次改动无关。
- 用临时 WinForms 窗体验证了 `WrapLabel` 的单列根表格和 AutoSize 横幅、`AddItemListBox` 的
  28/36 像素行高和右键退选锚点、`CreateActionBar` 的右对齐和自适应高度；`LocalDataForm`
  与 `TextPromptDialog` 缩到最小宽度后按钮仍在客户区内。
- 人工目视已由用户验收通过，包括连接编辑窗、JDBC 环境、设置、进度窗和模型/Skill 列表。
- 已对照当前代码复核 `AGENTS.md` 的界面、会话状态和文档章节，未发现其他因本次改动而失真的描述。
- 未运行 `DB2Sheet.Tests`：测试项目仍有三个既有编译错误，本轮未处理测试项目。
- 已提交 `90d20d3`（界面统一重构）、`7a1ec8c`（2.0.3 版本号）和 `a1cefb6`（发布预检空引用修复）。
- 2.0.3 已发布到 GitHub 和 Gitee：分支、标签 `v2.0.3`、Release 附件一致；
  `scripts/check-release.ps1` 返回 0。

## 阻塞项与下一步

- 无阻塞。
- 2.0.3 已发布并核验通过，工作区干净。
- 若以后调整折行或高度，优先改 `WrapLabel` 或对应窗体的根表格行样式，不要再加回各窗自己的
  `Resize`/`MaximumSize` 钩子。
- 下次发布仍先读 `.cursor/skills/publish-release/SKILL.md`；只改 `Properties/AssemblyInfo.cs`
  两处版本号，先跑 `scripts/publish-release.ps1 -CheckOnly`，再发布并核验。

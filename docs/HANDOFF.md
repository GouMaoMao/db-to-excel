# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md` 与 `docs/specs/`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- 查询窗打开时结束上次未归档对话并新建一条，关闭时归档本窗对话。输入框从对话正文下沿拖动，不设 240 上限。已归档列表可就地删除。Token 说明已改写。
- 验收状态：Excel 中拖动手柄和已归档删除已通过。

## 已完成

- 拖动手柄在对话正文下沿、引用芯片之上。向上拖输入框变高。最低约一行（48 像素），不再用 240 卡住；最多拖到正文被挤没。松开后写入 `session.aiComposerInputHeight`，注册上限 16384。
- 已归档条目左侧用 `Resources/icon_talk.png`。右键在该项旁边展开「删除」，列表保持打开。删除成功后该项从仍打开的列表里消失。左键仍恢复。页签删除使用 `Resources/icon_delete.png`。
- Token 说明改为：发给模型的文本计费单位，包括对话内容、系统提示词、勾选的 SKILL、引用的内容；不要附带 MB 级别的文件。
- 新打开 SQL 查询窗调用 `BeginQuerySession`：结束上次未归档对话再新建。只收起再展开 AI 栏不新建、不归档。关闭查询窗且未取消时 `EndQuerySession`。
- 对话正文字号改为 12 像素。`.kilo/` 已加入 `.gitignore`，不入库。

## 本次涉及文件

- `UI/AiChatControl.cs`
- `UI/AppPresentation.cs`
- `UI/QueryEditorForm.cs`
- `Services/AiTranscriptHtml.cs`
- `Services/SettingsRegistry.cs`
- `Resources/icon_delete.png`
- `Resources/icon_talk.png`
- `DB2Sheet.csproj`
- `docs/specs/ai-assistant.md`
- `docs/specs/sql-query-form.md`
- `.gitignore`
- `docs/HANDOFF.md`

## 验证结果

- 用户确认已在 Excel 中核对：从对话正文下沿向上拖，输入框能高过原先的 240，直到正文被挤没；向下拖能收回。已归档列表右键后列表仍在，旁边出现「删除」，确认删除后该条马上消失。
- 此前用 Visual Studio 18 Community 的 MSBuild 编译 `DB2Sheet.csproj`（Debug）通过，输出 `bin\Debug\DB2Sheet.dll`；只有既有的 `CS2002` 重复源文件警告（`UI/DatabaseTreeImageCatalog.cs`）。本轮未再编译。
- 未运行自动化测试：本地 `DB2Sheet.Tests` 未入库，且与当前源码不同步。
- `.kilo/` 已被忽略。`Resources/icon_session_add.png` 仍有未提交修改，此前记为与这次对话改动无关，未纳入本次提交。

## 阻塞项与下一步

- 无阻塞。
- 无功能待办。下次发布仍先读 `.cursor/skills/publish-release/SKILL.md`；只改 `Properties/AssemblyInfo.cs` 两处版本号，先跑 `scripts/publish-release.ps1 -CheckOnly`，再发布并核验。

# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md` 与 `docs/specs/`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- SQL 编辑器按宽度自动换行，不显示横向滚动条。换行只影响显示。
- 查询预览、写入、美化、复制和导出失败时，底部状态栏只留一句短说明，具体异常仍只出现在错误详情里。
- 打开提示放大为 440×148、18 磅，文案左侧为 32 像素 `icon_robot.png`。SQL 查询和批量刷新共用这一提示。
- 验收状态：已编译。尚未在 Excel 里打开查询窗核对换行、失败状态栏和打开提示。

## 已完成

- `SqlEditorControl` 打开 AvalonEdit 自动换行，横向滚动条设为 Disabled。
- `ShowError` 不再把 `exception.Message` 写入状态栏。预览失败写「查询预览失败。」或「选中文本预览失败。」，写入写「写入失败。」，保存后美化写「美化失败。」，复制写「复制失败。」，导出写「导出失败。」。错误详情和「问 AI」不变。
- 打开提示使用 `AppPresentation.OpeningHintFontSize`（18）。图标在构造提示窗时从嵌入资源解码，不在 Excel 启动时预加载；图片框在窗体关闭时释放该图。
- 连接树测试失败仍会把异常消息送到同一条状态栏，这次未改。

## 本次涉及文件

- `UI/SqlEditorControl.cs`
- `UI/QueryEditorForm.cs`
- `UI/OpeningHintForm.cs`
- `UI/AppPresentation.cs`
- `docs/specs/sql-query-form.md`
- `docs/specs/preview-and-write.md`
- `docs/specs/ui-common.md`
- `docs/HANDOFF.md`

## 验证结果

- 用 Visual Studio 18 Community 的 MSBuild 编译 `DB2Sheet.csproj`（Debug）通过，输出 `bin\Debug\DB2Sheet.dll`。只有既有的 `CS2002` 重复源文件警告（`UI/DatabaseTreeImageCatalog.cs`）。
- 未在 Excel 中打开查询窗核对。未运行自动化测试：本地 `DB2Sheet.Tests` 未入库，且与当前源码不同步。
- `Resources/icon_session_add.png` 仍有未提交修改，与这次改动无关，未纳入本次说明。

## 阻塞项与下一步

- 无阻塞。
- 请在 Excel 中打开 SQL 查询：长行应自动换行且没有横向滚动条；故意让预览失败，状态栏只留「查询预览失败。」，详情仍在错误窗里；打开过程中的「正在打开…」应更大，左侧有机器人图标。批量刷新打开时会看到同一块提示。
- 下次发布仍先读 `.cursor/skills/publish-release/SKILL.md`；只改 `Properties/AssemblyInfo.cs` 两处版本号，先跑 `scripts/publish-release.ps1 -CheckOnly`，再发布并核验。

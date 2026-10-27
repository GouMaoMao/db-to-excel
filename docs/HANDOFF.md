# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- `2.0.2` 已发布到 GitHub 和 Gitee，两边的标签、Release 和 `db-to-excel.exe` 附件一致。
- 下次发布前确认工作区干净，然后只运行 `scripts/publish-release.ps1`。

## 已完成

- `2.0.2` 已包含 `PoorMansTSqlFormatter` 1.4.3.1 的 SQL 格式化、「检测模型」列出接口返回的全部型号、AGPLv3 `LICENSE` 与 `THIRD_PARTY_NOTICES.md`，安装工程已引用这些文件。
- 本地标签 `v2.0.2` 原指向旧提交 `aa4944c`，远端没有该标签和 Release；删除该本地标签后在 `91c0beb` 由发布脚本重建。
- GitHub Release `v2.0.2` 已创建，附件 `db-to-excel.exe` 为 24,103,597 字节。
- Gitee 的 `main`、`v2.0.2` 标签、Release `2.0.2` 和同名附件已按 GitHub 对齐。
- 修复 `scripts/mirror-gitee.ps1` 的 JSON 解析：`Invoke-External` 先把原生输出按 UTF-8 解码并以换行拼接，不再用 `Out-String`，避免中文正文让 `ConvertFrom-Json` 失败。
- `.cursor/rules/db2sheet.mdc` 的约定正文入口改为 `AGENTS.md`。

## 本次涉及文件

- `.cursor/rules/db2sheet.mdc`
- `scripts/mirror-gitee.ps1`
- `docs/HANDOFF.md`

## 验证结果

- `gh release view v2.0.2` 显示附件 `db-to-excel.exe` 为 24,103,597 字节，说明首行是「## 变更」，中文正常。
- Gitee 接口 `GET /releases/tags/v2.0.2` 返回 200，附件含 `db-to-excel.exe`；`git ls-remote` 显示 Gitee 的 `main` 与 `v2.0.2` 和 GitHub 指向同一提交。
- `scripts/mirror-gitee.ps1` 本次运行退出码为 0，其余版本的 Release 判定一致并跳过。
- 修复前同一脚本在 `ConvertFrom-Json` 处失败：终端 `gb2312` 码页错误解码 `gh` 的 UTF-8 输出，中文 `body` 处截断 JSON。
- 测试项目仍因三个既有编译错误无法生成：`AiTranscriptHtmlTests` 参数类型过期、`AppPresentationTests` 缺少 Ribbon 构造参数、`QueryProfileRepositoryTests` 仍引用已删除的 `TargetSheetName`。
- 尚未在 Excel 或 WPS 宿主中人工查看编辑器格式化效果和模型检测弹窗。

## 阻塞项与下一步

- 无发布阻塞，提交本次记录后工作区应保持干净。
- 下次发布只改 `Properties/AssemblyInfo.cs` 的两处版本号，提交并推送后运行 `scripts/publish-release.ps1`。
- 待人工在 Excel 和 WPS 中验证本版的格式化输出与「检测模型」弹窗。

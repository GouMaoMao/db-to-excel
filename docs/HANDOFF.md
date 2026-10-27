# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- `2.0.2` 已发布到 GitHub 和 Gitee，两边的标签、Release 和 `db-to-excel.exe` 附件一致。
- 发布流程已有只读预检、后台运行和结果核验，失败时能定位阶段，网络类失败两次即停。
- 下次发布前先跑 `scripts/publish-release.ps1 -CheckOnly`，正式发布用 `scripts/start-release.ps1`。

## 已完成

- `2.0.2` 的 GitHub Release 与 Gitee Release 都带 `db-to-excel.exe`（24,103,597 字节），标签 `v2.0.2` 指向 `91c0beb`，两侧一致。
- 修复 `scripts/mirror-gitee.ps1` 解析 GitHub JSON 的缺陷：原生输出按 UTF-8 解码，不再用会折行的 `Out-String`。
- 发布与镜像脚本增加阶段标识和退出码：2 是网络类失败（最多两次尝试），3 是配置或状态失败。
- 镜像的标签同步改成先比较两侧提交、只把差异标签合并成一次推送；附件改用带连接和低速上限的 `curl` 下载，并按 GitHub 记录的字节数校验。
- 新增 `scripts/start-release.ps1`（后台运行，日志与状态写 `Logs\`）和 `scripts/check-release.ps1`（只读核验两侧分支、标签和附件）。
- `publish-release.ps1` 新增 `-CheckOnly`、`-ResetTag`、`-SkipMirror`；标签分新建、失败残留重建、已发布拒绝移动三种处理；已有同名同大小附件时跳过上传。
- `.cursor/skills/publish-release/SKILL.md`、`docs/DEPLOYMENT.md`、`docs/gitee-mirror.md` 写明预检、后台运行、失败分级、两次即停和只重跑镜像。
- `.gitignore` 忽略 `Logs/release-status.json`。

## 本次涉及文件

- `scripts/publish-release.ps1`
- `scripts/mirror-gitee.ps1`
- `scripts/start-release.ps1`（新增）
- `scripts/check-release.ps1`（新增）
- `.cursor/skills/publish-release/SKILL.md`
- `docs/DEPLOYMENT.md`、`docs/gitee-mirror.md`
- `.gitignore`
- `docs/HANDOFF.md`

## 验证结果

- `scripts/check-release.ps1` 退出码 0：本地、GitHub、Gitee 的 `main`、`v2.0.2` 标签对象和 24,103,597 字节附件一致。
- `scripts/start-release.ps1 -MirrorOnly` 在后台跑完，`Logs\release-status.json` 记录 `exitCode=0`；日志显示标签与 Gitee 一致、11 个 Release 判定一致并跳过，耗时约 16 秒。
- `scripts/publish-release.ps1 -CheckOnly` 在 `v2.0.2` 已发布时正确停在「检查标签」，提示改版本号，退出码 3。
- 四个 PowerShell 脚本都用 `[System.Management.Automation.Language.Parser]::ParseFile` 做过语法检查，无错误。
- `Invoke-Native` 的退出码参数改为 `Position=1`，并用最小脚本验证过位置绑定生效（`f { } 5` 返回 5）。
- 测试项目仍因三个既有编译错误无法生成：`AiTranscriptHtmlTests`、`AppPresentationTests`、`QueryProfileRepositoryTests`。本次未处理。
- 尚未在 Excel 或 WPS 宿主中人工查看编辑器格式化效果和模型检测弹窗。

## 阻塞项与下一步

- 无发布阻塞；提交本次记录后工作区应保持干净。
- 下次发布：只改 `Properties/AssemblyInfo.cs` 两处版本号，提交并推送，先跑 `-CheckOnly`，再用 `start-release.ps1` 发布，最后用 `check-release.ps1` 核验。
- 待人工在 Excel 和 WPS 中验证本版的格式化输出与「检测模型」弹窗。

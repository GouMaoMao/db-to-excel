# Agent 工作交接

本文件只记录当前工作状态，不保存长期项目规范。长期规范见仓库根目录 `AGENTS.md` 与 `docs/specs/`。新 Agent 必须用当前代码、`git status` 和相关差异核对以下内容。

## 当前目标与验收条件

- 发布 2.0.4：推送 `main`、创建 GitHub Release，并把分支、标签和 Release 镜像到 Gitee。
- 验收状态：`scripts/check-release.ps1` 已通过。标签 `v2.0.4` 停在版本提交上，不包含本交接记录。

## 已完成

- `Properties/AssemblyInfo.cs` 的 `AssemblyVersion` 与 `AssemblyFileVersion` 改为 `2.0.4`，提交 `0836b5a`（`chore: 发布 2.0.4`）。
- 安装包 `installer/Setup Files/db-to-excel.exe` 已挂到 GitHub Release `v2.0.4`，大小 24106989 字节。
- 发布前本地比 GitHub 超前的 SQL 查询窗改动和新建对话图标，已随 `main` 推上去。
- Gitee 的 `main`、标签 `v2.0.4` 和对应 Release 已按 GitHub 对齐。

## 本次涉及文件

- `Properties/AssemblyInfo.cs`
- `docs/HANDOFF.md`

## 验证结果

- 发布脚本退出码 0。日志：`Logs/release-2.0.4-20261031-064545.log`。
- `scripts/check-release.ps1` 退出码 0。当时本地、GitHub、Gitee 的 `main` 都是 `0836b5a8235ddf72865112fcfca1891415a7f960`，标签 `v2.0.4` 指向同一提交，附件字节数一致。
- SQL 查询窗的自动换行、失败状态栏和打开提示仍未在 Excel 里核对。

## 阻塞项与下一步

- 无阻塞。
- 请在 Excel 中打开 SQL 查询：长行应自动换行且没有横向滚动条；故意让预览失败，状态栏只留「查询预览失败。」，详情仍在错误窗里；打开过程中的「正在打开…」应更大，左侧有机器人图标。
- 下次发布仍先读 `.cursor/skills/publish-release/SKILL.md`；只改 `Properties/AssemblyInfo.cs` 两处版本号，先跑 `scripts/publish-release.ps1 -CheckOnly`，再发布并核验。

---
name: publish-release
description: >-
  Builds the db-to-excel setup exe from installer/db-to-excel.aip, attaches it
  to the GitHub Release for AssemblyFileVersion, then runs scripts/mirror-gitee.ps1
  to force-align the branch, version tags, and releases to Gitee. Use when the
  user asks to 发布, 打安装包, generate the installer, or upload the setup exe
  to a GitHub Release.
---

# 发布安装包

发布动作仍只由 `scripts/publish-release.ps1` 完成。不要另写发布步骤，也不要手动改安装工程的产品版本。预检、后台运行和核验用同目录的三个辅助脚本。

## 发布前

先跑只读预检：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/publish-release.ps1 -CheckOnly
```

> 预检检查工作区、分支、两处版本号、安装工程、MSBuild、`gh` 登录、Gitee 配置和标签状态，不写远端。它不检查 Advanced Installer 的实际安装和 ProductVersion 对齐，那一步在正式发布里做。

版本号只改 `Properties/AssemblyInfo.cs` 里的 `AssemblyVersion` 和 `AssemblyFileVersion`，两处必须是同一个三段号。改完先单独提交，发布期间不要夹带文档、规则或功能改动。

> 发布脚本要求工作区干净；一次发布混进两类改动会让失败后很难判断该保留还是回退。

## 运行

用后台启动脚本，不要在会随对话重启的前台终端里跑：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start-release.ps1
```

日志写到 `Logs\release-<版本>-<时间>.log`，状态写到 `Logs\release-status.json`，两者都在 `.gitignore` 里。

> 终端重启会杀死前台进程。上次发布就死在 `git push` 上，日志停在「正在推送」，看起来像断网。

读日志和状态文件判断进度，不要反复发命令试探。日志三分钟没有新增且状态里 `exitCode` 仍为空时，先确认后台 PowerShell 进程是否还在，再决定是否重跑。

## 失败怎么处理

- `exitCode=2` 是网络类失败（连接失败、超时、5xx、低速终止）。最多再重试一次；第二次仍失败就停止，把阶段、最后一条日志和结论报给用户，不要连续自动重跑。
- `exitCode=3` 是配置或状态失败（工作区脏、版本不一致、标签冲突、Gitee 配置缺失、HTTP 4xx）。改完对应内容再跑，不要原样重试。
- 失败信息形如「发布在「<阶段>」阶段失败」。阶段依次是：检查工作区、检查标签、构建 Release、生成安装包、推送分支与标签、创建 GitHub Release、镜像到 Gitee。

> 上一次把 JSON 解析失败误报成网络问题，来回试了好几轮。阶段名和退出码就是为了避免这种误判。

镜像阶段失败时 GitHub 侧已经发完，只重跑镜像：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start-release.ps1 -MirrorOnly
```

## 标签

- 本地没有、远端也没有：正常新建。
- 本地标签不指向 HEAD，但远端没有该标签也没有 GitHub Release：这是上次失败留下的，确认后加 `-ResetTag` 重建，不要手工删。
- 标签已推送或已有 Release：不要移动，改版本号后重新发布。

> 上次本地 `v2.0.2` 停在更早的提交、远端却没有它，正属于第二种状态；没有这条规则只能翻日志推断。

## 发布后核验

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/check-release.ps1
```

它只读地比较本地、GitHub、Gitee 的分支、标签对象和附件字节数，退出码 0 表示一致。不要临时拼 `gh`、`curl` 命令核验，也不要打印令牌。

## 配置与限制

本机 `GITEE_TOKEN`、`GITEE_REPO`（`owner/repo`）、`MIRROR_BRANCH` 和 `MIRROR_TAG_REGEX` 写在 `.github/gitee.local`，该文件已加入 `.gitignore`。发布脚本和镜像脚本都只读这个文件，不读环境变量。缺了立刻停止。令牌不入库、不写入 git 配置，也不打到日志里。

日常只做 Gitee 对齐时运行 `scripts/mirror-gitee.ps1`（或 `bash scripts/mirror-gitee.sh`）。它按 GitHub 对齐分支、符合规则的标签和这些标签的 Release，只推有差异的标签，附件文件名和字节数一致时不重新上传。说明见 `docs/gitee-mirror.md`。

不要提交 `installer/`、`bin/` 或 `*.pfx`。脚本不签名。

若脚本因 `bin\Release\x64\SQLite.Interop.dll` 没有进入安装工程而停止，让用户在 Advanced Installer 里把该文件加进安装目录，然后重新运行脚本。不要用脚本改文件表。

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

只运行 `scripts/publish-release.ps1`。不要另写发布步骤，也不要手动改安装工程的产品版本。

版本号只改 `Properties/AssemblyInfo.cs` 里的 `AssemblyVersion` 和 `AssemblyFileVersion`，两处必须是同一个三段号。提交到 `main` 并且工作区干净之后，在仓库根目录运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/publish-release.ps1
```

脚本会做完这些事：检查干净的 `main`、Release 构建、核对安装工程引用的文件、对齐低于程序集的产品版本并生成新的 ProductCode、用 Advanced Installer 22.9 生成 exe、打标签 `v` 加版本号、推送，再用 `gh` 把 exe 挂到对应 GitHub Release。同一版本再次打包不会更换 ProductCode。然后执行 `scripts/mirror-gitee.ps1`。

本机 `GITEE_TOKEN`、`GITEE_REPO`（`owner/repo`）、`MIRROR_BRANCH` 和 `MIRROR_TAG_REGEX` 写在 `.github/gitee.local`，该文件已加入 `.gitignore`。发布脚本和镜像脚本都只读这个文件，不读环境变量。缺了立刻停止。令牌不入库、不写入 git 配置，也不打到日志里。

日常把 GitHub 对齐到 Gitee 时，在仓库根目录执行 `bash scripts/mirror-gitee.sh`。它强制对齐分支和符合规则的标签，并同步这些标签上的 Release。附件的文件名和大小与 GitHub 一致时不重新上传。说明见 `docs/gitee-mirror.md`。不要在发布脚本里另写 Gitee 上传。

不要提交 `installer/`、`bin/` 或 `*.pfx`。脚本不签名。

若脚本因 `bin\Release\x64\SQLite.Interop.dll` 没有进入安装工程而停止，让用户在 Advanced Installer 里把该文件加进安装目录，然后重新运行脚本。不要用脚本改文件表。

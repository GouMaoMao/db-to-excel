---
name: publish-release
description: >-
  Builds the db-to-excel setup exe from installer/db-to-excel.aip, attaches it
  to the GitHub Release for AssemblyFileVersion, then uploads that same exe to
  the Gitee Release from this machine. GitHub Actions only mirrors commits and
  tags. Use when the user asks to 发布, 打安装包, generate the installer, or
  upload the setup exe to a GitHub Release.
---

# 发布安装包

只运行 `scripts/publish-release.ps1`。不要另写发布步骤，也不要手动改安装工程的产品版本。

版本号只改 `Properties/AssemblyInfo.cs` 里的 `AssemblyVersion` 和 `AssemblyFileVersion`，两处必须是同一个三段号。提交到 `main` 并且工作区干净之后，在仓库根目录运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/publish-release.ps1
```

脚本会做完这些事：检查干净的 `main`、Release 构建、核对安装工程引用的文件、对齐低于程序集的产品版本并生成新的 ProductCode、用 Advanced Installer 22.9 生成 exe、打标签 `v` 加版本号、推送，再用 `gh` 把 exe 挂到对应 GitHub Release。同一版本再次打包不会更换 ProductCode。然后用本机 exe 上传到 Gitee。

本机环境变量 `GITEE_TOKEN` 和 `GITEE_REPO`（`owner/repo`）必须已设置。脚本一开始就检查，缺了立刻停止。令牌不写入 git 配置，也不打到日志里。

`.github/workflows/mirror-gitee.yml` 只把同一次提交和标签镜像到 Gitee，不传安装包。

不要提交 `installer/`、`bin/` 或 `*.pfx`。脚本不签名。

若脚本因 `bin\Release\x64\SQLite.Interop.dll` 没有进入安装工程而停止，让用户在 Advanced Installer 里把该文件加进安装目录，然后重新运行脚本。不要用脚本改文件表。

# Gitee 镜像

本机脚本把 GitHub 上的一个分支、符合规则的标签，以及这些标签的 Release 强制对齐到 Gitee。GitHub 是主，Gitee 是从。

把 `scripts/mirror-gitee.sh`、`scripts/mirror-gitee.ps1` 和 `.github/gitee.local.example` 按原相对路径复制到目标项目。

> 在仓库根目录执行 `bash scripts/mirror-gitee.sh`。它不接受参数，也不读环境变量，只转调 PowerShell 脚本。

将 `.github/gitee.local.example` 复制为 `.github/gitee.local`，填上令牌、仓库、分支和标签正则。

> `GITEE_REPO` 形如 `owner/repo`。`MIRROR_TAG_REGEX` 是正则，示例只匹配 `v主.次.修订`。这份文件含令牌，不要提交。

确认目标项目的 `.gitignore` 包含 `.github/gitee.local`。

确认本机已安装并登录 `gh`，并且能运行 `curl.exe`。

> 草稿 Release 不复制。GitHub 自动生成的源码包不算附件。

先把要镜像的分支推到 GitHub。

在仓库根目录执行 `bash scripts/mirror-gitee.sh`。

> 脚本先读取 GitHub 的 `origin`，再强制推送分支和符合规则的标签。本机尚未推到 GitHub 的提交不会进入 Gitee。不符合规则的分支、标签和 Release 不会被删除。令牌需要能强制推送。不要在 Gitee 上直接提交。Windows 上用 Git Bash 执行。附件的文件名和字节大小与 GitHub 一致时，不下载也不上传。

标签会先比较两边指向的提交，只把有差异的标签合并在一次推送里发出。附件用带连接和低速上限的 `curl` 下载，下载后按 GitHub 记录的字节数校验，再上传到 Gitee。Gitee 接口和附件传输都最多尝试两次，第二次仍失败就停止，不再重跑整轮。

确认 Gitee 上该分支与 GitHub 指向同一次提交，并且符合规则的 Release 与 GitHub 的标题和附件一致。

需要一条命令核对时运行 `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/check-release.ps1`。它只读比较两侧的分支、标签和附件字节数，不打印令牌，退出码 0 表示一致。

> 安装包如何构建不在这套文件里。

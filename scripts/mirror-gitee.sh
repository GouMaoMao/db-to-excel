#!/usr/bin/env bash
# 转调 scripts/mirror-gitee.ps1。配置、强制对齐和 Release 同步都在那个脚本里。

set -euo pipefail

if [[ $# -ne 0 ]]; then
  echo "这个脚本不接受参数。配置写在 .github/gitee.local。"
  exit 1
fi

root=$(cd "$(dirname "$0")/.." && pwd)
cd "$root"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$root/scripts/mirror-gitee.ps1"

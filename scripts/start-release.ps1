# 在独立隐藏进程里运行发布或镜像，日志和状态写到 Logs\。
# 调用方终端关闭不会中断长任务；启动后读 Logs\release-status.json 和日志判断进度。
# 用法：
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start-release.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/start-release.ps1 -MirrorOnly

param(
    [switch]$MirrorOnly,
    [switch]$Run
)

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")

if ($Run) {
    # 子进程分支：执行目标脚本，输出进日志，结束时把退出码写进状态文件。
    # 统一按 UTF-8 解码子进程输出，避免脚本自己切换输出编码时日志出现乱码。
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
    $log = $env:DB2SHEET_RELEASE_LOG
    $status = $env:DB2SHEET_RELEASE_STATUS
    $target = $env:DB2SHEET_RELEASE_TARGET
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $target *>&1 | Out-File -LiteralPath $log -Encoding utf8
    $code = $LASTEXITCODE
    @{
        target = $target
        log = $log
        exitCode = $code
        finishedAt = (Get-Date).ToString("s")
    } | ConvertTo-Json | Out-File -LiteralPath $status -Encoding utf8
    exit $code
}

$logDir = Join-Path $Root "Logs"
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

$version = "unknown"
$assemblyInfo = Join-Path $Root "Properties\AssemblyInfo.cs"
if (Test-Path -LiteralPath $assemblyInfo) {
    $match = [regex]::Match((Get-Content -Raw -Encoding UTF8 $assemblyInfo), '\[assembly: AssemblyFileVersion\("([^"]+)"\)\]')
    if ($match.Success) { $version = $match.Groups[1].Value }
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$log = Join-Path $logDir "release-$version-$stamp.log"
$status = Join-Path $logDir "release-status.json"
$target = if ($MirrorOnly) {
    Join-Path $Root "scripts\mirror-gitee.ps1"
}
else {
    Join-Path $Root "scripts\publish-release.ps1"
}

$env:DB2SHEET_RELEASE_LOG = $log
$env:DB2SHEET_RELEASE_STATUS = $status
$env:DB2SHEET_RELEASE_TARGET = $target

@{
    target = $target
    log = $log
    exitCode = $null
    startedAt = (Get-Date).ToString("s")
    finishedAt = $null
} | ConvertTo-Json | Out-File -LiteralPath $status -Encoding utf8

Start-Process -FilePath "powershell.exe" -ArgumentList @(
    "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath, "-Run"
) -WindowStyle Hidden | Out-Null

Write-Host "已在后台启动：$target"
Write-Host "日志：$log"
Write-Host "状态：$status"
Write-Host "exitCode 为 0 表示成功，2 是网络类失败（不要自动整轮重跑），3 是配置或状态失败（先改东西再跑）。"

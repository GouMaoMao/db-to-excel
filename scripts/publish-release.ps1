# 按 AssemblyFileVersion 做 Release 构建，用 installer\db-to-excel.aip 生成 exe，并挂到 GitHub Release。
# GitHub 发布成功后执行 scripts\mirror-gitee.ps1，由它把分支、版本标签和 Release 对齐到 Gitee。
# 不修改版本号，不签名，不提交 installer、证书或 bin。
# -CheckOnly 只做只读预检；-ResetTag 重建上次失败留下、从未发布过的本地标签；-SkipMirror 只发 GitHub。

param(
    [switch]$CheckOnly,
    [switch]$ResetTag,
    [switch]$SkipMirror
)

$ErrorActionPreference = "Stop"
$script:Stage = "准备"
if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$Project = Join-Path $Root "DB2Sheet.csproj"
$AssemblyInfo = Join-Path $Root "Properties\AssemblyInfo.cs"
$InstallerProject = Join-Path $Root "installer\db-to-excel.aip"
$ReleaseDir = Join-Path $Root "bin\Release"
$SetupDir = Join-Path $Root "installer\Setup Files"
$SqliteInterop = Join-Path $ReleaseDir "x64\SQLite.Interop.dll"

function Stop-Release([string]$Message, [int]$Code = 3) {
    Write-Host "发布在「$($script:Stage)」阶段失败：$Message"
    exit $Code
}

function Set-Stage([string]$Name) {
    $script:Stage = $Name
    Write-Host "== $Name =="
}

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [scriptblock]$Command,
        [Parameter(Position = 1)]
        [int]$Code = 3
    )
    & $Command
    if ($LASTEXITCODE -ne 0) {
        Stop-Release "命令失败，退出码 $LASTEXITCODE。" $Code
    }
}

function Get-GitSubjects([string]$RevisionRange) {
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $previousEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = $utf8
    try {
        if ([string]::IsNullOrEmpty($RevisionRange)) {
            $subjects = @(git -c i18n.logOutputEncoding=utf-8 log --format="%s")
        }
        else {
            $subjects = @(git -c i18n.logOutputEncoding=utf-8 log $RevisionRange --format="%s")
        }
    }
    finally {
        [Console]::OutputEncoding = $previousEncoding
    }
    if ($LASTEXITCODE -ne 0) {
        Stop-Release "无法读取提交说明。"
    }
    return $subjects
}

function Get-ThreePartVersion([string]$Text, [string]$AttributeName) {
    $pattern = '\[assembly: ' + [regex]::Escape($AttributeName) + '\("([^"]+)"\)\]'
    $match = [regex]::Match($Text, $pattern)
    if (-not $match.Success) {
        Stop-Release "在 AssemblyInfo.cs 里找不到 $AttributeName。"
    }
    $value = $match.Groups[1].Value
    if ($value -notmatch '^\d+\.\d+\.\d+$') {
        Stop-Release "$AttributeName 必须是三段版本，当前是 $value。"
    }
    return $value
}

function Test-GitHubReleaseExists([string]$Tag) {
    $gh = Get-Command gh -ErrorAction SilentlyContinue
    if ($gh -eq $null) { return $false }
    & gh release view $Tag --json tagName 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Resolve-ReleaseTag([string]$Tag, [switch]$Repair) {
    $head = (git rev-parse HEAD).Trim()
    $remote = (git ls-remote --tags origin "refs/tags/$Tag").Trim()
    $remoteExists = -not [string]::IsNullOrWhiteSpace($remote)
    git rev-parse --verify --quiet "refs/tags/$Tag" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        if ($remoteExists) {
            Stop-Release "远端已有标签 $Tag，但本地没有。先 fetch 标签再发布。"
        }
        return "create"
    }
    $tagged = (git rev-list -n 1 $Tag).Trim()
    if ($tagged -eq $head) { return "reuse" }
    if ($remoteExists -or (Test-GitHubReleaseExists $Tag)) {
        Stop-Release "标签 $Tag 已发布并指向 $tagged，当前提交是 $head。请改版本号后发布，不要移动已发布的标签。"
    }
    if (-not $Repair) {
        Stop-Release "本地标签 $Tag 指向 $tagged，不是当前提交 $head；远端没有该标签和 Release，像是上次发布失败留下的。确认后加 -ResetTag 重建。"
    }
    Write-Host "删除上次失败留下的本地标签 $Tag，随后在当前提交重建。"
    git tag -d $Tag | Out-Null
    if ($LASTEXITCODE -ne 0) { Stop-Release "删除本地标签 $Tag 失败。" }
    return "create"
}

function Assert-GiteeMirrorConfig {
    $giteeLocal = Join-Path $Root ".github\gitee.local"
    if (-not (Test-Path -LiteralPath $giteeLocal)) {
        Stop-Release "缺少 .github\gitee.local。把 .github\gitee.local.example 复制为该文件并填上 GITEE_TOKEN、GITEE_REPO、MIRROR_BRANCH、MIRROR_TAG_REGEX。该文件已忽略，不要提交。"
    }
    $values = @{}
    foreach ($line in Get-Content -LiteralPath $giteeLocal -Encoding UTF8) {
        $trimmed = $line.Trim()
        if ($trimmed.Length -eq 0 -or $trimmed.StartsWith("#")) { continue }
        $separator = $trimmed.IndexOf("=")
        if ($separator -lt 1) { continue }
        $name = $trimmed.Substring(0, $separator).Trim()
        $value = $trimmed.Substring($separator + 1).Trim()
        $values[$name] = $value
    }
    $missing = @()
    foreach ($key in @("GITEE_TOKEN", "GITEE_REPO", "MIRROR_BRANCH", "MIRROR_TAG_REGEX")) {
        if ([string]::IsNullOrWhiteSpace($values[$key])) { $missing += $key }
    }
    if ($missing.Count -gt 0) {
        Stop-Release ".github\gitee.local 里缺少 $($missing -join '、')。对照 .github\gitee.local.example 填写。该文件已忽略，不要提交。"
    }
}

Set-Location $Root
Assert-GiteeMirrorConfig
Set-Stage "检查工作区"

$dirty = git status --porcelain
if ($LASTEXITCODE -ne 0) { Stop-Release "无法读取 git 状态。" }
if (-not [string]::IsNullOrWhiteSpace(($dirty | Out-String))) {
    Stop-Release "工作区有未提交的改动。先提交当前版本，再发布。"
}

$branch = (git rev-parse --abbrev-ref HEAD).Trim()
if ($LASTEXITCODE -ne 0) { Stop-Release "无法读取当前分支。" }
if ($branch -ne "main") {
    Stop-Release "请在 main 上发布。当前分支是 $branch。"
}

$assemblyText = Get-Content -Raw -Encoding UTF8 $AssemblyInfo
$fileVersion = Get-ThreePartVersion $assemblyText "AssemblyFileVersion"
$assemblyVersion = Get-ThreePartVersion $assemblyText "AssemblyVersion"
if ($fileVersion -ne $assemblyVersion) {
    Stop-Release "AssemblyVersion ($assemblyVersion) 与 AssemblyFileVersion ($fileVersion) 不一致。"
}
$tag = "v$fileVersion"

if (-not (Test-Path $InstallerProject)) {
    Stop-Release "找不到安装工程 $InstallerProject。"
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    Stop-Release "找不到 vswhere，无法定位 MSBuild。"
}
$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
if (-not $msbuild) {
    Stop-Release "找不到 MSBuild。"
}

Set-Stage "检查标签"
$tagAction = Resolve-ReleaseTag $tag -Repair:$ResetTag.IsPresent

if ($CheckOnly) {
    Set-Stage "预检"
    $ghTool = Get-Command gh -ErrorAction SilentlyContinue
    if ($ghTool -eq $null) { Stop-Release "找不到 GitHub CLI（gh）。安装并登录后再发布。" }
    & gh auth status | Out-Null
    if ($LASTEXITCODE -ne 0) { Stop-Release "gh 未登录。" }
    $remoteHead = ((git ls-remote origin refs/heads/main) -split "`t")[0].Trim()
    $head = (git rev-parse HEAD).Trim()
    if ($remoteHead -ne $head) {
        Write-Host "提示：GitHub main 是 $remoteHead，本地 HEAD 是 $head；正式发布会推送本地提交。"
    }
    $tagText = if ($tagAction -eq "create") { "将创建" } else { "已指向当前提交，可直接复用" }
    Write-Host "预检通过：版本 $fileVersion，标签 $tag $tagText。"
    Write-Host "预检不检查 Advanced Installer 的实际安装与 ProductVersion 对齐，那一步在正式发布里做。"
    exit 0
}

Set-Stage "构建 Release"
Write-Host "正在构建 Release $fileVersion。"
Invoke-Native { & $msbuild $Project /t:Build /p:Configuration=Release /v:minimal /nologo }

$installer = New-Object System.Xml.XmlDocument
$installer.Load($InstallerProject)
$createVersion = $installer.DocumentElement.GetAttribute("CreateVersion")
if ([string]::IsNullOrWhiteSpace($createVersion)) {
    Stop-Release "安装工程没有 CreateVersion。"
}

$missing = New-Object System.Collections.Generic.List[string]
$installerDir = Split-Path $InstallerProject
foreach ($row in $installer.SelectNodes("//ROW[@SourcePath]")) {
    $source = $row.GetAttribute("SourcePath")
    if ($source -notlike "..\bin\Release\*") { continue }
    $full = [System.IO.Path]::GetFullPath((Join-Path $installerDir $source))
    if (-not (Test-Path -LiteralPath $full)) {
        $missing.Add($full)
    }
}
if ($missing.Count -gt 0) {
    Stop-Release ("安装工程引用的文件不在 Release 输出里：`n" + ($missing -join "`n"))
}

$installerText = Get-Content -Raw -Encoding UTF8 $InstallerProject
if ($installerText -notmatch "SQLite\.Interop\.dll") {
    if (-not (Test-Path -LiteralPath $SqliteInterop)) {
        Stop-Release "Release 构建没有产出 x64\SQLite.Interop.dll。"
    }
    Stop-Release "Release 已生成 bin\Release\x64\SQLite.Interop.dll，但安装工程没有引用它。请在 Advanced Installer 里把该文件加进安装目录后再发布。脚本不会改安装工程的文件表。"
}

$advancedInstaller = Join-Path ${env:ProgramFiles(x86)} "Caphyon\Advanced Installer $createVersion\bin\x86\AdvancedInstaller.com"
if (-not (Test-Path -LiteralPath $advancedInstaller)) {
    $alternate = Join-Path $env:ProgramFiles "Caphyon\Advanced Installer $createVersion\bin\x86\AdvancedInstaller.com"
    if (Test-Path -LiteralPath $alternate) {
        $advancedInstaller = $alternate
    }
    else {
        Stop-Release "找不到 Advanced Installer $createVersion。安装工程要求这个版本。"
    }
}

$versionRow = $installer.SelectSingleNode("//COMPONENT[@cid='caphyon.advinst.msicomp.MsiPropsComponent']/ROW[@Property='ProductVersion']")
if ($versionRow -eq $null) {
    Stop-Release "安装工程里找不到 ProductVersion。"
}
$installerVersionText = $versionRow.GetAttribute("Value")
if ($installerVersionText -notmatch '^\d+\.\d+\.\d+(\.\d+)?$') {
    Stop-Release "安装工程的 ProductVersion 无法识别：$installerVersionText。"
}
$installerVersion = [version]$installerVersionText
$productVersion = [version]$fileVersion
if ($installerVersion -gt $productVersion) {
    Stop-Release "安装工程版本 $installerVersionText 高于程序集版本 $fileVersion。"
}
if ($installerVersion -lt $productVersion) {
    Write-Host "安装工程版本 $installerVersionText 低于 $fileVersion，正在对齐并生成新的 ProductCode。"
    Invoke-Native { & $advancedInstaller /edit $InstallerProject /SetVersion $fileVersion }
}

Set-Stage "生成安装包"
$buildStarted = Get-Date
Write-Host "正在生成安装包。"
Invoke-Native { & $advancedInstaller /rebuild $InstallerProject -buildslist DefaultBuild }

if (-not (Test-Path -LiteralPath $SetupDir)) {
    Stop-Release "没有找到安装包输出目录 $SetupDir。"
}
$package = Get-ChildItem -LiteralPath $SetupDir -Filter "*.exe" |
    Where-Object { $_.LastWriteTime -ge $buildStarted.AddMinutes(-1) } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($package -eq $null) {
    Stop-Release "安装包目录里没有新生成的 exe：$SetupDir。"
}

if ($tagAction -eq "create") {
    Invoke-Native { git tag -a $tag -m $fileVersion }
}

$previous = $null
foreach ($name in @(git tag -l "v*")) {
    $candidate = $name.Trim()
    if ($candidate -notmatch '^v(\d+\.\d+\.\d+)$') { continue }
    $candidateVersion = [version]$Matches[1]
    if ($candidateVersion -ge $productVersion) { continue }
    if ($previous -eq $null -or $candidateVersion -gt $previous) {
        $previous = $candidateVersion
    }
}

if ($previous -eq $null) {
    $subjects = @(Get-GitSubjects $null)
}
else {
    $subjects = @(Get-GitSubjects "v$previous..HEAD")
}

$notes = New-Object System.Text.StringBuilder
[void]$notes.AppendLine("## 变更")
[void]$notes.AppendLine()
foreach ($subject in $subjects) {
    if ([string]::IsNullOrWhiteSpace($subject)) { continue }
    [void]$notes.AppendLine("- $($subject.Trim())")
}
$notesFile = Join-Path ([System.IO.Path]::GetTempPath()) "db2sheet-release-$fileVersion.md"
$utf8 = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($notesFile, $notes.ToString().TrimEnd(), $utf8)

Set-Stage "推送分支与标签"
Write-Host "正在推送 $tag。"
$previousPrompt = $env:GIT_TERMINAL_PROMPT
$env:GIT_TERMINAL_PROMPT = "0"
try {
    # 低速 30 秒即失败，避免网络异常时一直挂着。
    Invoke-Native { git -c http.lowSpeedLimit=1000 -c http.lowSpeedTime=30 push origin HEAD } 2
    Invoke-Native { git -c http.lowSpeedLimit=1000 -c http.lowSpeedTime=30 push origin $tag } 2
}
finally {
    if ($null -eq $previousPrompt) {
        Remove-Item Env:GIT_TERMINAL_PROMPT -ErrorAction SilentlyContinue
    }
    else {
        $env:GIT_TERMINAL_PROMPT = $previousPrompt
    }
}

Set-Stage "创建 GitHub Release"

$gh = Get-Command gh -ErrorAction SilentlyContinue
if ($gh -eq $null) {
    Stop-Release "找不到 GitHub CLI（gh）。安装并登录后再发布。"
}
& gh auth status | Out-Null
if ($LASTEXITCODE -ne 0) {
    Stop-Release "gh 未登录。"
}

$previousEncoding = [Console]::OutputEncoding
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
try {
    $releaseJson = (& gh release view $tag --json tagName,assets) -join "`n"
    $releaseCode = $LASTEXITCODE
}
finally {
    [Console]::OutputEncoding = $previousEncoding
}
if ($releaseCode -eq 0) {
    Invoke-Native { gh release edit $tag --title $fileVersion --notes-file $notesFile } 2
    $asset = $null
    try {
        $release = $releaseJson | ConvertFrom-Json
        $asset = @($release.assets) | Where-Object { $_.name -eq $package.Name } | Select-Object -First 1
    }
    catch {
        # 解析失败时按没有附件处理，走重新上传；这里吞异常只为让上传兜底，失败仍会在上传步骤报错。
        $asset = $null
    }
    if ($null -ne $asset -and $asset.size -eq $package.Length) {
        Write-Host "GitHub Release 已有同名同大小的附件，跳过上传。"
    }
    else {
        Invoke-Native { gh release upload $tag $package.FullName --clobber } 2
    }
}
else {
    Invoke-Native { gh release create $tag --title $fileVersion --notes-file $notesFile --target main $package.FullName } 2
}

Write-Host "已发布 $fileVersion：$($package.FullName)"
if ($SkipMirror) {
    Write-Host "已跳过 Gitee 镜像。需要时运行 scripts\mirror-gitee.ps1。"
    exit 0
}
Set-Stage "镜像到 Gitee"
Write-Host "正在把 GitHub 镜像到 Gitee。"
$mirror = Join-Path $Root "scripts\mirror-gitee.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mirror
if ($LASTEXITCODE -ne 0) {
    $mirrorCode = $LASTEXITCODE
    Write-Host "GitHub 侧已发布完成，只有 Gitee 镜像失败。排查后只重跑 scripts\mirror-gitee.ps1，不要重跑整轮发布。"
    exit $mirrorCode
}
Write-Host "发布完成：GitHub 与 Gitee 均已就绪。"

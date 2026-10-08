# 按 AssemblyFileVersion 做 Release 构建，用 installer\db-to-excel.aip 生成 exe，并挂到 GitHub Release。
# 不修改版本号，不签名，不提交 installer、证书或 bin。

$ErrorActionPreference = "Stop"
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

function Stop-Release([string]$Message) {
    Write-Error $Message
    exit 1
}

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Command
    )
    & $Command
    if ($LASTEXITCODE -ne 0) {
        Stop-Release "命令失败，退出码 $LASTEXITCODE。"
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

Set-Location $Root

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

$tag = "v$fileVersion"
git rev-parse --verify --quiet "refs/tags/$tag" | Out-Null
if ($LASTEXITCODE -ne 0) {
    Invoke-Native { git tag -a $tag -m $fileVersion }
}
else {
    $tagged = (git rev-list -n 1 $tag).Trim()
    $head = (git rev-parse HEAD).Trim()
    if ($tagged -ne $head) {
        Stop-Release "标签 $tag 指向 $tagged，不是当前提交 $head。"
    }
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

Write-Host "正在推送 $tag。"
Invoke-Native { git push origin HEAD }
Invoke-Native { git push origin $tag }

$gh = Get-Command gh -ErrorAction SilentlyContinue
if ($gh -eq $null) {
    Stop-Release "找不到 GitHub CLI（gh）。安装并登录后再发布。"
}
& gh auth status | Out-Null
if ($LASTEXITCODE -ne 0) {
    Stop-Release "gh 未登录。"
}

& gh release view $tag --json tagName | Out-Null
if ($LASTEXITCODE -eq 0) {
    Invoke-Native { gh release edit $tag --title $fileVersion --notes-file $notesFile }
    Invoke-Native { gh release upload $tag $package.FullName --clobber }
}
else {
    Invoke-Native { gh release create $tag --title $fileVersion --notes-file $notesFile --target main $package.FullName }
}

Write-Host "已发布 $fileVersion：$($package.FullName)"

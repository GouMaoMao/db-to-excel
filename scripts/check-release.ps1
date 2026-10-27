# 只读核验本地、GitHub、Gitee 的版本、标签、分支和 Release 附件是否一致。
# 不写任何远端，不打印令牌；附件只比较文件名和字节数。
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File scripts/check-release.ps1 [-Version 2.0.2]
# 退出码：0 全部一致；1 有差异或缺失；3 配置缺失。

param(
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $Root

function Invoke-Native([string]$File, [string[]]$ArgumentList) {
    $lines = & $File @ArgumentList 2>&1
    return @{ Code = $LASTEXITCODE; Output = (@($lines) | ForEach-Object { [string]$_ }) -join "`n" }
}

function Get-RefSha($Listing, [string]$Ref) {
    foreach ($line in @($Listing.Output -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $tab = $line.IndexOf("`t")
        if ($tab -lt 0) { continue }
        if ($line.Substring($tab + 1).Trim() -eq $Ref) { return $line.Substring(0, $tab).Trim() }
    }
    return ""
}

function Get-TagCommit($Listing, [string]$Tag) {
    $sha = ""
    foreach ($line in @($Listing.Output -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $tab = $line.IndexOf("`t")
        if ($tab -lt 0) { continue }
        $ref = $line.Substring($tab + 1).Trim()
        if ($ref -eq "refs/tags/$Tag^{}") { return $line.Substring(0, $tab).Trim() }
        if ($ref -eq "refs/tags/$Tag" -and [string]::IsNullOrEmpty($sha)) { $sha = $line.Substring(0, $tab).Trim() }
    }
    return $sha
}

function Get-AssetMap($Files) {
    $map = @{}
    foreach ($file in @($Files)) {
        if ($null -eq $file) { continue }
        $name = [string]$file.name
        if ([string]::IsNullOrEmpty($name)) { $name = [string]$file.file_name }
        if ([string]::IsNullOrEmpty($name)) { continue }
        $size = $file.size
        if ($null -eq $size) { $size = $file.file_size }
        if ($null -eq $size) { $map[$name] = $null; continue }
        $map[$name] = [int64]$size
    }
    return $map
}

$problems = New-Object System.Collections.Generic.List[string]
function Add-Problem([string]$Text) {
    Write-Host "  [差异] $Text"
    $problems.Add($Text)
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    $assemblyInfo = Join-Path $Root "Properties\AssemblyInfo.cs"
    $match = [regex]::Match((Get-Content -Raw -Encoding UTF8 $assemblyInfo), '\[assembly: AssemblyFileVersion\("([^"]+)"\)\]')
    if (-not $match.Success) {
        Write-Host "无法从 Properties\AssemblyInfo.cs 读到 AssemblyFileVersion。"
        exit 3
    }
    $Version = $match.Groups[1].Value
}
$tag = "v$Version"

$giteeLocal = Join-Path $Root ".github\gitee.local"
if (-not (Test-Path -LiteralPath $giteeLocal)) {
    Write-Host "缺少 .github\gitee.local，无法核验 Gitee。"
    exit 3
}
$config = @{}
foreach ($line in Get-Content -LiteralPath $giteeLocal -Encoding UTF8) {
    $trimmed = $line.Trim()
    if ($trimmed.Length -eq 0 -or $trimmed.StartsWith("#")) { continue }
    $separator = $trimmed.IndexOf("=")
    if ($separator -lt 1) { continue }
    $config[$trimmed.Substring(0, $separator).Trim()] = $trimmed.Substring($separator + 1).Trim()
}
$repo = $config["GITEE_REPO"]
$token = $config["GITEE_TOKEN"]
$branch = $config["MIRROR_BRANCH"]
if ([string]::IsNullOrWhiteSpace($repo) -or [string]::IsNullOrWhiteSpace($token) -or [string]::IsNullOrWhiteSpace($branch)) {
    Write-Host ".github\gitee.local 缺少 GITEE_REPO、GITEE_TOKEN 或 MIRROR_BRANCH。"
    exit 3
}
$api = "https://gitee.com/api/v5/repos/$repo"
$giteeUrl = "https://gitee.com/$repo.git"

Write-Host "== 检查版本 $Version（标签 $tag）=="
Write-Host "== 本地 =="
$head = (& git rev-parse HEAD).Trim()
$currentBranch = (& git rev-parse --abbrev-ref HEAD).Trim()
$dirty = ((@(git status --porcelain) -join "`n")).Trim()
Write-Host "分支 $currentBranch，HEAD $head，工作区$(if ($dirty) { "有未提交改动" } else { "干净" })"
$localTagResult = Invoke-Native "git" @("rev-list", "-n", "1", $tag)
$localTagCommit = if ($localTagResult.Code -eq 0) { $localTagResult.Output.Trim() } else { "" }
if ([string]::IsNullOrEmpty($localTagCommit)) {
    Write-Host "本地没有标签 $tag。"
}
else {
    Write-Host "本地标签指向 $localTagCommit"
}

Write-Host "== GitHub =="
$originUrl = (& git remote get-url origin).Trim()
$slug = ""
if ($originUrl -match 'github\.com[:/]([^/]+)/(.+?)(?:\.git)?$') {
    $slug = "$($Matches[1])/$($Matches[2])"
}
$githubRefs = Invoke-Native "git" @("ls-remote", "origin", "refs/heads/$branch", "refs/tags/$tag")
if ($githubRefs.Code -ne 0) {
    if (-not [string]::IsNullOrWhiteSpace($githubRefs.Output)) { Write-Host $githubRefs.Output.TrimEnd() }
    Add-Problem "无法读取 GitHub 引用。"
}
$githubMain = Get-RefSha $githubRefs "refs/heads/$branch"
$githubTag = Get-TagCommit $githubRefs $tag
Write-Host "$branch $githubMain；标签 $githubTag"

$githubAssets = @{}
$releaseArgs = @("release", "view", $tag, "--json", "tagName,assets")
if (-not [string]::IsNullOrEmpty($slug)) { $releaseArgs += @("--repo", $slug) }
$releaseResult = Invoke-Native "gh" $releaseArgs
if ($releaseResult.Code -ne 0) {
    Add-Problem "GitHub 没有 Release $tag。"
}
else {
    try {
        $release = $releaseResult.Output | ConvertFrom-Json
        $githubAssets = Get-AssetMap $release.assets
    }
    catch {
        Add-Problem "GitHub Release $tag 的附件清单无法解析。"
    }
}
foreach ($name in @($githubAssets.Keys)) {
    Write-Host "附件 $name $($githubAssets[$name]) 字节"
}

Write-Host "== Gitee =="
$giteeRefs = Invoke-Native "git" @(
    "-c", "credential.helper=",
    "-c", "url.https://oauth2:$($token)@gitee.com/.insteadOf=https://gitee.com/",
    "ls-remote", $giteeUrl, "refs/heads/$branch", "refs/tags/$tag"
)
$giteeMain = ""
$giteeTag = ""
if ($giteeRefs.Code -ne 0) {
    Add-Problem "无法读取 Gitee 引用，可能是网络或令牌问题。"
}
else {
    $giteeMain = Get-RefSha $giteeRefs "refs/heads/$branch"
    $giteeTag = Get-TagCommit $giteeRefs $tag
    Write-Host "$branch $giteeMain；标签 $giteeTag"
}

$giteeAssets = @{}
$releaseFile = Join-Path ([System.IO.Path]::GetTempPath()) ("check-release-" + [guid]::NewGuid().ToString("N") + ".json")
$attachFile = Join-Path ([System.IO.Path]::GetTempPath()) ("check-attach-" + [guid]::NewGuid().ToString("N") + ".json")
try {
    $code = (& curl.exe -sS -G -o $releaseFile -w "%{http_code}" "$api/releases/tags/$tag" "--data-urlencode" "access_token=$token") -join ""
    if ($code -ne "200") {
        Add-Problem "Gitee 没有 Release $tag（HTTP $code）。"
    }
    else {
        $giteeRelease = [System.IO.File]::ReadAllText($releaseFile) | ConvertFrom-Json
        $releaseId = [string]$giteeRelease.id
        if ($releaseId -match '^\d+$') {
            $attachCode = (& curl.exe -sS -G -o $attachFile -w "%{http_code}" "$api/releases/$releaseId/attach_files" "--data-urlencode" "access_token=$token") -join ""
            if ($attachCode -eq "200") {
                $giteeAssets = Get-AssetMap ([System.IO.File]::ReadAllText($attachFile) | ConvertFrom-Json)
            }
            else {
                Add-Problem "无法读取 Gitee Release $tag 的附件清单（HTTP $attachCode）。"
            }
        }
    }
}
finally {
    Remove-Item -LiteralPath $releaseFile -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $attachFile -Force -ErrorAction SilentlyContinue
}

Write-Host "== 比对 =="
if ([string]::IsNullOrEmpty($githubMain) -or [string]::IsNullOrEmpty($giteeMain)) {
    Add-Problem "有一侧读不到 $branch。"
}
elseif ($githubMain -ne $giteeMain) {
    Add-Problem "GitHub 与 Gitee 的 $branch 不一致。"
}
else {
    Write-Host "两边 $branch 一致。"
}

if ([string]::IsNullOrEmpty($githubTag) -or [string]::IsNullOrEmpty($giteeTag)) {
    Add-Problem "有一侧缺少标签 $tag。"
}
elseif ($githubTag -ne $giteeTag) {
    Add-Problem "GitHub 与 Gitee 的标签 $tag 指向不同提交。"
}
else {
    Write-Host "两边标签 $tag 指向同一提交。"
}

foreach ($name in @($githubAssets.Keys)) {
    if (-not $giteeAssets.ContainsKey($name)) {
        Add-Problem "Gitee Release 缺少附件 $name。"
        continue
    }
    if ($null -ne $githubAssets[$name] -and $null -ne $giteeAssets[$name] -and $githubAssets[$name] -ne $giteeAssets[$name]) {
        Add-Problem "附件 $name 字节数不一致：GitHub $($githubAssets[$name])，Gitee $($giteeAssets[$name])。"
    }
}
if ($githubAssets.Count -gt 0) { Write-Host "已比对 $($githubAssets.Count) 个附件。" }

if ($problems.Count -eq 0) {
    Write-Host "核验通过：本地、GitHub、Gitee 一致。"
    exit 0
}
Write-Host "核验发现 $($problems.Count) 处差异。"
exit 1

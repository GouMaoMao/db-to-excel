# 按 AssemblyFileVersion 做 Release 构建，用 installer\db-to-excel.aip 生成 exe，并挂到 GitHub Release。
# GitHub 发布成功后，用本机 exe 上传到 Gitee。不从 GitHub Actions 传这个大文件。
# 不修改版本号，不签名，不提交 installer、证书或 bin。令牌不写入 git 配置，也不打到日志里。

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

function Hide-GiteeSecret([string]$Text) {
    if ([string]::IsNullOrEmpty($Text) -or [string]::IsNullOrEmpty($script:GiteeToken)) {
        return $Text
    }
    return $Text.Replace($script:GiteeToken, "***")
}

function Read-GiteeBody([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return "" }
    return [System.IO.File]::ReadAllText($Path)
}

function Show-GiteeFailure([string]$Path) {
    $raw = Read-GiteeBody $Path
    if ([string]::IsNullOrWhiteSpace($raw)) { return }
    try {
        $obj = $raw | ConvertFrom-Json
        if ($obj.message) {
            Write-Host (Hide-GiteeSecret ([string]$obj.message))
            return
        }
    }
    catch {
        # 不是 JSON 时打印原文，令牌已打码。
    }
    Write-Host (Hide-GiteeSecret $raw)
}

function Invoke-GiteeCurl {
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Url,
        [string[]]$ArgumentList
    )
    $out = Join-Path ([System.IO.Path]::GetTempPath()) ("gitee-" + [guid]::NewGuid().ToString("N") + ".json")
    $code = "000"
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $curlArgs = @(
            "-sS", "--connect-timeout", "20", "--max-time", "180",
            "-o", $out, "-w", "%{http_code}", "-X", $Method
        )
        if ($Method -eq "GET" -or $Method -eq "DELETE") {
            $curlArgs += "-G"
        }
        if ($ArgumentList) { $curlArgs += $ArgumentList }
        $curlArgs += $Url
        $err = Join-Path ([System.IO.Path]::GetTempPath()) ("gitee-" + [guid]::NewGuid().ToString("N") + ".err")
        $codeText = & curl.exe @curlArgs 2>$err
        $exit = $LASTEXITCODE
        $code = ""
        if ($null -ne $codeText) { $code = ([string]$codeText).Trim() }
        if ($code -notmatch '^[0-9]{3}$') { $code = "000" }
        $retry = $exit -ne 0 -or $code -in @("000", "408", "429", "500", "502", "503", "504")
        if (-not $retry) { return @{ Code = $code; Body = $out } }
        if ($attempt -lt 3) {
            Write-Host "Gitee 接口暂时没有响应（HTTP $code），5 秒后重试。"
            Start-Sleep -Seconds 5
        }
    }
    return @{ Code = $code; Body = $out }
}

function Invoke-GiteeGitPush([string]$Refspec) {
    $remote = "https://oauth2:$($script:GiteeToken)@gitee.com/$($script:GiteeRepo).git"
    $previousPrompt = $env:GIT_TERMINAL_PROMPT
    $env:GIT_TERMINAL_PROMPT = "0"
    try {
        $output = & git -c credential.helper= push --quiet $remote $Refspec 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            Stop-Release ("推送到 Gitee 失败。" + (Hide-GiteeSecret $output))
        }
    }
    finally {
        if ($null -eq $previousPrompt) {
            Remove-Item Env:GIT_TERMINAL_PROMPT -ErrorAction SilentlyContinue
        }
        else {
            $env:GIT_TERMINAL_PROMPT = $previousPrompt
        }
    }
}

function Get-GiteeReleaseId($Release) {
    if ($null -eq $Release) { return "" }
    if ($Release.id -match '^\d+$') { return [string]$Release.id }
    return ""
}

function Publish-GiteeRelease {
    param(
        [Parameter(Mandatory = $true)][string]$Tag,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$NotesPath,
        [Parameter(Mandatory = $true)][string]$PackagePath
    )
    $parts = $script:GiteeRepo.Split("/")
    $api = "https://gitee.com/api/v5/repos/$($parts[0])/$($parts[1])"
    $tokenArg = @("--data-urlencode", "access_token=$($script:GiteeToken)")

    Write-Host "正在把 $Tag 推到 Gitee。"
    Invoke-GiteeGitPush "HEAD:refs/heads/main"
    Invoke-GiteeGitPush "refs/tags/${Tag}:refs/tags/${Tag}"

    $found = Invoke-GiteeCurl -Method GET -Url "$api/releases/tags/$Tag" -ArgumentList $tokenArg
    $release = $null
    if ($found.Code -eq "200") {
        $raw = Read-GiteeBody $found.Body
        if (-not [string]::IsNullOrWhiteSpace($raw) -and $raw.Trim() -ne "null") {
            $release = $raw | ConvertFrom-Json
        }
    }
    elseif ($found.Code -ne "404") {
        Show-GiteeFailure $found.Body
        Stop-Release "查询 Gitee Release $Tag 失败，HTTP $($found.Code)。"
    }

    $releaseId = Get-GiteeReleaseId $release
    $meta = $tokenArg + @(
        "--data-urlencode", "tag_name=$Tag",
        "--data-urlencode", "name=$Title",
        "--data-urlencode", "body@$NotesPath",
        "--data-urlencode", "target_commitish=main",
        "--data-urlencode", "prerelease=false"
    )
    if ($releaseId) {
        $files = @()
        if ($release.assets) { $files += @($release.assets) }
        if ($release.attach_files) { $files += @($release.attach_files) }
        $removed = @{}
        foreach ($file in $files) {
            $name = [string]$file.name
            if ([string]::IsNullOrEmpty($name)) { $name = [string]$file.file_name }
            if ($name -ne "db-to-excel.exe") { continue }
            if ($file.id -notmatch '^\d+$') { continue }
            $fileId = [string]$file.id
            if ($removed.ContainsKey($fileId)) { continue }
            $removed[$fileId] = $true
            Write-Host "删除 Gitee 上的旧 db-to-excel.exe（$fileId）。"
            $deleted = Invoke-GiteeCurl -Method DELETE -Url "$api/releases/$releaseId/attach_files/$($file.id)" -ArgumentList $tokenArg
            if ($deleted.Code -ne "204" -and $deleted.Code -ne "200") {
                Show-GiteeFailure $deleted.Body
                Stop-Release "删除 Gitee 附件 $($file.id) 失败，HTTP $($deleted.Code)。"
            }
        }
        $updated = Invoke-GiteeCurl -Method PATCH -Url "$api/releases/$releaseId" -ArgumentList $meta
        if ($updated.Code -ne "200") {
            Show-GiteeFailure $updated.Body
            Stop-Release "更新 Gitee Release $Tag 失败，HTTP $($updated.Code)。"
        }
    }
    else {
        $releaseId = ""
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            $created = Invoke-GiteeCurl -Method POST -Url "$api/releases" -ArgumentList $meta
            $createdBody = Read-GiteeBody $created.Body
            $createdRelease = $null
            if (-not [string]::IsNullOrWhiteSpace($createdBody) -and $createdBody.Trim() -ne "null") {
                $createdRelease = $createdBody | ConvertFrom-Json
            }
            $releaseId = Get-GiteeReleaseId $createdRelease
            if ($releaseId) { break }
            if ($attempt -lt 3) {
                Write-Host "创建 Gitee Release $Tag 未成功（HTTP $($created.Code)），5 秒后重试。"
                Show-GiteeFailure $created.Body
                Start-Sleep -Seconds 5
                continue
            }
            Show-GiteeFailure $created.Body
            Stop-Release "创建 Gitee Release $Tag 失败，HTTP $($created.Code)。"
        }
    }

    Write-Host "正在上传 db-to-excel.exe 到 Gitee。"
    $packageForm = "file=@`"$($PackagePath.Replace('\', '/'))`";filename=db-to-excel.exe"
    $uploaded = Invoke-GiteeCurl -Method POST -Url "$api/releases/$releaseId/attach_files" -ArgumentList @(
        "--http1.1", "-H", "Expect:",
        "-F", "access_token=$($script:GiteeToken)",
        "-F", $packageForm
    )
    if ($uploaded.Code -ne "201" -and $uploaded.Code -ne "200") {
        Show-GiteeFailure $uploaded.Body
        Stop-Release "上传 db-to-excel.exe 到 Gitee Release $Tag 失败，HTTP $($uploaded.Code)。"
    }
    Write-Host "已把安装包上传到 Gitee Release $Tag。"
}

Set-Location $Root

$giteeLocal = Join-Path $Root ".github\gitee.local"
if (Test-Path -LiteralPath $giteeLocal) {
    foreach ($line in Get-Content -LiteralPath $giteeLocal -Encoding UTF8) {
        $trimmed = $line.Trim()
        if ($trimmed.Length -eq 0 -or $trimmed.StartsWith("#")) { continue }
        $separator = $trimmed.IndexOf("=")
        if ($separator -lt 1) { continue }
        $name = $trimmed.Substring(0, $separator).Trim()
        $value = $trimmed.Substring($separator + 1).Trim()
        if ($name -eq "GITEE_TOKEN" -and [string]::IsNullOrWhiteSpace($env:GITEE_TOKEN)) {
            $env:GITEE_TOKEN = $value
        }
        elseif ($name -eq "GITEE_REPO" -and [string]::IsNullOrWhiteSpace($env:GITEE_REPO)) {
            $env:GITEE_REPO = $value
        }
    }
}

$script:GiteeToken = $env:GITEE_TOKEN
$script:GiteeRepo = $env:GITEE_REPO
if ([string]::IsNullOrWhiteSpace($script:GiteeToken) -or [string]::IsNullOrWhiteSpace($script:GiteeRepo)) {
    Stop-Release "缺少 GITEE_TOKEN 或 GITEE_REPO。写在本机 .github\gitee.local，或设成环境变量。GITEE_REPO 形如 owner/repo。该文件已忽略，不要提交。"
}
if ($script:GiteeRepo -notmatch '^[^/\s]+/[^/\s]+$') {
    Stop-Release "GITEE_REPO 必须是 owner/repo，当前是：$($script:GiteeRepo)"
}
if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
    Stop-Release "找不到 curl.exe，无法把安装包上传到 Gitee。"
}

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
Publish-GiteeRelease -Tag $tag -Title $fileVersion -NotesPath $notesFile -PackagePath $package.FullName

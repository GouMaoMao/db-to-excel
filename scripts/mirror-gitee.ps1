# 把 GitHub origin 上的一个分支、符合规则的标签，以及这些标签的 Release 强制对齐到 Gitee。
# GitHub 是主，Gitee 是从。配置只读 .github/gitee.local，不读环境变量，不接受参数。
# 附件的文件名和字节大小与 GitHub 一致时不下载、不删除、不上传。
# 令牌只通过当次 git -c url.insteadOf 和 Gitee 接口使用，不写入 git 配置，远程地址本身不含令牌。

$ErrorActionPreference = "Stop"
if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
    $PSNativeCommandUseErrorActionPreference = $false
}

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$script:Stage = "读取配置"
$script:NetworkFailure = $false

function Stop-Mirror([string]$Message, [int]$Code) {
    if (-not $PSBoundParameters.ContainsKey("Code")) {
        # 传输类失败用退出码 2，其余（配置、状态、校验）用 3，方便调用方判断要不要改东西再跑。
        $Code = if ($script:NetworkFailure) { 2 } else { 3 }
    }
    Write-Host "镜像在「$($script:Stage)」阶段失败：$Message"
    exit $Code
}

function Set-Stage([string]$Name) {
    $script:Stage = $Name
    Write-Host "== $Name =="
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

function Invoke-External {
    param(
        [Parameter(Mandatory = $true)][string]$File,
        [string[]]$ArgumentList
    )
    $previous = $ErrorActionPreference
    $previousEncoding = [Console]::OutputEncoding
    $ErrorActionPreference = "Continue"
    # gh 返回 UTF-8 编码的 JSON。默认按本机 ANSI 码页解码中文会吃掉后续字节，JSON 随之失效。
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
    try {
        $lines = & $File @ArgumentList 2>&1
        $code = $LASTEXITCODE
        # 不用 Out-String：它会按终端宽度折行，把 GitHub 返回的单行 JSON 断在字符串中间，导致 ConvertFrom-Json 失败。
        $output = (@($lines) | ForEach-Object { [string]$_ }) -join "`n"
        if ($code -ne 0 -and $output -match 'Could not resolve|Failed to connect|Connection (reset|refused|timed out)|timed out|TLS|SSL|unable to access|RPC failed') {
            $script:NetworkFailure = $true
        }
        return @{ Code = $code; Output = $output }
    }
    finally {
        [Console]::OutputEncoding = $previousEncoding
        $ErrorActionPreference = $previous
    }
}

function Invoke-GiteeCurl {
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Url,
        [string[]]$ArgumentList,
        [int]$MaxTime = 180
    )
    $out = Join-Path ([System.IO.Path]::GetTempPath()) ("gitee-" + [guid]::NewGuid().ToString("N") + ".json")
    $code = "000"
    # 网络类失败最多试两次：首次加一次重试，仍失败就停止，不做无意义的重跑。
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $curlArgs = @(
            "-sS", "--connect-timeout", "20", "--max-time", "$MaxTime",
            "-o", $out, "-w", "%{http_code}", "-X", $Method
        )
        if ($Method -eq "GET" -or $Method -eq "DELETE") {
            $curlArgs += "-G"
        }
        if ($ArgumentList) { $curlArgs += $ArgumentList }
        $curlArgs += $Url
        $previous = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        try {
            $codeText = & curl.exe @curlArgs 2>$null
            $exit = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previous
        }
        $code = ""
        if ($null -ne $codeText) { $code = ([string]$codeText).Trim() }
        if ($code -notmatch '^[0-9]{3}$') { $code = "000" }
        $retry = $exit -ne 0 -or $code -in @("000", "408", "429", "500", "502", "503", "504")
        if (-not $retry) { return @{ Code = $code; Body = $out } }
        $script:NetworkFailure = $true
        if ($attempt -lt 2) {
            Write-Host "Gitee 接口暂时没有响应（HTTP $code），5 秒后重试。"
            Start-Sleep -Seconds 5
        }
    }
    return @{ Code = $code; Body = $out }
}

function Get-GiteeGitConfig {
    # 令牌只通过本次 git 调用的 insteadOf 注入，不写 git 配置，也不打印到日志。
    return @(
        "-c", "credential.helper=",
        "-c", "url.https://oauth2:$($script:GiteeToken)@gitee.com/.insteadOf=https://gitee.com/"
    )
}

function Invoke-GiteeGit([string[]]$ArgumentList) {
    $gitArgs = (Get-GiteeGitConfig) + $ArgumentList
    $previousPrompt = $env:GIT_TERMINAL_PROMPT
    $env:GIT_TERMINAL_PROMPT = "0"
    try {
        return Invoke-External -File "git" -ArgumentList $gitArgs
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

function Invoke-GiteeGitPush([string[]]$Refspec) {
    $remote = "https://gitee.com/$($script:GiteeRepo).git"
    $targets = @()
    foreach ($item in @($Refspec)) {
        $destination = $item
        $colon = $item.IndexOf(":")
        if ($colon -ge 0) { $destination = $item.Substring($colon + 1) }
        $targets += $destination
    }
    Write-Host ("正在强制推送 " + ($targets -join "、") + "。")
    $pushed = Invoke-GiteeGit (@("push", "--force", "--no-verify", $remote) + @($Refspec))
    if (-not [string]::IsNullOrWhiteSpace($pushed.Output)) {
        Write-Host (Hide-GiteeSecret $pushed.Output.TrimEnd())
    }
    if ($pushed.Code -ne 0) {
        Stop-Mirror "推送到 Gitee 失败。"
    }
}

function Get-GiteeReleaseId($Release) {
    if ($null -eq $Release) { return "" }
    if ($Release.id -match '^\d+$') { return [string]$Release.id }
    return ""
}

function Convert-JsonList([string]$Raw) {
    if ([string]::IsNullOrWhiteSpace($Raw) -or $Raw.Trim() -eq "[]" -or $Raw.Trim() -eq "null") {
        return @()
    }
    $parsed = $Raw | ConvertFrom-Json
    if ($null -eq $parsed) { return @() }
    return @($parsed)
}

function Get-FileName($File) {
    if ($null -eq $File) { return "" }
    $name = [string]$File.name
    if ([string]::IsNullOrEmpty($name)) { $name = [string]$File.file_name }
    return $name
}

function Get-FileSize($File) {
    if ($null -eq $File) { return $null }
    foreach ($prop in @("size", "file_size")) {
        $property = $File.PSObject.Properties[$prop]
        if ($null -eq $property -or $null -eq $property.Value) { continue }
        $text = [string]$property.Value
        if ($text -match '^\d+$') { return [int64]$text }
    }
    return $null
}

function Get-GiteeAttachFiles([string]$ReleaseId) {
    $found = Invoke-GiteeCurl -Method GET -Url "$($script:GiteeApi)/releases/$ReleaseId/attach_files" -ArgumentList $script:GiteeTokenArg
    if ($found.Code -ne "200") {
        Show-GiteeFailure $found.Body
        Stop-Mirror "查询 Gitee Release $ReleaseId 的附件失败，HTTP $($found.Code)。"
    }
    return Convert-JsonList (Read-GiteeBody $found.Body)
}

function Get-AssetMap($Files) {
    $map = @{}
    foreach ($file in @($Files)) {
        if ($null -eq $file) { continue }
        $name = Get-FileName $file
        if ([string]::IsNullOrEmpty($name)) { continue }
        $size = Get-FileSize $file
        if ($null -eq $size) { return $null }
        if ($map.ContainsKey($name)) { return $null }
        $map[$name] = $size
    }
    return $map
}

function Test-SameAssets($Left, $Right) {
    $leftMap = Get-AssetMap $Left
    $rightMap = Get-AssetMap $Right
    if ($null -eq $leftMap -or $null -eq $rightMap) { return $false }
    if ($leftMap.Count -ne $rightMap.Count) { return $false }
    foreach ($key in $leftMap.Keys) {
        if (-not $rightMap.ContainsKey($key) -or $rightMap[$key] -ne $leftMap[$key]) { return $false }
    }
    return $true
}

function Get-NormalizedText($Value) {
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Replace("`r`n", "`n").Replace("`r", "`n").TrimEnd()
}

function Test-Flag($Value) {
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return [bool]$Value }
    $text = ([string]$Value).Trim().ToLowerInvariant()
    return $text -eq "true" -or $text -eq "1"
}

function Read-GiteeConfig {
    $config = Join-Path $Root ".github\gitee.local"
    $example = ".github/gitee.local.example"
    if (-not (Test-Path -LiteralPath $config)) {
        Stop-Mirror "缺少 .github\gitee.local。把 $example 复制为该文件并填上 GITEE_TOKEN、GITEE_REPO、MIRROR_BRANCH、MIRROR_TAG_REGEX。该文件已忽略，不要提交。"
    }
    $values = @{}
    foreach ($line in Get-Content -LiteralPath $config -Encoding UTF8) {
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
        Stop-Mirror ".github\gitee.local 缺少：$($missing -join '、')。对照 $example 填上这四项。该文件已忽略，不要提交。"
    }
    $script:GiteeToken = $values["GITEE_TOKEN"]
    $script:GiteeRepo = $values["GITEE_REPO"]
    $script:MirrorBranch = $values["MIRROR_BRANCH"]
    $script:MirrorTagRegex = $values["MIRROR_TAG_REGEX"]
    if ($script:GiteeRepo -notmatch '^[^/\s]+/[^/\s]+$') {
        Stop-Mirror "GITEE_REPO 必须是 owner/repo，当前是：$($script:GiteeRepo)"
    }
    if ($script:MirrorBranch -match '\s|:' -or $script:MirrorBranch.StartsWith("-")) {
        Stop-Mirror "MIRROR_BRANCH 不能含空白或冒号，也不能以减号开头，当前是：$($script:MirrorBranch)"
    }
    try {
        [void][regex]::new($script:MirrorTagRegex)
    }
    catch {
        Stop-Mirror "MIRROR_TAG_REGEX 不是合法的正则：$($script:MirrorTagRegex)"
    }
}

function Get-RemoteTagMap($Listing, [string]$Source) {
    if ($Listing.Code -ne 0) {
        if (-not [string]::IsNullOrWhiteSpace($Listing.Output)) { Write-Host $Listing.Output.TrimEnd() }
        Stop-Mirror "无法读取 $Source 标签。"
    }
    # 注解标签会同时列出标签对象和解引用后的提交，用后者比较才能判断两边是不是同一个版本。
    $map = @{}
    foreach ($line in @($Listing.Output -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace([string]$line)) { continue }
        $tab = ([string]$line).IndexOf("`t")
        if ($tab -lt 0) { continue }
        $sha = ([string]$line).Substring(0, $tab).Trim()
        $ref = ([string]$line).Substring($tab + 1).Trim()
        $isPeeled = $ref.EndsWith("^{}")
        if ($isPeeled) { $ref = $ref.Substring(0, $ref.Length - 3) }
        if (-not $ref.StartsWith("refs/tags/")) { continue }
        $name = $ref.Substring("refs/tags/".Length)
        if (-not [regex]::IsMatch($name, $script:MirrorTagRegex)) { continue }
        if ($isPeeled -or -not $map.ContainsKey($name)) { $map[$name] = $sha }
    }
    return $map
}

function Sync-GitRefs {
    Set-Stage "同步分支与标签"
    Write-Host "正在从 GitHub 读取 $($script:MirrorBranch)。"
    $fetched = Invoke-External -File "git" -ArgumentList @(
        "fetch", "origin", "+refs/heads/$($script:MirrorBranch):refs/remotes/origin/$($script:MirrorBranch)"
    )
    if ($fetched.Code -ne 0) {
        if (-not [string]::IsNullOrWhiteSpace($fetched.Output)) { Write-Host $fetched.Output.TrimEnd() }
        Stop-Mirror "无法从 GitHub 读取 $($script:MirrorBranch)。"
    }
    Invoke-GiteeGitPush "refs/remotes/origin/$($script:MirrorBranch):refs/heads/$($script:MirrorBranch)"

    Write-Host "正在读取 GitHub 上的标签。"
    $githubTags = Get-RemoteTagMap -Listing (Invoke-External -File "git" -ArgumentList @("ls-remote", "--tags", "origin")) -Source "GitHub"
    $names = @($githubTags.Keys | Sort-Object)
    if ($names.Count -eq 0) {
        Write-Host "没有符合规则的标签。"
        return
    }

    # 一次取回全部符合规则的标签，再一次性比较 Gitee；逐个标签推送每次都要一次网络往返。
    $fetchArgs = @("fetch", "origin") + ($names | ForEach-Object { "+refs/tags/$($_):refs/tags/$($_)" })
    $tagFetch = Invoke-External -File "git" -ArgumentList $fetchArgs
    if ($tagFetch.Code -ne 0) {
        if (-not [string]::IsNullOrWhiteSpace($tagFetch.Output)) { Write-Host $tagFetch.Output.TrimEnd() }
        Stop-Mirror "无法从 GitHub 读取标签。"
    }

    $giteeUrl = "https://gitee.com/$($script:GiteeRepo).git"
    $giteeTags = Get-RemoteTagMap -Listing (Invoke-GiteeGit @("ls-remote", "--tags", $giteeUrl)) -Source "Gitee"
    $refspecs = @()
    foreach ($name in $names) {
        if ($giteeTags.ContainsKey($name) -and $giteeTags[$name] -eq $githubTags[$name]) { continue }
        $refspecs += "refs/tags/$($name):refs/tags/$($name)"
    }
    if ($refspecs.Count -eq 0) {
        Write-Host "标签与 Gitee 一致，跳过推送。"
        return
    }
    Invoke-GiteeGitPush $refspecs
}

function Get-GitHubRepo {
    $origin = Invoke-External -File "git" -ArgumentList @("remote", "get-url", "origin")
    if ($origin.Code -ne 0) {
        Stop-Mirror "无法读取 origin。"
    }
    $url = $origin.Output.Trim()
    if ($url -match 'github\.com[:/]([^/]+)/(.+?)(?:\.git)?$') {
        return "$($Matches[1])/$($Matches[2])"
    }
    Stop-Mirror "无法从 origin 解析 GitHub 仓库：$url"
}

function Get-GitHubReleases([string]$GitHubRepo) {
    $page = 1
    $all = @()
    while ($page -le 20) {
        $rawResult = Invoke-External -File "gh" -ArgumentList @("api", "repos/$GitHubRepo/releases?per_page=100&page=$page")
        if ($rawResult.Code -ne 0) {
            if (-not [string]::IsNullOrWhiteSpace($rawResult.Output)) { Write-Host $rawResult.Output.TrimEnd() }
            Stop-Mirror "无法读取 GitHub Release。"
        }
        $raw = $rawResult.Output
        $batch = Convert-JsonList $raw
        if ($batch.Count -eq 0) { break }
        $all += $batch
        if ($batch.Count -lt 100) { break }
        $page++
    }
    $selected = @{}
    foreach ($release in $all) {
        if ($release.draft) { continue }
        $tag = [string]$release.tag_name
        if ([string]::IsNullOrWhiteSpace($tag)) { continue }
        if (-not [regex]::IsMatch($tag, $script:MirrorTagRegex)) { continue }
        $selected[$tag] = $release
    }
    return $selected
}

function Get-GiteeReleases {
    $page = 1
    $all = @()
    while ($page -le 20) {
        $found = Invoke-GiteeCurl -Method GET -Url "$($script:GiteeApi)/releases" -ArgumentList @(
            "--data-urlencode", "access_token=$($script:GiteeToken)",
            "--data-urlencode", "page=$page",
            "--data-urlencode", "per_page=100"
        )
        if ($found.Code -ne "200") {
            Show-GiteeFailure $found.Body
            Stop-Mirror "查询 Gitee Release 列表失败，HTTP $($found.Code)。"
        }
        $batch = Convert-JsonList (Read-GiteeBody $found.Body)
        if ($batch.Count -eq 0) { break }
        $all += $batch
        if ($batch.Count -lt 100) { break }
        $page++
    }
    return $all
}

function Get-GiteeReleaseByTag([string]$Tag) {
    $found = Invoke-GiteeCurl -Method GET -Url "$($script:GiteeApi)/releases/tags/$Tag" -ArgumentList $script:GiteeTokenArg
    if ($found.Code -eq "404") { return $null }
    if ($found.Code -ne "200") {
        Show-GiteeFailure $found.Body
        Stop-Mirror "查询 Gitee Release $Tag 失败，HTTP $($found.Code)。"
    }
    $raw = Read-GiteeBody $found.Body
    if ([string]::IsNullOrWhiteSpace($raw) -or $raw.Trim() -eq "null") { return $null }
    return $raw | ConvertFrom-Json
}

function Remove-GiteeRelease([string]$ReleaseId, [string]$Tag) {
    Write-Host "正在删除 Gitee 上多出的 Release $Tag。"
    $deleted = Invoke-GiteeCurl -Method DELETE -Url "$($script:GiteeApi)/releases/$ReleaseId" -ArgumentList $script:GiteeTokenArg
    if ($deleted.Code -ne "204" -and $deleted.Code -ne "200") {
        Show-GiteeFailure $deleted.Body
        Stop-Mirror "删除 Gitee Release $Tag 失败，HTTP $($deleted.Code)。"
    }
}

function Remove-GiteeAssets([string]$ReleaseId, $Files) {
    $removed = @{}
    foreach ($file in @($Files)) {
        if ($null -eq $file) { continue }
        if ($file.id -notmatch '^\d+$') { continue }
        $fileId = [string]$file.id
        if ($removed.ContainsKey($fileId)) { continue }
        $removed[$fileId] = $true
        $deleted = Invoke-GiteeCurl -Method DELETE -Url "$($script:GiteeApi)/releases/$ReleaseId/attach_files/$fileId" -ArgumentList $script:GiteeTokenArg
        if ($deleted.Code -ne "204" -and $deleted.Code -ne "200") {
            Show-GiteeFailure $deleted.Body
            Stop-Mirror "删除 Gitee 附件 $fileId 失败，HTTP $($deleted.Code)。"
        }
    }
}

function Write-ReleaseBody([string]$Tag, [string]$Body) {
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ("gitee-body-" + [guid]::NewGuid().ToString("N") + ".md")
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($path, $(if ($null -eq $Body) { "" } else { [string]$Body }), $utf8)
    return $path
}

function Get-ReleaseMeta([string]$Tag, [string]$Title, [string]$BodyPath, [bool]$Prerelease) {
    $flag = "false"
    if ($Prerelease) { $flag = "true" }
    return $script:GiteeTokenArg + @(
        "--data-urlencode", "tag_name=$Tag",
        "--data-urlencode", "name=$Title",
        "--data-urlencode", "body@$BodyPath",
        "--data-urlencode", "target_commitish=$($script:MirrorBranch)",
        "--data-urlencode", "prerelease=$flag"
    )
}

function New-GiteeRelease([string]$Tag, [string[]]$Meta) {
    $releaseId = ""
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $created = Invoke-GiteeCurl -Method POST -Url "$($script:GiteeApi)/releases" -ArgumentList $Meta
        $createdRelease = $null
        $createdBody = Read-GiteeBody $created.Body
        if (-not [string]::IsNullOrWhiteSpace($createdBody) -and $createdBody.Trim() -ne "null") {
            $createdRelease = $createdBody | ConvertFrom-Json
        }
        $releaseId = Get-GiteeReleaseId $createdRelease
        if ($releaseId) { return $releaseId }
        if ($attempt -lt 3) {
            Write-Host "创建 Gitee Release $Tag 未成功（HTTP $($created.Code)），5 秒后重试。"
            Show-GiteeFailure $created.Body
            Start-Sleep -Seconds 5
            continue
        }
        Show-GiteeFailure $created.Body
        Stop-Mirror "创建 Gitee Release $Tag 失败，HTTP $($created.Code)。"
    }
    return ""
}

function Update-GiteeRelease([string]$ReleaseId, [string]$Tag, [string[]]$Meta) {
    $updated = Invoke-GiteeCurl -Method PATCH -Url "$($script:GiteeApi)/releases/$ReleaseId" -ArgumentList $Meta
    if ($updated.Code -ne "200") {
        Show-GiteeFailure $updated.Body
        Stop-Mirror "更新 Gitee Release $Tag 失败，HTTP $($updated.Code)。"
    }
}

function Send-GiteeAsset([string]$ReleaseId, [string]$Tag, [string]$FilePath, [string]$FileName) {
    $packageForm = "file=@`"$($FilePath.Replace('\', '/'))`";filename=$FileName"
    $uploaded = Invoke-GiteeCurl -Method POST -Url "$($script:GiteeApi)/releases/$ReleaseId/attach_files" -MaxTime 600 -ArgumentList @(
        "--http1.1", "-H", "Expect:",
        "-F", "access_token=$($script:GiteeToken)",
        "-F", $packageForm
    )
    if ($uploaded.Code -ne "201" -and $uploaded.Code -ne "200") {
        Show-GiteeFailure $uploaded.Body
        Stop-Mirror "上传 $FileName 到 Gitee Release $Tag 失败，HTTP $($uploaded.Code)。"
    }
}

function Copy-GitHubAssets([string]$Tag, [string]$ReleaseId, $GitHubFiles, $ExistingFiles) {
    $files = @($GitHubFiles)
    if ($files.Count -eq 0) { return }
    $dest = Join-Path ([System.IO.Path]::GetTempPath()) ("gitee-assets-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $dest | Out-Null
    try {
        Write-Host "正在下载 GitHub Release $Tag 的附件。"
        $paths = [ordered]@{}
        foreach ($file in $files) {
            $fileName = Get-FileName $file
            $expected = Get-FileSize $file
            $url = [string]$file.browser_download_url
            if ([string]::IsNullOrWhiteSpace($fileName) -or [string]::IsNullOrWhiteSpace($url)) {
                Stop-Mirror "GitHub Release $Tag 的附件缺少文件名或下载地址。"
            }
            $path = Join-Path $dest $fileName
            $ok = $false
            # 网络类失败最多试两次，仍失败就停止，不再整轮重跑。
            for ($attempt = 1; $attempt -le 2; $attempt++) {
                # 用 curl 而不是 gh release download：连接、总时长和低速都能设上限，慢连接不会一直挂着。
                $download = Invoke-External -File "curl.exe" -ArgumentList @(
                    "-sS", "-L", "--fail", "--connect-timeout", "15", "--max-time", "600",
                    "--speed-limit", "1024", "--speed-time", "60", "-o", $path, $url
                )
                if ($download.Code -eq 0 -and (Test-Path -LiteralPath $path)) {
                    $size = (Get-Item -LiteralPath $path).Length
                    if ($null -eq $expected -or $size -eq $expected) { $ok = $true; break }
                    Write-Host "下载的 $fileName 是 $size 字节，GitHub 记录 $expected 字节。"
                }
                elseif (-not [string]::IsNullOrWhiteSpace($download.Output)) {
                    Write-Host $download.Output.TrimEnd()
                }
                if ($attempt -lt 2) {
                    Write-Host "下载 $fileName 未成功，5 秒后重试。"
                    Start-Sleep -Seconds 5
                }
            }
            if (-not $ok) {
                $script:NetworkFailure = $true
                Stop-Mirror "下载 GitHub Release $Tag 的附件 $fileName 失败。"
            }
            $paths[$fileName] = $path
        }
        Remove-GiteeAssets -ReleaseId $ReleaseId -Files $ExistingFiles
        foreach ($name in $paths.Keys) {
            Write-Host "正在上传 $name 到 Gitee Release $Tag。"
            Send-GiteeAsset -ReleaseId $ReleaseId -Tag $Tag -FilePath $paths[$name] -FileName $name
        }
    }
    finally {
        Remove-Item -LiteralPath $dest -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Sync-OneRelease($GitHubRelease, $Existing) {
    $tag = [string]$GitHubRelease.tag_name
    $title = [string]$GitHubRelease.name
    if ([string]::IsNullOrWhiteSpace($title)) { $title = $tag }
    $body = [string]$GitHubRelease.body
    $prerelease = Test-Flag $GitHubRelease.prerelease
    $githubFiles = @()
    if ($GitHubRelease.assets) { $githubFiles = @($GitHubRelease.assets) }
    $bodyPath = Write-ReleaseBody -Tag $tag -Body $body
    try {
        $meta = Get-ReleaseMeta -Tag $tag -Title $title -BodyPath $bodyPath -Prerelease $prerelease
        $releaseId = Get-GiteeReleaseId $Existing
        if (-not $releaseId) {
            Write-Host "正在创建 Gitee Release $tag。"
            $releaseId = New-GiteeRelease -Tag $tag -Meta $meta
            Copy-GitHubAssets -Tag $tag -ReleaseId $releaseId -GitHubFiles $githubFiles -ExistingFiles @()
            return
        }
        $giteeTitle = [string]$Existing.name
        if ([string]::IsNullOrWhiteSpace($giteeTitle)) { $giteeTitle = [string]$Existing.tag_name }
        $giteePrerelease = Test-Flag $Existing.prerelease
        $sameText = (Get-NormalizedText $title) -eq (Get-NormalizedText $giteeTitle) -and
            (Get-NormalizedText $body) -eq (Get-NormalizedText $Existing.body) -and
            $prerelease -eq $giteePrerelease
        $giteeFiles = @(Get-GiteeAttachFiles $releaseId)
        $sameAssets = Test-SameAssets $githubFiles $giteeFiles
        if ($sameText -and $sameAssets) {
            Write-Host "Release $tag 与 GitHub 一致，跳过。"
            return
        }
        if (-not $sameText) {
            Write-Host "正在更新 Gitee Release $tag 的说明。"
            Update-GiteeRelease -ReleaseId $releaseId -Tag $tag -Meta $meta
        }
        if ($sameAssets) { return }
        Write-Host "正在替换 Gitee Release $tag 的附件。"
        if (@($githubFiles).Count -eq 0) {
            Remove-GiteeAssets -ReleaseId $releaseId -Files $giteeFiles
            return
        }
        Copy-GitHubAssets -Tag $tag -ReleaseId $releaseId -GitHubFiles $githubFiles -ExistingFiles $giteeFiles
    }
    finally {
        Remove-Item -LiteralPath $bodyPath -Force -ErrorAction SilentlyContinue
    }
}

function Sync-Releases {
    Set-Stage "同步 Release"
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        Stop-Mirror "找不到 GitHub CLI（gh）。安装并登录后再镜像。"
    }
    $auth = Invoke-External -File "gh" -ArgumentList @("auth", "status")
    if ($auth.Code -ne 0) {
        Stop-Mirror "gh 未登录。"
    }
    $gitHubRepo = Get-GitHubRepo
    $gitHubReleases = Get-GitHubReleases $gitHubRepo
    foreach ($release in @(Get-GiteeReleases)) {
        $tag = [string]$release.tag_name
        if ([string]::IsNullOrWhiteSpace($tag)) { continue }
        if (-not [regex]::IsMatch($tag, $script:MirrorTagRegex)) { continue }
        if ($gitHubReleases.ContainsKey($tag)) { continue }
        $releaseId = Get-GiteeReleaseId $release
        if (-not $releaseId) { continue }
        Remove-GiteeRelease -ReleaseId $releaseId -Tag $tag
    }
    foreach ($tag in @($gitHubReleases.Keys)) {
        $existing = Get-GiteeReleaseByTag $tag
        Sync-OneRelease -GitHubRelease $gitHubReleases[$tag] -Existing $existing
    }
}

if ($args.Count -gt 0) {
    Stop-Mirror "这个脚本不接受参数。配置写在 .github/gitee.local。"
}

Set-Location $Root
Read-GiteeConfig
Set-Stage "同步分支与标签"
if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
    Stop-Mirror "找不到 curl.exe，无法把 Release 同步到 Gitee。"
}
$repoParts = $script:GiteeRepo.Split("/")
$script:GiteeApi = "https://gitee.com/api/v5/repos/$($repoParts[0])/$($repoParts[1])"
$script:GiteeTokenArg = @("--data-urlencode", "access_token=$($script:GiteeToken)")

Sync-GitRefs
Sync-Releases
Write-Host "已按 GitHub 把 $($script:MirrorBranch)、符合规则的标签和 Release 对齐到 Gitee。"

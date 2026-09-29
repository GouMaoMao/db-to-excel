# 用本机 JDK 重新编译 JDBC 转接程序。插件构建不依赖此脚本；用户机器也不需要运行它。
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root "db2sheet-jdbc-bridge\src"
$classes = Join-Path $root "db2sheet-jdbc-bridge\classes"
$jar = Join-Path $root "db2sheet-jdbc-bridge.jar"

if (Test-Path $classes) {
    Remove-Item $classes -Recurse -Force
}
New-Item -ItemType Directory -Path $classes | Out-Null

$sources = @(Get-ChildItem $src -Filter *.java -Recurse | ForEach-Object { $_.FullName })
& javac --release 8 -encoding UTF-8 -d $classes @sources
if ($LASTEXITCODE -ne 0) {
    & javac -encoding UTF-8 -source 8 -target 8 -d $classes @sources
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

if (Test-Path $jar) {
    Remove-Item $jar -Force
}
& jar cfe $jar db2sheet.jdbc.Bridge -C $classes .
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Remove-Item $classes -Recurse -Force
Write-Output $jar

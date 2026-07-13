$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Python = "D:\conda_envs\edc\python.exe"

if (-not (Test-Path -LiteralPath $Python)) {
    throw "找不到 Python：$Python"
}

Set-Location -LiteralPath $ProjectRoot
& $Python -m excel_splitter.web.app


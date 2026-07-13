$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$Python = "D:\conda_envs\edc\python.exe"

Set-Location -LiteralPath $ProjectRoot
& $Python --version
& $Python -c "import flask, openpyxl; from importlib.metadata import version; print('Flask', version('flask')); print('openpyxl', openpyxl.__version__)"
& $Python -m pytest -q


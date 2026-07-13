param(
    [string]$Python = "D:\conda_envs\edc\python.exe"
)

$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$PythonRoot = Split-Path -Parent $Python
$CondaLibraryBin = Join-Path $PythonRoot "Library\bin"
$SpecFile = Join-Path $ProjectRoot "packaging\excel_splitter.spec"
$BuildRoot = Join-Path $ProjectRoot "build"
$DistRoot = Join-Path $ProjectRoot "dist"
$PackagingDir = Join-Path $ProjectRoot "packaging"
$ZipFile = Join-Path $DistRoot "ExcelSplitter-portable.zip"

if (-not (Test-Path -LiteralPath $Python -PathType Leaf)) {
    throw "Python 3.11 was not found: $Python"
}

$PythonPlatform = & $Python -c "import struct,sys;print(str(sys.version_info.major)+'.'+str(sys.version_info.minor)+'|'+str(struct.calcsize('P')*8))"
if ($PythonPlatform.Trim() -ne "3.11|64") {
    throw "Portable builds require CPython 3.11 x64. Found: $PythonPlatform"
}

if (Test-Path -LiteralPath $CondaLibraryBin -PathType Container) {
    $env:PATH = "$CondaLibraryBin;$env:PATH"
}

$ProjectRootFull = [System.IO.Path]::GetFullPath($ProjectRoot).TrimEnd('\')
foreach ($TargetRoot in @($BuildRoot, $DistRoot)) {
    $TargetRootFull = [System.IO.Path]::GetFullPath($TargetRoot)
    if (-not $TargetRootFull.StartsWith("$ProjectRootFull\", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean a path outside the project: $TargetRootFull"
    }
}
if (Test-Path -LiteralPath $BuildRoot) {
    Remove-Item -LiteralPath $BuildRoot -Recurse -Force
}
if (Test-Path -LiteralPath $DistRoot) {
    Remove-Item -LiteralPath $DistRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $BuildRoot, $DistRoot -Force | Out-Null

Set-Location -LiteralPath $PackagingDir
& $Python -m PyInstaller --noconfirm --clean `
    --workpath $BuildRoot `
    --distpath $DistRoot `
    $SpecFile

$DistDir = Get-ChildItem -LiteralPath $DistRoot -Directory |
    Where-Object {
        Get-ChildItem -LiteralPath $_.FullName -Filter "*.exe" -File
    } |
    Select-Object -First 1
if ($null -eq $DistDir) {
    throw "PyInstaller did not create a portable application directory."
}

$ReleaseTextFiles = Get-ChildItem -LiteralPath $PackagingDir -Filter "*.txt" -File
foreach ($ReleaseTextFile in $ReleaseTextFiles) {
    Copy-Item -LiteralPath $ReleaseTextFile.FullName -Destination $DistDir.FullName -Force
}

if (Test-Path -LiteralPath $ZipFile) {
    Remove-Item -LiteralPath $ZipFile -Force
}
Compress-Archive -Path $DistDir.FullName -DestinationPath $ZipFile -CompressionLevel Optimal

$ExeFile = Get-ChildItem -LiteralPath $DistDir.FullName -Filter "*.exe" -File |
    Select-Object -First 1
$Hash = Get-FileHash -LiteralPath $ZipFile -Algorithm SHA256
[pscustomobject]@{
    Exe = $ExeFile.FullName
    Zip = $ZipFile
    ZipSizeMB = [math]::Round((Get-Item -LiteralPath $ZipFile).Length / 1MB, 2)
    SHA256 = $Hash.Hash
} | Format-List

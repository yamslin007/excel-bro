$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$workDir = Join-Path $projectRoot ".tmp\installer"
$payloadDir = Join-Path $workDir "payload"
$pyInstallerDir = Join-Path $workDir "pyinstaller"
$outputDir = Join-Path $projectRoot "dist"
$installerPath = Join-Path $outputDir "Excel-Bro-Setup-0.1.0.exe"

if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $payloadDir | Out-Null
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

Push-Location $projectRoot
try {
    npm run build:addin
    python (Join-Path $PSScriptRoot "generate_certificate.py") $payloadDir
    python -m PyInstaller `
        --noconfirm `
        --clean `
        --onedir `
        --noconsole `
        --name ExcelBro `
        --distpath (Join-Path $pyInstallerDir "dist") `
        --workpath (Join-Path $pyInstallerDir "work") `
        --specpath (Join-Path $pyInstallerDir "spec") `
        --add-data "$projectRoot\server;server" `
        --add-data "$projectRoot\config;config" `
        --hidden-import tkinter `
        --hidden-import tkinter.filedialog `
        --exclude-module cv2 `
        --exclude-module IPython `
        --exclude-module matplotlib `
        --exclude-module numba `
        --exclude-module pyarrow `
        --exclude-module PyQt5 `
        --exclude-module pytest `
        --exclude-module scipy `
        --exclude-module torch `
        --exclude-module yt_dlp `
        (Join-Path $PSScriptRoot "runtime.py")
} finally {
    Pop-Location
}

Compress-Archive `
    -Path (Join-Path $pyInstallerDir "dist\ExcelBro\*") `
    -DestinationPath (Join-Path $payloadDir "runtime.zip") `
    -CompressionLevel Optimal `
    -Force
Compress-Archive `
    -Path (Join-Path $projectRoot "apps\excel-addin\dist\*") `
    -DestinationPath (Join-Path $payloadDir "frontend.zip") `
    -Force
Copy-Item (Join-Path $projectRoot "apps\excel-addin\manifest.xml") $payloadDir
Copy-Item (Join-Path $PSScriptRoot "uninstall.ps1") $payloadDir

if (Test-Path -LiteralPath $installerPath) {
    Remove-Item -LiteralPath $installerPath -Force
}
python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --noconsole `
    --name "Excel-Bro-Setup-0.1.0" `
    --distpath $outputDir `
    --workpath (Join-Path $pyInstallerDir "installer-work") `
    --specpath (Join-Path $pyInstallerDir "installer-spec") `
    --add-data "$payloadDir\runtime.zip;." `
    --add-data "$payloadDir\frontend.zip;." `
    --add-data "$payloadDir\localhost.cer;." `
    --add-data "$payloadDir\localhost.crt;." `
    --add-data "$payloadDir\localhost.key;." `
    --add-data "$payloadDir\manifest.xml;." `
    --add-data "$payloadDir\uninstall.ps1;." `
    (Join-Path $PSScriptRoot "installer.py")
if (-not (Test-Path -LiteralPath $installerPath)) {
    throw "安装包生成失败：$installerPath"
}
Get-Item -LiteralPath $installerPath

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$workDir = Join-Path $projectRoot ".tmp\installer"
$appStageDir = Join-Path $workDir "app"
$certificateStageDir = Join-Path $workDir "certificate"
$pyInstallerDir = Join-Path $workDir "pyinstaller"
$outputDir = Join-Path $projectRoot "dist"
$package = Get-Content -Raw -LiteralPath (Join-Path $projectRoot "package.json") |
    ConvertFrom-Json
$appVersion = [string]$package.version
$installerName = "Excel-Bro-Setup-$appVersion.exe"
$installerPath = Join-Path $outputDir $installerName
$manifestPath = Join-Path $projectRoot "apps\excel-addin\manifest.xml"

$manifestText = [IO.File]::ReadAllText(
    $manifestPath,
    [Text.UTF8Encoding]::new($false)
)
[xml]$manifest = $manifestText
$manifestVersion = [Version]$manifest.OfficeApp.Version
if ($manifestVersion -lt [Version]"1.0.0.0") {
    throw "Office 加载项 manifest 版本不能低于 1.0.0.0：$manifestVersion"
}
if (
    [string]::IsNullOrWhiteSpace($manifest.OfficeApp.IconUrl.DefaultValue) -or
    [string]::IsNullOrWhiteSpace(
        $manifest.OfficeApp.HighResolutionIconUrl.DefaultValue
    )
) {
    throw "Office 加载项 manifest 必须包含 IconUrl 和 HighResolutionIconUrl。"
}

function Find-InnoSetupCompiler {
    $configured = $env:INNO_SETUP_COMPILER
    if (-not [string]::IsNullOrWhiteSpace($configured)) {
        $configuredPath = [IO.Path]::GetFullPath($configured)
        if (-not (Test-Path -LiteralPath $configuredPath -PathType Leaf)) {
            throw "INNO_SETUP_COMPILER 指向的文件不存在：$configuredPath"
        }
        return $configuredPath
    }

    $command = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    $candidates = @(
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 7\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 7\ISCC.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"),
        (Join-Path $env:ProgramFiles "Inno Setup 6\ISCC.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 7\ISCC.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 6\ISCC.exe")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    throw (
        "未找到 Inno Setup 命令行编译器 ISCC.exe。请安装 Inno Setup 7，" +
        "或设置 INNO_SETUP_COMPILER 指向 ISCC.exe。"
    )
}

if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $appStageDir | Out-Null
New-Item -ItemType Directory -Force -Path $certificateStageDir | Out-Null
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

Push-Location $projectRoot
try {
    npm run build:addin
    python (Join-Path $PSScriptRoot "generate_certificate.py") `
        $certificateStageDir
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

Copy-Item `
    -Path (Join-Path $pyInstallerDir "dist\ExcelBro\*") `
    -Destination $appStageDir `
    -Recurse `
    -Force

$frontendDir = Join-Path $appStageDir "frontend"
$catalogDir = Join-Path $appStageDir "catalog"
New-Item -ItemType Directory -Force -Path $frontendDir | Out-Null
New-Item -ItemType Directory -Force -Path $catalogDir | Out-Null
Copy-Item `
    -Path (Join-Path $projectRoot "apps\excel-addin\dist\*") `
    -Destination $frontendDir `
    -Recurse `
    -Force

Copy-Item -LiteralPath $manifestPath -Destination $appStageDir
Copy-Item -LiteralPath $manifestPath -Destination $catalogDir
Copy-Item `
    -Path (Join-Path $certificateStageDir "localhost.*") `
    -Destination $appStageDir `
    -Force
Copy-Item `
    -LiteralPath (Join-Path $PSScriptRoot "install_tasks.ps1") `
    -Destination $appStageDir `
    -Force

if (Test-Path -LiteralPath $installerPath) {
    Remove-Item -LiteralPath $installerPath -Force
}

$iscc = Find-InnoSetupCompiler
$innoScript = Join-Path $PSScriptRoot "excel_bro.iss"
& $iscc `
    "/Qp" `
    "/DSourceDir=$appStageDir" `
    "/DOutputDir=$outputDir" `
    "/DAppVersion=$appVersion" `
    $innoScript
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup 构建失败，退出代码：$LASTEXITCODE"
}
if (-not (Test-Path -LiteralPath $installerPath -PathType Leaf)) {
    throw "安装包生成失败：$installerPath"
}

Get-Item -LiteralPath $installerPath

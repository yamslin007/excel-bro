param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("Install", "Uninstall")]
    [string]$Action,

    [Parameter(Mandatory = $true)]
    [string]$InstallDir
)

$ErrorActionPreference = "Stop"

$shareName = "ExcelBroAddins"
$certificateSubject = "CN=Excel Bro localhost"
$manifestName = "manifest.xml"
$executableName = "ExcelBro.exe"

function ConvertTo-PowerShellLiteral {
    param([Parameter(Mandatory = $true)][string]$Value)

    return "'" + $Value.Replace("'", "''") + "'"
}

function Resolve-ValidatedInstallDir {
    param([Parameter(Mandatory = $true)][string]$Candidate)

    $resolved = [IO.Path]::GetFullPath($Candidate)
    $root = [IO.Path]::GetPathRoot($resolved)
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "安装目录必须是绝对路径：$resolved"
    }
    if ($resolved.TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar
    ).Equals(
        $root.TrimEnd(
            [IO.Path]::DirectorySeparatorChar,
            [IO.Path]::AltDirectorySeparatorChar
        ),
        [StringComparison]::OrdinalIgnoreCase
    )) {
        throw "不能把 Excel Bro 直接安装到磁盘根目录：$resolved"
    }
    return $resolved
}

function Stop-ExcelBroRuntime {
    Get-Process -Name "ExcelBro" -ErrorAction SilentlyContinue |
        Stop-Process -Force

    for ($attempt = 0; $attempt -lt 20; $attempt += 1) {
        $remaining = Get-Process `
            -Name "ExcelBro" `
            -ErrorAction SilentlyContinue
        if ($null -eq $remaining) {
            return
        }
        Start-Sleep -Milliseconds 250
    }
    throw "Excel Bro 本地服务未能停止，请关闭后重试。"
}

function Remove-LegacyInstall {
    param([Parameter(Mandatory = $true)][string]$CurrentInstallDir)

    $legacyDir = [IO.Path]::GetFullPath(
        (Join-Path $env:LOCALAPPDATA "Programs\Excel Bro")
    )
    if ($legacyDir.Equals(
        $CurrentInstallDir,
        [StringComparison]::OrdinalIgnoreCase
    )) {
        return
    }

    $legacyExecutable = Join-Path $legacyDir $executableName
    $legacyManifest = Join-Path $legacyDir $manifestName
    $legacyUninstaller = Join-Path $legacyDir "uninstall.ps1"
    if (
        (Test-Path -LiteralPath $legacyExecutable -PathType Leaf) -and
        (Test-Path -LiteralPath $legacyManifest -PathType Leaf) -and
        (Test-Path -LiteralPath $legacyUninstaller -PathType Leaf)
    ) {
        Remove-Item -LiteralPath $legacyDir -Recurse -Force
    }
}

function Invoke-ElevatedShareAction {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Install", "Uninstall")]
        [string]$ShareAction,

        [Parameter(Mandatory = $true)]
        [string]$CatalogDir
    )

    $principal = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $shareLiteral = ConvertTo-PowerShellLiteral $shareName
    $catalogLiteral = ConvertTo-PowerShellLiteral $CatalogDir
    $principalLiteral = ConvertTo-PowerShellLiteral $principal

    if ($ShareAction -eq "Install") {
        $operation = @"
`$existing = Get-SmbShare -Name $shareLiteral -ErrorAction SilentlyContinue
if (`$null -ne `$existing) {
    Remove-SmbShare -Name $shareLiteral -Force -Confirm:`$false
}
New-SmbShare -Name $shareLiteral -Path $catalogLiteral -ReadAccess $principalLiteral |
    Out-Null
"@
    }
    else {
        $operation = @"
`$existing = Get-SmbShare -Name $shareLiteral -ErrorAction SilentlyContinue
if (`$null -ne `$existing) {
    `$actualPath = [IO.Path]::GetFullPath(`$existing.Path)
    `$expectedPath = [IO.Path]::GetFullPath($catalogLiteral)
    if (-not `$actualPath.Equals(
        `$expectedPath,
        [StringComparison]::OrdinalIgnoreCase
    )) {
        throw "共享路径与当前安装目录不一致，拒绝删除：`$actualPath"
    }
    Remove-SmbShare -Name $shareLiteral -Force -Confirm:`$false
}
"@
    }

    $elevatedScript = @"
`$ErrorActionPreference = "Stop"
$operation
"@
    $encoded = [Convert]::ToBase64String(
        [Text.Encoding]::Unicode.GetBytes($elevatedScript)
    )

    try {
        $process = Start-Process `
            -FilePath "powershell.exe" `
            -Verb RunAs `
            -WindowStyle Hidden `
            -ArgumentList "-NoProfile", "-EncodedCommand", $encoded `
            -PassThru `
            -Wait
    }
    catch {
        throw "管理员授权已取消，无法$(
            if ($ShareAction -eq "Install") { "创建" } else { "移除" }
        ) Excel 本机加载项目录。"
    }

    if ($process.ExitCode -ne 0) {
        throw "Excel 本机加载项目录操作失败，退出代码：$($process.ExitCode)。"
    }
}

function Install-RootCertificate {
    param([Parameter(Mandatory = $true)][string]$CertificatePath)

    $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $CertificatePath
    )
    $store = [Security.Cryptography.X509Certificates.X509Store]::new(
        "Root",
        [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
    )
    $store.Open(
        [Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite
    )
    try {
        $store.Add($certificate)
        foreach ($existing in @($store.Certificates)) {
            if (
                $existing.Subject -eq $certificateSubject -and
                $existing.Thumbprint -ne $certificate.Thumbprint
            ) {
                $store.Remove($existing)
            }
        }
    }
    finally {
        $store.Close()
    }
}

function Remove-RootCertificates {
    $store = [Security.Cryptography.X509Certificates.X509Store]::new(
        "Root",
        [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
    )
    $store.Open(
        [Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite
    )
    try {
        foreach ($certificate in @($store.Certificates)) {
            if ($certificate.Subject -eq $certificateSubject) {
                $store.Remove($certificate)
            }
        }
    }
    finally {
        $store.Close()
    }
}

function Start-ExcelBroRuntime {
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string]$WorkingDirectory
    )

    $runtimeProcess = Start-Process `
        -FilePath $Executable `
        -WorkingDirectory $WorkingDirectory `
        -WindowStyle Hidden `
        -PassThru

    for ($attempt = 0; $attempt -lt 40; $attempt += 1) {
        Start-Sleep -Milliseconds 500
        if ($runtimeProcess.HasExited) {
            throw "Excel Bro 本地服务启动后异常退出。"
        }
        $healthy = $false
        try {
            $response = Invoke-WebRequest `
                -UseBasicParsing `
                -Uri "http://127.0.0.1:8765/health" `
                -TimeoutSec 1
            if ($response.StatusCode -eq 200) {
                $healthy = $true
            }
        }
        catch {
            continue
        }
        if ($healthy) {
            $listener = Get-NetTCPConnection `
                -LocalPort 8765 `
                -State Listen `
                -ErrorAction SilentlyContinue
            if ($null -ne $listener -and
                $listener.OwningProcess -eq $runtimeProcess.Id) {
                return
            }
            throw "健康检查对应的进程不是当前安装目录中的 Excel Bro。"
        }
    }
    throw "Excel Bro 本地服务未能启动。"
}

$resolvedInstallDir = Resolve-ValidatedInstallDir $InstallDir
$catalogDir = Join-Path $resolvedInstallDir "catalog"
$manifestPath = Join-Path $catalogDir $manifestName
$executablePath = Join-Path $resolvedInstallDir $executableName
$certificatePath = Join-Path $resolvedInstallDir "localhost.cer"

if ($Action -eq "Install") {
    foreach ($requiredPath in (
        $manifestPath,
        $executablePath,
        $certificatePath
    )) {
        if (-not (Test-Path -LiteralPath $requiredPath)) {
            throw "安装文件不完整：$requiredPath"
        }
    }

    $shareCreated = $false
    try {
        Stop-ExcelBroRuntime
        Invoke-ElevatedShareAction -ShareAction Install -CatalogDir $catalogDir
        $shareCreated = $true
        Install-RootCertificate $certificatePath
        Remove-LegacyInstall $resolvedInstallDir
        Start-ExcelBroRuntime `
            -Executable $executablePath `
            -WorkingDirectory $resolvedInstallDir
    }
    catch {
        Stop-ExcelBroRuntime -ErrorAction SilentlyContinue
        Remove-RootCertificates
        if ($shareCreated) {
            try {
                Invoke-ElevatedShareAction `
                    -ShareAction Uninstall `
                    -CatalogDir $catalogDir
            }
            catch {
                # Preserve the original installation error.
            }
        }
        throw
    }
}
else {
    try {
        Stop-ExcelBroRuntime
        Invoke-ElevatedShareAction `
            -ShareAction Uninstall `
            -CatalogDir $catalogDir
        Remove-RootCertificates
        Remove-LegacyInstall $resolvedInstallDir
    }
    catch {
        if (Test-Path -LiteralPath $executablePath) {
            Start-Process `
                -FilePath $executablePath `
                -WorkingDirectory $resolvedInstallDir `
                -WindowStyle Hidden `
                -ErrorAction SilentlyContinue
        }
        throw
    }
}

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
Get-NetFirewallRule -Direction Inbound |
    Where-Object { `$_.DisplayName -match 'SMB-In' } |
    Enable-NetFirewallRule -ErrorAction SilentlyContinue
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

$certificateForegroundHelper = @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class ExcelBroForeground {
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint a, uint b, bool f);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    private delegate bool EnumWindowsProc(IntPtr h, IntPtr p);
    private const int SW_SHOW = 5;
    public static bool BringDialogToFront(uint targetPid) {
        IntPtr target = IntPtr.Zero;
        EnumWindows(delegate(IntPtr h, IntPtr p) {
            uint pid;
            GetWindowThreadProcessId(h, out pid);
            if (pid != targetPid) return true;
            if (!IsWindowVisible(h)) return true;
            StringBuilder sb = new StringBuilder(64);
            GetClassName(h, sb, sb.Capacity);
            if (sb.ToString() == "#32770") { target = h; return false; }
            return true;
        }, IntPtr.Zero);
        if (target == IntPtr.Zero) return false;
        uint dummy;
        uint fgThread = GetWindowThreadProcessId(GetForegroundWindow(), out dummy);
        uint thisThread = GetCurrentThreadId();
        AttachThreadInput(thisThread, fgThread, true);
        ShowWindow(target, SW_SHOW);
        BringWindowToTop(target);
        SetForegroundWindow(target);
        AttachThreadInput(thisThread, fgThread, false);
        return true;
    }
}
'@

function Install-RootCertificate {
    param([Parameter(Mandatory = $true)][string]$CertificatePath)

    $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $CertificatePath
    )

    if (-not ("ExcelBroForeground" -as [type])) {
        Add-Type -TypeDefinition $certificateForegroundHelper
    }

    # $store.Add() 会弹出系统的证书信任"安全警告"对话框。因为本进程
    # 由安装器以隐藏窗口方式启动，该对话框默认会被压在安装向导后面。
    # 把 Add 放到后台 runspace 执行，主线程轮询把对话框强制拉到前台。
    $powerShell = [PowerShell]::Create()
    $powerShell.AddScript({
        param($StorePath, $Subject, $Thumbprint)
        $cert = [Security.Cryptography.X509Certificates.X509Certificate2]::new(
            $StorePath
        )
        $rootStore = [Security.Cryptography.X509Certificates.X509Store]::new(
            "Root",
            [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
        )
        $rootStore.Open(
            [Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite
        )
        try {
            $rootStore.Add($cert)
            foreach ($existing in @($rootStore.Certificates)) {
                if (
                    $existing.Subject -eq $Subject -and
                    $existing.Thumbprint -ne $Thumbprint
                ) {
                    $rootStore.Remove($existing)
                }
            }
        }
        finally {
            $rootStore.Close()
        }
    }).AddArgument($CertificatePath).AddArgument(
        $certificateSubject
    ).AddArgument($certificate.Thumbprint) | Out-Null

    $currentPid = [System.Diagnostics.Process]::GetCurrentProcess().Id
    $async = $powerShell.BeginInvoke()
    try {
        while (-not $async.IsCompleted) {
            [ExcelBroForeground]::BringDialogToFront($currentPid) | Out-Null
            Start-Sleep -Milliseconds 150
        }
        $powerShell.EndInvoke($async)
    }
    finally {
        $powerShell.Dispose()
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

function Register-ExcelBroAddin {
    param([Parameter(Mandatory = $true)][string]$ManifestPath)

    $manifestText = [IO.File]::ReadAllText(
        $ManifestPath,
        [Text.UTF8Encoding]::new($false)
    )
    [xml]$manifestXml = $manifestText
    $manifestId = [string]$manifestXml.OfficeApp.Id
    $manifestVersion = [string]$manifestXml.OfficeApp.Version

    $storeId = "\\$env:COMPUTERNAME\$shareName"
    $storeIdJson = $storeId.Replace("\", "\\")

    $wefRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\WEF"
    $filesystemExcelDir = Join-Path $wefRoot "AddinInfo\1\filesystem\Excel"
    $addinDir = Join-Path $filesystemExcelDir "1\${manifestId}_${manifestVersion}"
    New-Item -ItemType Directory -Force -Path $addinDir | Out-Null

    $bootJson = (
        '{"solid":"' + $manifestId +
        '","storeid":"' + $storeIdJson +
        '","appversion":"' + $manifestVersion +
        '","equivalentaddins":[]}'
    )
    [IO.File]::WriteAllText(
        (Join-Path $addinDir "boot.json"),
        $bootJson,
        [Text.UTF8Encoding]::new($false)
    )

    $metaPath = Join-Path $filesystemExcelDir "meta.json"
    $unixNow = [int64][DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $metaJson = '{"minorversion":1,"lastupdate":' + $unixNow + '}'
    [IO.File]::WriteAllText(
        $metaPath,
        $metaJson,
        [Text.UTF8Encoding]::new($false)
    )

    # 功能区（AppCommands）只有在加载项被「添加」到受信任目录后才会生成。
    # 「添加」状态记录在 Providers\<hash>\AppStates\<id>_<version>，其中
    # <hash> = base64(MD5(UTF-16LE(目录 URL)))，'/' 替换为 '_'。
    $providerHash = [Convert]::ToBase64String(
        [Security.Cryptography.MD5]::Create().ComputeHash(
            [Text.Encoding]::Unicode.GetBytes($storeId)
        )
    ).Replace('/', '_')

    $providerKey = (
        "HKCU:\Software\Microsoft\Office\16.0\WEF\Providers\" + $providerHash
    )
    if (-not (Test-Path $providerKey)) {
        New-Item -Path $providerKey -Force | Out-Null
    }
    New-ItemProperty `
        -Path $providerKey `
        -Name "UniqueId" `
        -Value $storeId `
        -PropertyType String `
        -Force | Out-Null

    New-ItemProperty `
        -Path $providerKey `
        -Name "Entitlements" `
        -Value ([DateTime]::UtcNow.ToFileTime()) `
        -PropertyType QWord `
        -Force | Out-Null

    $appStatesKey = Join-Path $providerKey "AppStates"
    if (-not (Test-Path $appStatesKey)) {
        New-Item -Path $appStatesKey -Force | Out-Null
    }
    New-ItemProperty `
        -Path $appStatesKey `
        -Name "${manifestId}_${manifestVersion}" `
        -Value ([DateTime]::UtcNow.ToFileTime()) `
        -PropertyType QWord `
        -Force | Out-Null
}

function Get-WefProviderHash {
    param([Parameter(Mandatory = $true)][string]$UniqueId)

    return [Convert]::ToBase64String(
        [Security.Cryptography.MD5]::Create().ComputeHash(
            [Text.Encoding]::Unicode.GetBytes($UniqueId)
        )
    ).Replace('/', '_')
}

function Remove-StaleWefState {
    $wefRoot = "HKCU:\Software\Microsoft\Office\16.0\WEF"
    $cacheRoot = Join-Path `
        $env:LOCALAPPDATA `
        "Microsoft\Office\16.0\WEF\{96F651A8-A028-49B0-BA44-B9088B0ACC78}"

    # 功能区（AppCommands）只由共享文件夹 (TrustedCatalog) 的「已添加」状态
    # 驱动；旧的开发者旁加载直连（Wef\Developer）不会生成功能区，还会把同一
    # 清单 Id 抢占过去。这里连根移除，让 Excel 只走共享文件夹路径。
    $developerKey = Join-Path $wefRoot "Developer"
    if (Test-Path $developerKey) {
        Remove-Item -Path $developerKey -Recurse -Force
    }

    $stalePrefixes = @("developer", "\\localhost\", "\\127.0.0.1\")
    $providersRoot = Join-Path $wefRoot "Providers"
    if (Test-Path $providersRoot) {
        foreach ($provider in Get-ChildItem $providersRoot) {
            $uniqueId = (Get-ItemProperty `
                -Path $provider.PSPath `
                -Name "UniqueId" `
                -ErrorAction SilentlyContinue).UniqueId
            if ([string]::IsNullOrWhiteSpace($uniqueId)) {
                continue
            }
            $stale = $false
            foreach ($prefix in $stalePrefixes) {
                if ($uniqueId.StartsWith(
                    $prefix,
                    [StringComparison]::OrdinalIgnoreCase
                )) {
                    $stale = $true
                    break
                }
            }
            if (-not $stale) {
                continue
            }
            Remove-Item -Path $provider.PSPath -Recurse -Force

            $hash = Get-WefProviderHash $uniqueId
            $cacheDir = Join-Path $cacheRoot $hash
            if (Test-Path -LiteralPath $cacheDir) {
                Remove-Item -LiteralPath $cacheDir -Recurse -Force
            }
        }
    }
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
        Remove-StaleWefState
        Register-ExcelBroAddin $manifestPath
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

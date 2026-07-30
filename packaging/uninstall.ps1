param(
    [switch]$Silent
)

$ErrorActionPreference = "Stop"

# Resolve the actual installation directory from this script. This keeps the
# uninstaller working if Windows is installed on a non-default drive.
$installDir = $PSScriptRoot
$manifestId = "9c758d40-c2b8-42d8-a6bc-735bd5c4f34c"
$catalogId = "{41f62f5c-cd95-44f2-a0c6-c8cb847fe4e0}"
$catalogShareName = "ExcelBroAddins"
$developerKey = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
$trustedCatalogKey = (
    "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\$catalogId"
)
$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"

& taskkill.exe /F /IM ExcelBro.exe *> $null
for ($attempt = 0; $attempt -lt 20; $attempt += 1) {
    if (-not (Get-Process -Name "ExcelBro" -ErrorAction SilentlyContinue)) {
        break
    }
    Start-Sleep -Milliseconds 250
}

# Creating and removing an SMB share requires administrator approval. Do this
# before removing the installation records so a cancelled UAC prompt leaves a
# retryable installation instead of a partial uninstall.
$removeShareCommand = @"
`$ErrorActionPreference = 'Stop'
`$share = Get-SmbShare -Name '$catalogShareName' -ErrorAction SilentlyContinue
if (`$null -ne `$share) {
    Remove-SmbShare -Name '$catalogShareName' -Force -Confirm:`$false
}
"@
$encodedRemoveShareCommand = [Convert]::ToBase64String(
    [Text.Encoding]::Unicode.GetBytes($removeShareCommand)
)
$shareProcess = Start-Process `
    -FilePath "powershell.exe" `
    -Verb RunAs `
    -WindowStyle Hidden `
    -ArgumentList "-NoProfile", "-EncodedCommand", $encodedRemoveShareCommand `
    -PassThru `
    -Wait
if ($shareProcess.ExitCode -ne 0) {
    throw "无法移除 Excel Bro 本机加载项目录。"
}

Remove-ItemProperty -Path $developerKey -Name $manifestId -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $trustedCatalogKey -Recurse -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $runKey -Name "Excel Bro" -ErrorAction SilentlyContinue
Remove-Item `
    -LiteralPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel Bro" `
    -Recurse `
    -Force `
    -ErrorAction SilentlyContinue

$certificatePath = Join-Path $installDir "localhost.cer"
if (Test-Path -LiteralPath $certificatePath) {
    $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $certificatePath
    )
    & certutil.exe -user -delstore Root $certificate.Thumbprint *> $null
}

# Data directory (%LOCALAPPDATA%\Excel Bro) holds the log, saved model
# connections and API keys. A clean uninstall removes it too so nothing is left
# behind. Kept separate from $installDir (%LOCALAPPDATA%\Programs\Excel Bro).
$dataDir = Join-Path $env:LOCALAPPDATA "Excel Bro"

# Run final deletion from a separate process so this script never tries to
# delete its own active directory. EncodedCommand avoids quoting failures when
# the installation path contains spaces or apostrophes.
$escapedInstallDir = $installDir.Replace("'", "''")
$escapedDataDir = $dataDir.Replace("'", "''")
$cleanupCommand = @"
`$target = '$escapedInstallDir'
`$dataTarget = '$escapedDataDir'
Remove-Item -LiteralPath `$dataTarget -Recurse -Force -ErrorAction SilentlyContinue
for (`$attempt = 0; `$attempt -lt 30; `$attempt += 1) {
    Start-Sleep -Milliseconds 500
    Remove-Item -LiteralPath `$target -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath `$target)) { exit 0 }
}
exit 1
"@
$encodedCommand = [Convert]::ToBase64String(
    [Text.Encoding]::Unicode.GetBytes($cleanupCommand)
)
Start-Process `
    -FilePath "powershell.exe" `
    -WindowStyle Hidden `
    -ArgumentList "-NoProfile", "-EncodedCommand", $encodedCommand

exit 0

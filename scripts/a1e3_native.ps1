# ASCII-only: also works in Windows PowerShell 5.1.
[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [ValidateSet("prepare", "status", "arm", "cancel", "evidence", "stop", "frontend")]
    [string]$Command,
    [string]$HostId = "",
    [string]$Name = "",
    [string]$Identity = "",
    [string]$Base = "",
    [ValidateRange(1024,65535)][int]$Port = 8877,
    [string]$Manifest = "",
    [string]$Out = ""
)
$explicitPort = $PSBoundParameters.ContainsKey("Port")
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"
$cliPath = Join-Path $PSScriptRoot "a1e3_native.py"
$runsRoot = Join-Path $projectRoot "outputs\a1e3-native\runs"
$latestPath = Join-Path $projectRoot "outputs\a1e3-native\latest.json"

function Save-Json($Path, $Value) {
    $json = $Value | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText($Path, $json, (New-Object Text.UTF8Encoding($false)))
}
function Process-Identity([int]$ProcessId) {
    $item = Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId" -ErrorAction Stop
    if (-not $item) { return $null }
    return @{ pid = $item.ProcessId; created = $item.CreationDate.ToUniversalTime().ToString("o");
              executable = $item.ExecutablePath; command = $item.CommandLine }
}
function Stop-OwnedServer($Run, [string]$Path) {
    if (-not $Run.process) { Write-Host "[stop] No recorded process."; return }
    $recorded = $Run.process
    $actual = Process-Identity ([int]$recorded.pid)
    if (-not $actual) { Write-Host "[stop] Recorded process has exited."; return }
    # JSON readers may normalize fractional-second precision; compare the instant, not its spelling.
    $sameCreated = $false
    try {
        $sameCreated = ([DateTimeOffset]$actual.created).UtcTicks -eq
                       ([DateTimeOffset]$recorded.created).UtcTicks
    } catch {
        # Malformed timestamps are an identity mismatch, never permission to stop.
        $sameCreated = $false
    }
    if (-not $sameCreated -or $actual.executable -ne $recorded.executable -or
        $actual.command -ne $recorded.command -or
        $actual.command -notlike '*scripts.a1e3_native_server:create_app*') {
        throw "Process identity mismatch. Refusing to stop PID $($recorded.pid)."
    }
    & taskkill.exe /PID $recorded.pid /T /F | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "taskkill failed: $LASTEXITCODE" }
    Write-Host "[stop] Stopped the verified test process tree only."
}
function Load-Run {
    if (-not $script:Manifest) {
        if (-not (Test-Path -LiteralPath $latestPath)) { throw "Run prepare first, or provide -Manifest." }
        $script:Manifest = (Get-Content -LiteralPath $latestPath -Raw | ConvertFrom-Json).manifest
    }
    $script:Manifest = (Resolve-Path -LiteralPath $script:Manifest).Path
    $loaded = Get-Content -LiteralPath $script:Manifest -Raw | ConvertFrom-Json
    if ($Base -and $Base -ne $loaded.base) { throw "Base disagrees with manifest." }
    if ($explicitPort -and $Port -ne $loaded.port) { throw "Port disagrees with manifest." }
    return $loaded
}
if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) { throw "Project venv not found." }
if ($Command -eq "prepare") {
    if ($Port -eq 8765) { throw "8765 is reserved for the user's service." }
    $expectedBase = "http://127.0.0.1:$Port"
    if ($Base -and $Base -ne $expectedBase) { throw "Base must agree with Port and use 127.0.0.1." }
    if (Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue) {
        throw "Port $Port is occupied. Existing process and evidence were not touched."
    }
    $runId = (Get-Date -Format "yyyyMMddTHHmmss") + "-" + [guid]::NewGuid().ToString("N").Substring(0,8)
    $runDir = Join-Path $runsRoot $runId
    New-Item -ItemType Directory -Path $runDir | Out-Null
    $configDir = Join-Path $runDir "server-config"
    New-Item -ItemType Directory -Path $configDir | Out-Null
    $sourceConfig = if ($env:EXCEL_BRO_CONFIG_DIR) { $env:EXCEL_BRO_CONFIG_DIR } else { Join-Path $projectRoot "server" }
    foreach ($file in @("model-connections.json", ".env")) {
        $source = Join-Path $sourceConfig $file
        if (Test-Path -LiteralPath $source -PathType Leaf) { Copy-Item -LiteralPath $source -Destination $configDir }
    }
    & $pythonPath (Join-Path $PSScriptRoot "make_a1e3_native_fixture.py") --out $runDir
    if ($LASTEXITCODE -ne 0) { throw "Fixture creation failed." }
    $Manifest = Join-Path $runDir "manifest.json"
    $sdkStateDir = Join-Path $env:TEMP ("excel-bro-native-sdk-" + $runId)
    New-Item -ItemType Directory -Path $sdkStateDir | Out-Null
    $run = @{ sdkStateDir=$sdkStateDir; runId=$runId; base=$expectedBase; port=$Port; token=([guid]::NewGuid().ToString("N")); process=$null }
    Save-Json $Manifest $run
    $keys = @("EXCEL_BRO_CONFIG_DIR", "EXCEL_BRO_SOURCE_CHANGED_FAULT", "CLAUDE_AGENT_SDK_STATE_DIR", "EB_NATIVE_MANIFEST")
    $previous = @{}
    foreach ($key in $keys) { $previous[$key] = [Environment]::GetEnvironmentVariable($key, "Process") }
    try {
        $env:EXCEL_BRO_CONFIG_DIR = $configDir
        $env:EXCEL_BRO_SOURCE_CHANGED_FAULT = "1"
        $env:CLAUDE_AGENT_SDK_STATE_DIR = $sdkStateDir
        $env:EB_NATIVE_MANIFEST = $Manifest
        $proc = Start-Process -FilePath $pythonPath -ArgumentList @("-m", "uvicorn", "scripts.a1e3_native_server:create_app", "--factory", "--host", "127.0.0.1", "--port", "$Port", "--loop", "asyncio", "--workers", "1") -WorkingDirectory $projectRoot -RedirectStandardOutput (Join-Path $runDir "server.log") -RedirectStandardError (Join-Path $runDir "server.err.log") -WindowStyle Hidden -PassThru
        $run.process = Process-Identity $proc.Id
        if (-not $run.process) { throw "Server exited before process identity could be captured." }
        Save-Json $Manifest $run
    } finally {
        foreach ($key in $keys) { [Environment]::SetEnvironmentVariable($key, $previous[$key], "Process") }
    }
    $ready = $false
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        try {
            $state = Invoke-RestMethod "$expectedBase/api/dev/native-acceptance/session" -Headers @{"X-Native-Acceptance-Token"=$run.token} -TimeoutSec 2
            if ($state.runId -ne $runId) { throw "Wrong run ID" }
            $ready = $true
            break
        } catch { Start-Sleep -Milliseconds 400 }
    }
    if (-not $ready) { Stop-OwnedServer $run $Manifest; throw "Server not ready. Logs retained at $runDir" }
    Save-Json $latestPath @{ manifest=$Manifest }
    Write-Host "[ready] $expectedBase; run=$runId"
    Write-Host "[manifest] $Manifest"
    Write-Host "Next: run this script frontend -HostId excel-office (or wps-et) in another terminal."
    exit 0
}
$run = Load-Run
if ($Command -eq "stop") { Stop-OwnedServer $run $Manifest; exit 0 }
if ($Command -eq "frontend") {
    if ($HostId -notin @("excel-office", "wps-et")) { throw "frontend requires -HostId excel-office or wps-et." }
    $frontendPort = if ($HostId -eq "wps-et") { 3889 } else { 3000 }
    if (Get-NetTCPConnection -LocalPort $frontendPort -State Listen -ErrorAction SilentlyContinue) { throw "Frontend port occupied. No process was stopped." }
    $oldApi = $env:VITE_API_BASE_URL
    $oldProxy = $env:EB_TEST_API_TARGET
    Push-Location $projectRoot
    try {
        $env:VITE_API_BASE_URL = $run.base
        $env:EB_TEST_API_TARGET = $run.base
        $workspace = if ($HostId -eq "wps-et") { "excel-bro-wps" } else { "@excel-bro/addin" }
        & npm.cmd run dev --workspace $workspace
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
        [Environment]::SetEnvironmentVariable("VITE_API_BASE_URL", $oldApi, "Process")
        [Environment]::SetEnvironmentVariable("EB_TEST_API_TARGET", $oldProxy, "Process")
    }
    exit $code
}
$cliArgs = @($cliPath, $Command, "--manifest", $Manifest)
if ($Command -in @("arm", "evidence") -and $HostId) { $cliArgs += @("--host", $HostId) }
if ($Command -eq "arm") {
    if ($Name) { $cliArgs += @("--name", $Name) }
    if ($Identity) { $cliArgs += @("--identity", $Identity) }
}
if ($Command -eq "evidence") {
    if (-not $Out) { $Out = Join-Path (Split-Path -Parent $Manifest) "evidence" }
    $cliArgs += @("--out", $Out)
}
& $pythonPath @cliArgs
exit $LASTEXITCODE

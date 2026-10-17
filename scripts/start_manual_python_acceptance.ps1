# ASCII-only for Windows PowerShell 5.1.
[CmdletBinding()]
param([switch]$PreflightOnly)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"
$outputDir = Join-Path $projectRoot "outputs\manual-acceptance-20261010"
$capabilitiesPath = Join-Path $outputDir "server-capabilities.json"
$previousCapabilities = [Environment]::GetEnvironmentVariable("EXCEL_BRO_CAPABILITIES_PATH", "Process")

if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw "Project virtual environment not found."
}

try {
    $capabilities = Get-Content -LiteralPath (Join-Path $projectRoot "config\capabilities.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    $capabilities.pythonExecution.enabled = $true
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    [IO.File]::WriteAllText($capabilitiesPath, ($capabilities | ConvertTo-Json -Depth 100), (New-Object Text.UTF8Encoding($false)))
    $env:EXCEL_BRO_CAPABILITIES_PATH = $capabilitiesPath

    Push-Location $projectRoot
    try {
        # Read-only engine/image checks; no model request and no analysis execution.
        & $pythonPath -c "from server.app.python_analysis import _runner_status, python_execution_enabled; import importlib.util; ready = _runner_status(); print('Verified Python runner:', ready or 'UNAVAILABLE'); print('Python enabled for this launch:', python_execution_enabled()); print('Agent SDK installed:', importlib.util.find_spec('claude_agent_sdk') is not None); raise SystemExit(0 if ready and python_execution_enabled() else 2)"
        if ($LASTEXITCODE -ne 0) {
            throw "Verified Docker engine or pinned sandbox image unavailable. No host fallback is permitted."
        }
        Write-Host "Manual acceptance capabilities: $capabilitiesPath"
        Write-Host "Saved model connections are reused. No model double is installed."
        Write-Host "Default config/capabilities.json is unchanged."
        if (-not $PreflightOnly) {
            & (Join-Path $PSScriptRoot "start_claude_sdk_server.ps1")
        }
    }
    finally { Pop-Location }
}
finally {
    [Environment]::SetEnvironmentVariable("EXCEL_BRO_CAPABILITIES_PATH", $previousCapabilities, "Process")
}

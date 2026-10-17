# Keep this file ASCII-only: Windows PowerShell 5.1 reads UTF-8 without BOM as ANSI.
$ErrorActionPreference = "Stop"

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).Path
$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw "Project virtual environment not found. Install server/requirements-claude-sdk.txt first."
}

$listener = Get-NetTCPConnection -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue
if ($listener) {
    throw "Port 8765 is already in use. Stop the existing backend (including npm run dev:server) before starting the SDK server."
}

$previousStateDir = [Environment]::GetEnvironmentVariable("CLAUDE_AGENT_SDK_STATE_DIR", "Process")

try {
    if ([string]::IsNullOrWhiteSpace($env:CLAUDE_AGENT_SDK_STATE_DIR)) {
        $localAppData = [Environment]::GetFolderPath("LocalApplicationData")
        if ([string]::IsNullOrWhiteSpace($localAppData)) {
            throw "Local app data directory not found. Set CLAUDE_AGENT_SDK_STATE_DIR manually."
        }
        $env:CLAUDE_AGENT_SDK_STATE_DIR = Join-Path $localAppData "ExcelBro\ClaudeAgentSdk"
    }

    Push-Location $projectRoot
    try {
        # SDK availability does not imply that isolated Python tools are enabled.
        & $pythonPath -c "from server.app.python_analysis import python_execution_enabled; enabled = python_execution_enabled(); print('Isolated Python tools enabled:', enabled); print('For Python analysis use scripts/start_manual_python_acceptance.ps1' if not enabled else 'Python tools will use the configured verified isolated runner.')"
        if ($LASTEXITCODE -ne 0) {
            throw "Unable to read the effective Python capability configuration."
        }
        # Windows reload/multiple workers use a Selector loop without subprocess support.
        & $pythonPath -m uvicorn server.app.main:app --host 127.0.0.1 --port 8765 --loop asyncio --workers 1
        if ($LASTEXITCODE -ne 0) {
            throw "Local server exited with code $LASTEXITCODE."
        }
    }
    finally {
        Pop-Location
    }
}
finally {
    [Environment]::SetEnvironmentVariable("CLAUDE_AGENT_SDK_STATE_DIR", $previousStateDir, "Process")
}

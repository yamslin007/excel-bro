"""Exercise the actual PowerShell stop function without stopping any OS process."""
import shutil
import subprocess
from pathlib import Path

import pytest


@pytest.mark.skipif(not shutil.which("pwsh"), reason="PowerShell 7 unavailable")
def test_stop_matches_timestamp_instant_and_preserves_identity_guards():
    script = r"""
$ErrorActionPreference = 'Stop'
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PWD 'scripts/a1e3_native.ps1'), [ref]$null, [ref]$null)
$func = $ast.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Stop-OwnedServer'}, $true)
Invoke-Expression $func.Extent.Text
$script:actual = @{pid=42;created='2026-10-26T05:20:25.4412640Z';executable='python.exe';command='python -m scripts.a1e3_native_server:create_app'}
function Process-Identity { param($ProcessId) return $script:actual }
$script:kills = 0
function taskkill.exe { $script:kills++; $global:LASTEXITCODE=0 }
foreach ($stamp in @('2026-10-26T05:20:25.441264Z', [DateTime]::Parse('2026-10-26T05:20:25.441264Z').ToUniversalTime())) {
    $run=@{process=@{pid=42;created=$stamp;executable='python.exe';command=$script:actual.command}}
    Stop-OwnedServer $run ''
}
if ($script:kills -ne 2) { throw 'Equivalent string/JSON DateTime rejected' }
foreach ($field in @('created','executable','command')) {
    $record=@{pid=42;created='2026-10-26T05:20:25.441264Z';executable='python.exe';command=$script:actual.command}
    $record[$field]= if ($field -eq 'created') {'2026-10-26T05:20:25.441265Z'} else {'wrong'}
    try { Stop-OwnedServer @{process=$record} ''; throw 'Expected refusal' }
    catch { if ($_.Exception.Message -notlike 'Process identity mismatch*') { throw } }
}
if ($script:kills -ne 2) { throw 'Mismatched process would have been stopped' }
exit 0
"""
    result = subprocess.run(
        [shutil.which("pwsh"), "-NoProfile", "-NonInteractive", "-Command", script],
        cwd=Path(__file__).resolve().parents[2], capture_output=True, text=True, timeout=20,
    )
    assert result.returncode == 0, result.stdout + result.stderr

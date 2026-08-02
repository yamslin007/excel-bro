from __future__ import annotations

from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
BUILD_SCRIPT_PATH = PROJECT_ROOT / "packaging" / "build_installer.ps1"
INNO_SCRIPT_PATH = PROJECT_ROOT / "packaging" / "excel_bro.iss"
TASK_SCRIPT_PATH = PROJECT_ROOT / "packaging" / "install_tasks.ps1"
MANIFEST_PATH = PROJECT_ROOT / "apps" / "excel-addin" / "manifest.xml"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_build_uses_standard_inno_setup_compiler() -> None:
    script = _read(BUILD_SCRIPT_PATH)

    assert "Find-InnoSetupCompiler" in script
    assert "ISCC.exe" in script
    assert "excel_bro.iss" in script
    assert "/DSourceDir=" in script
    assert "--onefile" not in script
    assert "installer.py" not in script


def test_installer_shows_normal_directory_and_ready_pages() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "DefaultDirName={localappdata}\\Programs\\{#AppName}" in script
    assert "DisableDirPage=no" in script
    assert "DisableReadyPage=no" in script
    assert "DisableFinishedPage=no" in script
    assert "WizardStyle=modern" in script
    assert "AllowNetworkDrive=no" in script
    assert "AllowRootDirectory=no" in script
    assert "AllowUNCPath=no" in script


def test_installer_force_closes_only_its_background_runtime_on_upgrade() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "CloseApplications=force" in script
    assert "CloseApplicationsFilter={#AppExeName}" in script
    assert "CloseApplicationsFilter=EXCEL.EXE" not in script


def test_installer_uses_standard_uninstaller_and_start_menu_entry() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "Uninstallable=yes" in script
    assert "CreateUninstallRegKey=yes" in script
    assert 'Filename: "{uninstallexe}"' in script
    assert "CurUninstallStepChanged" in script
    assert "IntegrationParameters('Uninstall')" in script


def test_installer_is_per_user_and_preserves_model_configuration() -> None:
    inno_script = _read(INNO_SCRIPT_PATH)
    task_script = _read(TASK_SCRIPT_PATH)

    assert "PrivilegesRequired=lowest" in inno_script
    assert "Root: HKCU" in inno_script
    assert "model-connections.json" not in inno_script
    assert 'Join-Path $env:LOCALAPPDATA "Excel Bro"' not in task_script


def test_packaged_manifest_meets_office_minimum_version_and_icon_requirements() -> None:
    manifest = _read(MANIFEST_PATH)
    build_script = _read(BUILD_SCRIPT_PATH)

    assert "<Version>1.0.0.0</Version>" in manifest
    assert '<IconUrl DefaultValue="https://localhost:3000/' in manifest
    assert '<HighResolutionIconUrl DefaultValue="https://localhost:3000/' in manifest
    assert '$manifestVersion -lt [Version]"1.0.0.0"' in build_script
    assert "IconUrl 和 HighResolutionIconUrl" in build_script


def test_catalog_share_is_scoped_and_requires_elevation() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "New-SmbShare" in script
    assert "-Name $shareLiteral" in script
    assert "-Path $catalogLiteral" in script
    assert "-ReadAccess $principalLiteral" in script
    assert "-Verb RunAs" in script
    assert "Remove-SmbShare" in script


def test_custom_local_install_directory_is_allowed_but_root_is_rejected() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert 'GetFileName($resolved) -ne "Excel Bro"' not in script
    assert "不能把 Excel Bro 直接安装到磁盘根目录" in script


def test_uninstall_rejects_an_unexpected_share_path() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "共享路径与当前安装目录不一致，拒绝删除" in script
    assert "$actualPath.Equals(" in script
    assert "[StringComparison]::OrdinalIgnoreCase" in script


def test_certificate_cleanup_covers_previous_installer_builds() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert '$certificateSubject = "CN=Excel Bro localhost"' in script
    assert "Install-RootCertificate" in script
    assert "Remove-RootCertificates" in script
    assert "StoreLocation]::CurrentUser" in script
    assert "$store.Remove(" in script


def test_runtime_is_health_checked_during_install() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "Start-ExcelBroRuntime" in script
    assert "http://127.0.0.1:8765/health" in script
    assert "本地服务未能启动" in script
    assert "$runtimeProcess.HasExited" in script
    assert "$listener.OwningProcess -eq $runtimeProcess.Id" in script


def test_legacy_installer_is_migrated_without_touching_user_data() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert "Remove-LegacyInstall" in script
    assert 'Join-Path $legacyDir "uninstall.ps1"' in script
    assert "Remove-Item -LiteralPath $legacyDir -Recurse -Force" in script
    assert 'Join-Path $env:LOCALAPPDATA "Excel Bro"' not in script


def test_all_old_excel_bro_processes_are_stopped_before_health_check() -> None:
    script = _read(TASK_SCRIPT_PATH)

    assert 'Get-Process -Name "ExcelBro"' in script
    assert "Stop-Process -Force" in script


def test_catalog_url_uses_the_actual_computer_name() -> None:
    script = _read(INNO_SCRIPT_PATH)

    assert "GetEnv('COMPUTERNAME')" in script
    assert "\\ExcelBroAddins" in script

from __future__ import annotations

from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
BUILD_SCRIPT_PATH = PROJECT_ROOT / "packaging" / "build_mac.sh"
PREINSTALL_PATH = PROJECT_ROOT / "packaging" / "mac" / "pkg-scripts" / "preinstall"
POSTINSTALL_PATH = PROJECT_ROOT / "packaging" / "mac" / "pkg-scripts" / "postinstall"
UNINSTALL_PATH = PROJECT_ROOT / "packaging" / "mac" / "uninstall.sh"
RUNTIME_PATH = PROJECT_ROOT / "packaging" / "runtime.py"
WORKFLOW_PATH = PROJECT_ROOT / ".github" / "workflows" / "build-mac.yml"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def test_build_script_guards_manifest_and_uses_onedir_pyinstaller() -> None:
    script = _read(BUILD_SCRIPT_PATH)

    assert "1, 0, 0, 0" in script
    assert "IconUrl 和 HighResolutionIconUrl" in script
    assert "--onedir" in script
    assert "--onefile" not in script
    assert "--noconsole" not in script
    assert "pkgbuild" in script
    assert "productbuild" in script
    assert "--identifier com.excelbro.app" in script


def test_build_script_stages_frontend_manifest_and_certificate() -> None:
    script = _read(BUILD_SCRIPT_PATH)

    assert 'Applications/Excel Bro' in script
    assert "frontend" in script
    assert "catalog/manifest.xml" in script
    assert "localhost.*" in script
    assert "uninstall.sh" in script


def test_postinstall_sideloads_manifest_into_excel_container_as_console_user() -> None:
    script = _read(POSTINSTALL_PATH)

    assert "com.microsoft.Excel/Data/Documents/wef" in script
    assert "stat -f %Su /dev/console" in script
    assert 'sudo -u "$CONSOLE_USER" mkdir -p "$WEF_DIR"' in script
    assert 'sudo -u "$CONSOLE_USER" cp "$INSTALL_DIR/manifest.xml"' in script


def test_postinstall_trusts_certificate_and_replaces_previous_ones() -> None:
    script = _read(POSTINSTALL_PATH)

    assert 'CERT_SUBJECT="Excel Bro localhost"' in script
    assert "security add-trusted-cert -d -r trustRoot" in script
    assert "security delete-certificate" in script
    assert "/Library/Keychains/System.keychain" in script


def test_postinstall_registers_launchagent_and_health_checks_with_rollback() -> None:
    script = _read(POSTINSTALL_PATH)

    assert "com.excelbro.runtime" in script
    assert "LaunchAgents" in script
    assert "launchctl bootstrap" in script
    assert "RunAtLoad" in script
    assert "http://127.0.0.1:8765/health" in script
    assert "未能启动" in script
    assert "launchctl bootout" in script


def test_preinstall_stops_runtime_and_removes_launchagent() -> None:
    script = _read(PREINSTALL_PATH)

    assert "pkill -x ExcelBro" in script
    assert "launchctl bootout" in script
    assert "com.excelbro.runtime" in script


def test_uninstall_removes_integration_but_preserves_user_configuration() -> None:
    script = _read(UNINSTALL_PATH)

    assert "launchctl bootout" in script
    assert "wef/manifest.xml" in script
    assert "security delete-certificate" in script
    assert 'rm -rf "$INSTALL_DIR"' in script
    assert "Application Support/Excel Bro" in script
    assert "保留" in script
    # 配置目录只能被提及，不能被删除
    assert 'rm -rf "$HOME/Library/Application Support' not in script


def test_runtime_uses_application_support_on_macos() -> None:
    script = _read(RUNTIME_PATH)

    assert 'sys.platform == "darwin"' in script
    assert 'Path.home() / "Library" / "Application Support" / "Excel Bro"' in script


def test_workflow_builds_both_mac_architectures() -> None:
    workflow = _read(WORKFLOW_PATH)

    assert "macos-13" in workflow
    assert "macos-14" in workflow
    assert "workflow_dispatch" in workflow
    assert "bash packaging/build_mac.sh" in workflow
    assert "dist/Excel-Bro-Setup-*.pkg" in workflow

from __future__ import annotations

import base64
import importlib.util
import os
from pathlib import Path
from unittest.mock import patch


INSTALLER_PATH = (
    Path(__file__).resolve().parents[2] / "packaging" / "installer.py"
)
SPEC = importlib.util.spec_from_file_location(
    "excel_bro_packaging_installer",
    INSTALLER_PATH,
)
assert SPEC is not None
assert SPEC.loader is not None
installer = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(installer)

def test_powershell_literal_escapes_apostrophes() -> None:
    assert installer._powershell_literal("C:\\O'Brien") == "'C:\\O''Brien'"


def test_elevated_script_encoding_uses_windows_powershell_unicode() -> None:
    script = "Write-Output 'Excel Bro'"

    encoded = installer._encode_powershell(script)

    assert base64.b64decode(encoded).decode("utf-16-le") == script


def test_catalog_share_script_scopes_access_to_current_user() -> None:
    script = installer._catalog_share_script(
        Path(r"C:\Users\O'Brien\Excel Bro\catalog"),
        r"WORKGROUP\O'Brien",
    )

    assert "New-SmbShare" in script
    assert "-Name 'ExcelBroAddins'" in script
    assert "-Path 'C:\\Users\\O''Brien\\Excel Bro\\catalog'" in script
    assert "-ReadAccess 'WORKGROUP\\O''Brien'" in script
    assert "Remove-SmbShare" in script


def test_current_windows_principal_uses_original_user_environment() -> None:
    with patch.dict(
        os.environ,
        {"USERDOMAIN": "WORKGROUP", "USERNAME": "个人用户"},
        clear=False,
    ):
        assert installer._current_windows_principal() == "WORKGROUP\\个人用户"


def test_catalog_url_uses_computer_name_instead_of_localhost() -> None:
    with patch.dict(
        os.environ,
        {"COMPUTERNAME": "PERSONAL-PC"},
        clear=False,
    ):
        assert (
            installer._catalog_url()
            == r"\\PERSONAL-PC\ExcelBroAddins"
        )

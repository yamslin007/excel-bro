from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

from server.app.folder_mutex import FolderWriteBusyError, folder_write_guard


def test_same_work_directory_cannot_be_locked_twice(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    root = tmp_path / "work"
    root.mkdir()

    with folder_write_guard(root):
        with pytest.raises(FolderWriteBusyError, match="正在被另一项"):
            with folder_write_guard(root):
                pass


def test_different_work_directories_can_be_locked_together(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    first = tmp_path / "first"
    second = tmp_path / "second"
    first.mkdir()
    second.mkdir()

    with folder_write_guard(first):
        with folder_write_guard(second):
            assert not (first / ".excel-bro-backups").exists()
            assert not (second / ".excel-bro-backups").exists()


def test_cross_process_lock_rejects_second_writer(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    app_data = tmp_path / "app-data"
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(app_data))
    root = tmp_path / "work"
    root.mkdir()
    script = """
import sys
from pathlib import Path
from server.app.folder_mutex import FolderWriteBusyError, folder_write_guard
try:
    with folder_write_guard(Path(sys.argv[1])):
        pass
except FolderWriteBusyError:
    raise SystemExit(23)
raise SystemExit(0)
"""
    environment = os.environ.copy()
    environment["EXCEL_BRO_CONFIG_DIR"] = str(app_data)

    with folder_write_guard(root):
        completed = subprocess.run(
            [sys.executable, "-c", script, str(root)],
            cwd=str(Path(__file__).resolve().parents[2]),
            env=environment,
            check=False,
            capture_output=True,
            text=True,
        )

    assert completed.returncode == 23, completed.stderr

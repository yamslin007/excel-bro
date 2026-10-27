"""Keep test collection and default ledgers out of the user's application data."""
import os
import tempfile
from pathlib import Path
import pytest

_KEYS = ("EXCEL_BRO_CONFIG_DIR", "CLAUDE_AGENT_SDK_STATE_DIR")

def pytest_configure(config):
    config._eb_previous_env = {key: os.environ.get(key) for key in _KEYS}
    root = Path(tempfile.mkdtemp(prefix="excel-bro-pytest-isolated-"))
    config._eb_isolated_root = root
    os.environ["EXCEL_BRO_CONFIG_DIR"] = str(root / "config")
    os.environ["CLAUDE_AGENT_SDK_STATE_DIR"] = str(root / "sdk")

def pytest_unconfigure(config):
    for key, value in getattr(config, "_eb_previous_env", {}).items():
        if value is None:
            os.environ.pop(key, None)
        else:
            os.environ[key] = value

@pytest.fixture(autouse=True)
def isolated_a1_ledger(request, monkeypatch, tmp_path):
    if Path(str(request.node.fspath)).name in {
        "test_host_write_observation.py", "test_host_write_audit.py",
        "test_source_changed_stop.py", "test_a1e3_native_regressions.py",
    }:
        from server.app import main
        from server.app.task_ledger import FolderOperationLedger
        monkeypatch.setattr(main, "_folder_operation_ledger",
                            FolderOperationLedger(lambda: tmp_path / "state.sqlite3"))

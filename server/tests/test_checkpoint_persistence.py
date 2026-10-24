"""Fault injection for state-file replacement; never use a native workbook."""
import errno
import json
from pathlib import Path

import pytest

from server.app import worksheet_checkpoint as checkpoint
from test_worksheet_restore import env  # noqa: F401 -- shared synthetic workbook fixture
from test_worksheet_restore import SheetModel
from test_worksheet_checkpoint import Book, Native, request


def deny_replace(monkeypatch, fail):
    original = Path.replace
    attempts = []

    def replace(source, target):
        if Path(target).name == "state.json":
            record = json.loads(source.read_text(encoding="utf-8"))
            attempts.append(record)
            if fail(record, len(attempts)):
                raise PermissionError(errno.EACCES, "synthetic lock")
        return original(source, target)

    monkeypatch.setattr(Path, "replace", replace)
    monkeypatch.setattr(checkpoint.time, "sleep", lambda _: None)
    return attempts


def test_transient_replace_retries_only_persistence(monkeypatch, tmp_path):
    attempts = deny_replace(monkeypatch, lambda _record, n: n < 3)
    checkpoint._write(tmp_path, {"revision": 2})
    assert len(attempts) == 3
    assert json.loads((tmp_path / "state.json").read_text()) == {"revision": 2}


def test_persistent_replace_keeps_last_committed_state(monkeypatch, tmp_path):
    checkpoint._write(tmp_path, {"revision": 1})
    attempts = deny_replace(monkeypatch, lambda *_: True)
    with pytest.raises(OSError):
        checkpoint._write(tmp_path, {"revision": 2})
    assert 1 < len(attempts) <= 10
    assert json.loads((tmp_path / "state.json").read_text()) == {"revision": 1}


def test_restore_progress_lock_does_not_repeat_copy(env, monkeypatch):
    service, native, book, payload = env
    attempts = deny_replace(monkeypatch, lambda record, n:
        (record["result"].get("restore") or {}).get("restoredSheets") == ["Sheet2"]
        and n < 4)
    result = service.restore(payload.operationId, "https://pane")
    assert result.status == "completed", result.message
    assert len(attempts) >= 4
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2", "Sheet3", "Sheet4"]
    assert service.restore(payload.operationId, "https://pane") == result


def test_failed_restore_progress_is_unknown_after_reload_not_replayed(env, monkeypatch):
    service, native, book, payload = env
    original_replace = Path.replace
    # Intent can commit; every write after the actual first copy is denied.
    deny_replace(monkeypatch, lambda record, _n:
        bool((record["result"].get("restore") or {}).get("restoredSheets")))
    with pytest.raises(OSError):
        service.restore(payload.operationId, "https://pane")
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2"]
    monkeypatch.setattr(Path, "replace", original_replace)
    reloaded = checkpoint.WorksheetCheckpointService(native)
    status = reloaded.status(payload.operationId, "https://pane")
    assert status.restore is not None
    assert status.restore.status != "completed"
    assert status.restore.unverifiedSheets == ["Sheet2"]
    retried = reloaded.restore(payload.operationId, "https://pane")
    assert retried.unverifiedSheets == ["Sheet2"]
    assert "核对" in retried.message
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2"]
    assert reloaded.open(payload.operationId, "https://pane")["opened"]


def test_restore_intent_must_commit_before_copy(env, monkeypatch):
    service, native, book, payload = env
    deny_replace(monkeypatch, lambda record, _n: bool(record.get("restorePendingSheet")))
    with pytest.raises(OSError):
        service.restore(payload.operationId, "https://pane")
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1"]


def test_copy_then_throw_is_unknown_not_unchanged_and_never_replayed(env, monkeypatch):
    service, native, book, payload = env
    original_copy = SheetModel.Copy
    calls = []

    def copy_then_throw(sheet, Before=None, After=None):
        calls.append(sheet.Name)
        original_copy(sheet, Before, After)
        raise RuntimeError("host lost reply after mutation")

    monkeypatch.setattr(SheetModel, "Copy", copy_then_throw)
    result = service.restore(payload.operationId, "https://pane")
    assert result.status == "partial"
    assert result.unverifiedSheets == ["Sheet2"]
    assert "原工作簿未改动" not in result.message
    assert "未放回任何" not in result.message
    assert service.restore(payload.operationId, "https://pane").unverifiedSheets == ["Sheet2"]
    assert calls == ["Sheet2"]


def test_restore_api_returns_fixed_error_then_unknown_status(env, monkeypatch):
    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from server.app.worksheet_checkpoint_api import create_worksheet_checkpoint_router

    service, native, book, payload = env
    app = FastAPI()
    app.include_router(create_worksheet_checkpoint_router(["https://pane"], service))
    client = TestClient(app, base_url="http://127.0.0.1", client=("127.0.0.1", 5000))
    original_replace = Path.replace
    deny_replace(monkeypatch, lambda record, _n:
        bool((record["result"].get("restore") or {}).get("restoredSheets")))
    response = client.post("/api/worksheet-checkpoints/restore",
        json={"operationId": payload.operationId}, headers={"origin": "https://pane"})
    assert response.status_code == 409
    assert "暂时无法保存" in response.json()["detail"]
    assert "synthetic lock" not in response.text
    monkeypatch.setattr(Path, "replace", original_replace)
    status = client.post("/api/worksheet-checkpoints/status",
        json={"operationId": payload.operationId}, headers={"origin": "https://pane"})
    assert status.status_code == 200
    assert status.json()["restore"]["unverifiedSheets"] == ["Sheet2"]
    assert [sheet.Name for sheet in book.sheets] == ["Sheet1", "Sheet2"]


@pytest.mark.parametrize("host", ["excel", "wps-et"])
@pytest.mark.parametrize("after_delete", [False, True])
def test_delete_persistence_failure_is_not_replayed(tmp_path, monkeypatch, host, after_delete):
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path))
    native = Native(Book())
    service = checkpoint.WorksheetCheckpointService(native)
    payload = request(host=host)
    original_replace = Path.replace
    deny_replace(monkeypatch, lambda record, _n:
        bool(record["result"]["deletedSheets"]) if after_delete
        else record["result"]["status"] != "prepared")
    with pytest.raises(OSError):
        service.execute(payload, "https://pane")
    assert native.book.calls == (["Sheet2"] if after_delete else [])
    monkeypatch.setattr(Path, "replace", original_replace)
    reloaded = checkpoint.WorksheetCheckpointService(native)
    result = reloaded.status(payload.operationId, "https://pane")
    assert result.status in {"failed", "partial_or_unknown"}
    assert reloaded.execute(payload, "https://pane").status == result.status
    assert native.book.calls == (["Sheet2"] if after_delete else [])
    if after_delete:
        assert result.recoveryAvailable
        assert reloaded.open(payload.operationId, "https://pane")["opened"]

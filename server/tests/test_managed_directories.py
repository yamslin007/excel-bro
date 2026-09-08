from __future__ import annotations

import json
from pathlib import Path

import pytest

import server.app.main as main_module
from server.app.folder_workbooks import inspect_managed_folder_recovery
from server.app.task_ledger import FolderOperationLedger


def test_directory_registry_survives_restart_and_honors_removal(tmp_path) -> None:
    database = tmp_path / "app-data" / "tasks" / "state.sqlite3"
    root = tmp_path / "work"
    root.mkdir()
    first = FolderOperationLedger(lambda: database)
    registered = first.register_directory(str(root), max_entries=2)

    restarted = FolderOperationLedger(lambda: database)
    records = restarted.managed_directories()

    assert len(records) == 1
    assert records[0].directory_id == registered.directory_id
    assert Path(records[0].path) == root
    assert restarted.remove_directory(registered.directory_id) is True
    assert restarted.managed_directories() == []
    assert restarted.managed_directories(include_removed=True)[0].status == "removed"


def test_directory_registry_enforces_configured_capacity(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    first = tmp_path / "first"
    second = tmp_path / "second"
    first.mkdir()
    second.mkdir()
    ledger.register_directory(str(first), max_entries=1)

    with pytest.raises(RuntimeError, match="达到上限"):
        ledger.register_directory(str(second), max_entries=1)


def test_managed_scan_of_empty_directory_creates_no_backup_storage(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    root = tmp_path / "work"
    root.mkdir()
    main_module._folder_operation_ledger.register_directory(str(root))

    records = main_module._scan_registered_work_directories()

    assert len(records) == 1
    assert records[0].status == "active"
    assert records[0].recovery_count == 0
    assert records[0].attention_count == 0
    assert not (root / ".excel-bro-backups").exists()


def test_managed_scan_marks_missing_directory_unavailable_without_replacement(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    missing = tmp_path / "missing-work"
    replacement = tmp_path / "elsewhere" / "missing-work"
    replacement.mkdir(parents=True)
    registered = main_module._folder_operation_ledger.register_directory(str(missing))

    records = main_module._scan_registered_work_directories()

    record = next(item for item in records if item.directory_id == registered.directory_id)
    assert record.status == "unavailable"
    assert record.path == str(missing)
    assert record.last_error is not None
    assert "不存在" in record.last_error
    assert not (replacement / ".excel-bro-backups").exists()


def test_managed_scan_finds_interrupted_execution_and_recovery_transaction(
    tmp_path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    root = tmp_path / "work"
    execution_id = "a" * 32
    operation_id = "recovery-operation-1"
    execution = root / ".excel-bro-backups" / "executions" / execution_id
    operation = root / ".excel-bro-backups" / "operations" / operation_id
    execution.mkdir(parents=True)
    operation.mkdir(parents=True)
    (execution / "manifest.json").write_text(
        json.dumps(
            {
                "version": 2,
                "undoId": execution_id,
                "status": "applying",
                "createdAt": "2026-09-08T10:00:00+08:00",
                "planId": "plan-1",
                "planTitle": "测试计划",
                "files": [
                    {
                        "originalRelativePath": "result.xlsx",
                        "backupRelativePath": None,
                        "recoveryAction": "remove",
                        "sizeBytes": 0,
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    (operation / "manifest.json").write_text(
        json.dumps(
            {
                "version": 1,
                "operationId": operation_id,
                "operationType": "undo",
                "status": "applying",
                "files": [],
            }
        ),
        encoding="utf-8",
    )

    inspection = inspect_managed_folder_recovery(root)

    assert inspection["recoveryCount"] == 1
    assert inspection["operationCount"] == 1
    assert inspection["attentionCount"] == 2

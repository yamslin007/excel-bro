from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import os
from pathlib import Path
import sqlite3
import subprocess
import sys

import pytest

from server.app.task_ledger import ExpiredOperationError, FolderOperationLedger


def test_register_is_atomic_and_reuses_existing_identity(tmp_path) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    ledger = FolderOperationLedger(lambda: database)

    def register() -> bool:
        _record, created = ledger.register("operation-1", "fingerprint-1")
        return created

    with ThreadPoolExecutor(max_workers=8) as pool:
        created_values = list(pool.map(lambda _index: register(), range(16)))

    assert created_values.count(True) == 1
    record = ledger.get("operation-1")
    assert record is not None
    assert record.request_fingerprint == "fingerprint-1"
    assert record.status == "running"


def test_ledger_replays_events_and_result_after_new_instance(tmp_path) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    first = FolderOperationLedger(lambda: database, executor_id="first-process")
    first.register("operation-2", "fingerprint-2")
    first.append_event(
        "operation-2",
        1,
        {"operationId": "operation-2", "sequence": 1, "title": "已完成第一步"},
    )
    first.finish(
        "operation-2",
        status="completed",
        result={"filesModified": ["结果.xlsx"]},
    )

    restarted = FolderOperationLedger(lambda: database, executor_id="second-process")
    record = restarted.get("operation-2")

    assert record is not None
    assert record.status == "completed"
    assert record.last_sequence == 1
    assert record.result == {"filesModified": ["结果.xlsx"]}
    assert restarted.events("operation-2") == [
        {"operationId": "operation-2", "sequence": 1, "title": "已完成第一步"}
    ]


def test_ledger_rejects_event_gaps_and_terminal_rewrite(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    ledger.register("operation-3", "fingerprint-3")

    with pytest.raises(RuntimeError, match="序号不连续"):
        ledger.append_event("operation-3", 2, {"sequence": 2})

    ledger.finish(
        "operation-3",
        status="needs_reconciliation",
        error={"code": "UNKNOWN_EFFECT"},
    )
    with pytest.raises(RuntimeError, match="状态已经改变"):
        ledger.finish(
            "operation-3",
            status="completed",
            result={"filesModified": []},
        )


def test_same_operation_id_keeps_original_fingerprint(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    original, created = ledger.register("operation-4", "fingerprint-a")
    existing, created_again = ledger.register("operation-4", "fingerprint-b")

    assert created is True
    assert created_again is False
    assert original.request_fingerprint == "fingerprint-a"
    assert existing.request_fingerprint == "fingerprint-a"


def test_abrupt_child_process_exit_keeps_committed_event_for_reconciliation(
    tmp_path,
) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    script = """
import os
import sys
from pathlib import Path
from server.app.task_ledger import FolderOperationLedger

database = Path(sys.argv[1])
ledger = FolderOperationLedger(lambda: database, executor_id="crashed-process")
ledger.register("operation-crash", "fingerprint-crash")
ledger.append_event(
    "operation-crash",
    1,
    {"operationId": "operation-crash", "sequence": 1, "title": "写入意图已登记"},
)
os._exit(23)
"""
    environment = os.environ.copy()
    completed = subprocess.run(
        [sys.executable, "-c", script, str(database)],
        cwd=str(Path(__file__).resolve().parents[2]),
        env=environment,
        check=False,
    )

    restarted = FolderOperationLedger(
        lambda: database, executor_id="replacement-process"
    )
    record = restarted.get("operation-crash")

    assert completed.returncode == 23
    assert record is not None
    assert record.status == "running"
    assert record.executor_id == "crashed-process"
    assert restarted.events("operation-crash") == [
        {
            "operationId": "operation-crash",
            "sequence": 1,
            "title": "写入意图已登记",
        }
    ]


def test_terminal_compaction_keeps_permanent_idempotency_tombstone(tmp_path) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    ledger = FolderOperationLedger(lambda: database)
    ledger.register("operation-expired", "fingerprint-expired")
    ledger.append_event("operation-expired", 1, {"sequence": 1, "detail": "large"})
    ledger.finish(
        "operation-expired",
        status="completed",
        result={"filesModified": ["结果.xlsx"]},
    )
    with sqlite3.connect(database) as connection:
        connection.execute(
            "UPDATE folder_operations SET updated_at = 0 WHERE operation_id = ?",
            ("operation-expired",),
        )

    assert ledger.compact_terminal_operations(30) == 1
    assert ledger.get("operation-expired") is None
    assert ledger.events("operation-expired") == []
    assert ledger.expired_fingerprint("operation-expired") == "fingerprint-expired"
    with pytest.raises(ExpiredOperationError):
        ledger.register("operation-expired", "fingerprint-expired")

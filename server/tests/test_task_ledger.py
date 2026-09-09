from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import os
from pathlib import Path
import sqlite3
import subprocess
import sys

import pytest

import server.app.task_ledger as task_ledger_module
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


def test_execution_context_is_immutable_and_checkpoint_survives_restart(
    tmp_path,
) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    first = FolderOperationLedger(lambda: database, executor_id="first-process")
    first.register("operation-checkpoint", "fingerprint-checkpoint")
    first.save_execution_context(
        "operation-checkpoint",
        request={"sessionId": "session-1", "plan": {"id": "plan-1"}},
        authorization={
            "root": str(tmp_path / "work"),
            "targets": [{"reference": "book.xlsx › Data"}],
        },
    )
    committed = first.commit_checkpoint(
        "operation-checkpoint",
        kind="folder_publish",
        payload={"schemaVersion": 1, "stage": "publish_prepared"},
    )

    assert committed.generation == 1
    with pytest.raises(RuntimeError, match="确认请求不可被替换"):
        first.save_execution_context(
            "operation-checkpoint",
            request={"sessionId": "different", "plan": {"id": "plan-1"}},
            authorization={"root": str(tmp_path / "work"), "targets": []},
        )

    restarted = FolderOperationLedger(
        lambda: database, executor_id="replacement-process"
    )
    record = restarted.get("operation-checkpoint")
    checkpoint = restarted.checkpoint("operation-checkpoint")

    assert record is not None
    assert record.request == {"sessionId": "session-1", "plan": {"id": "plan-1"}}
    assert record.authorization == {
        "root": str(tmp_path / "work"),
        "targets": [{"reference": "book.xlsx › Data"}],
    }
    assert checkpoint is not None
    assert checkpoint.payload["stage"] == "publish_prepared"


def test_checkpoint_rejects_payload_over_configured_limit(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    ledger.register("operation-large-checkpoint", "fingerprint")
    monkeypatch.setattr(task_ledger_module, "capability_int", lambda *_args: 32)

    with pytest.raises(ValueError, match="容量上限"):
        ledger.commit_checkpoint(
            "operation-large-checkpoint",
            kind="test",
            payload={"value": "x" * 100},
        )

    assert ledger.checkpoint("operation-large-checkpoint") is None


def test_resume_claim_is_compare_and_swap_and_fences_old_executor(tmp_path) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    original = FolderOperationLedger(lambda: database, executor_id="old-process")
    record, _created = original.register("operation-claim", "fingerprint-claim")
    record = original.save_execution_context(
        "operation-claim",
        request={"sessionId": "session-1", "plan": {"id": "plan-1"}},
        authorization={"root": str(tmp_path / "work"), "targets": []},
    )
    original.commit_checkpoint(
        "operation-claim",
        kind="folder_publish",
        payload={"schemaVersion": 1, "stage": "publish_prepared"},
    )
    record = original.get("operation-claim")
    assert record is not None

    replacement = FolderOperationLedger(
        lambda: database, executor_id="replacement-process"
    )
    claimed = replacement.claim_for_resume(
        "operation-claim", expected_revision=record.revision
    )
    assert claimed.executor_id == "replacement-process"
    assert claimed.revision == record.revision + 1

    with pytest.raises(RuntimeError, match="过期执行者"):
        original.commit_checkpoint(
            "operation-claim",
            kind="folder_publish",
            payload={"schemaVersion": 1, "stage": "published"},
        )
    with pytest.raises(RuntimeError, match="状态已变化"):
        replacement.claim_for_resume(
            "operation-claim", expected_revision=record.revision
        )


def test_planning_pause_survives_restart_and_extension_request_is_idempotent(
    tmp_path,
) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    first = FolderOperationLedger(lambda: database, executor_id="planning-first")
    record, created = first.register_planning_task(
        "turn-planning-1",
        "request-fingerprint",
        {"turnId": "turn-planning-1", "prompt": "分析数据"},
        approved_wall_time_seconds=60,
    )
    assert created is True
    first.append_planning_event(
        record.turn_id,
        1,
        {"operationId": record.turn_id, "sequence": 1, "title": "已完成读取"},
    )
    checkpoint = {
        "schemaVersion": 1,
        "kind": "planning",
        "stage": "tool_batch_completed",
        "messages": [{"role": "user", "content": "分析数据"}],
        "toolCache": {"read:1": "result"},
        "loopState": {"activeElapsedSeconds": 60},
        "toolMode": "terminal",
        "repairingSubmission": False,
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
    }
    paused = first.pause_planning_task(
        record.turn_id,
        checkpoint,
        active_elapsed_seconds=60,
        reason="等待用户继续",
    )

    restarted = FolderOperationLedger(
        lambda: database, executor_id="planning-restarted"
    )
    restored = restarted.planning_task(record.turn_id)
    assert restored is not None
    assert restored.status == "waiting_user"
    assert restored.checkpoint == checkpoint
    assert restarted.planning_events(record.turn_id)[0]["sequence"] == 1

    claimed, should_run = restarted.claim_planning_continuation(
        record.turn_id,
        expected_revision=paused.revision,
        request_id="extension-request-0001",
        extension_seconds=60,
        maximum_extensions=4,
        hard_wall_time_seconds=300,
    )
    assert should_run is True
    assert claimed.status == "running"
    assert claimed.approved_wall_time_seconds == 120
    assert claimed.extension_count == 1

    duplicate, duplicate_should_run = restarted.claim_planning_continuation(
        record.turn_id,
        expected_revision=paused.revision,
        request_id="extension-request-0001",
        extension_seconds=60,
        maximum_extensions=4,
        hard_wall_time_seconds=300,
    )
    assert duplicate_should_run is True
    assert duplicate.approved_wall_time_seconds == 120
    assert duplicate.extension_count == 1


def test_planning_continuation_rejects_stale_revision(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    record, _created = ledger.register_planning_task(
        "turn-planning-stale",
        "request-fingerprint",
        {"turnId": "turn-planning-stale", "prompt": "分析数据"},
        approved_wall_time_seconds=60,
    )
    paused = ledger.pause_planning_task(
        record.turn_id,
        {
            "schemaVersion": 1,
            "kind": "planning",
            "loopState": {"activeElapsedSeconds": 60},
        },
        active_elapsed_seconds=60,
        reason="等待用户继续",
    )

    with pytest.raises(RuntimeError, match="状态已变化"):
        ledger.claim_planning_continuation(
            record.turn_id,
            expected_revision=paused.revision - 1,
            request_id="extension-request-stale",
            extension_seconds=60,
            maximum_extensions=4,
            hard_wall_time_seconds=300,
        )

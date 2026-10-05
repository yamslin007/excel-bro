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


def test_planning_model_switch_is_persisted_and_request_id_is_idempotent(
    tmp_path,
) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    ledger = FolderOperationLedger(lambda: database, executor_id="model-switch")
    record, _created = ledger.register_planning_task(
        "turn-model-switch",
        "request-fingerprint",
        {
            "turnId": "turn-model-switch",
            "prompt": "分析数据",
            "modelId": "connection:primary",
        },
        approved_wall_time_seconds=60,
    )
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "before_model",
        "messages": [{"role": "user", "content": "分析数据"}],
        "toolCache": {},
        "loopState": {"activeElapsedSeconds": 18},
        "toolMode": "full",
        "repairingSubmission": False,
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }
    paused = ledger.pause_planning_task(
        record.turn_id,
        checkpoint,
        active_elapsed_seconds=18,
        reason="模型连接中断",
        error={"code": "MODEL_TIMEOUT", "retryable": True},
    )

    claimed, should_run = ledger.claim_planning_continuation(
        record.turn_id,
        expected_revision=paused.revision,
        request_id="model-switch-request-0001",
        extension_seconds=0,
        maximum_extensions=4,
        hard_wall_time_seconds=300,
        model_id="connection:fallback",
        model_switch_reason="用户明确选择备用模型",
    )

    assert should_run is True
    assert claimed.current_model_id == "connection:fallback"
    assert claimed.approved_wall_time_seconds == 60
    assert claimed.extension_count == 0
    assert len(claimed.model_switches) == 1
    assert claimed.model_switches[0]["fromModelId"] == "connection:primary"
    assert claimed.model_switches[0]["toModelId"] == "connection:fallback"

    duplicate, duplicate_should_run = ledger.claim_planning_continuation(
        record.turn_id,
        expected_revision=paused.revision,
        request_id="model-switch-request-0001",
        extension_seconds=0,
        maximum_extensions=4,
        hard_wall_time_seconds=300,
        model_id="connection:fallback",
    )
    assert duplicate_should_run is True
    assert len(duplicate.model_switches) == 1

    with pytest.raises(RuntimeError, match="同一继续请求"):
        ledger.claim_planning_continuation(
            record.turn_id,
            expected_revision=paused.revision,
            request_id="model-switch-request-0001",
            extension_seconds=0,
            maximum_extensions=4,
            hard_wall_time_seconds=300,
            model_id="connection:another",
        )


def test_office_action_checkpoint_reconciles_without_blind_replay(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    plan = {
        "id": "plan-office",
        "title": "写入区域",
        "summary": "写入两个单元格",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {
                "type": "writeValues",
                "sheet": "Sheet1",
                "range": "A1:A1",
                "values": [["完成"]],
            }
        ],
    }
    record, created = ledger.register_office_execution(
        "office-execution-0001",
        "fingerprint-office",
        plan=plan,
        authorization={"authorizedSheetNames": ["Sheet1"]},
        workbook_identity="url:c:/work/report.xlsx",
    )
    assert created is True
    original_owner = record.owner_token
    before = {
        "actionIndex": 0,
        "sheet": "Sheet1",
        "range": "A1:A1",
        "formulas": [[None]],
        "numberFormat": [["General"]],
        "fillColor": "#FFFFFF",
        "fontBold": False,
        "fontColor": "#000000",
    }
    started = ledger.begin_office_action(
        record.operation_id,
        expected_revision=record.revision,
        owner_token=record.owner_token,
        action_index=0,
        before_snapshot=before,
    )
    assert started.in_flight_action_index == 0

    with pytest.raises(RuntimeError, match="先核对在途动作"):
        ledger.claim_office_resume(
            record.operation_id, expected_revision=started.revision
        )

    checkpoint = {
        "planId": "plan-office",
        "executionId": record.operation_id,
        "workbookIdentity": "url:c:/work/report.xlsx",
        "capturedAt": "2026-09-09T00:00:00Z",
        "coverage": "full",
        "completedActionCount": 1,
        "uncertainActionIndexes": [],
        "postStateCoverage": "complete",
        "ranges": [{**before, "postFingerprint": "fnv1a32:12345678"}],
        "createdWorksheets": [],
        "createdWorksheetStates": [],
        "deletedWorksheets": [],
        "skippedWorksheets": [],
    }
    reconciled = ledger.reconcile_office_action(
        record.operation_id,
        expected_revision=started.revision,
        action_index=0,
        outcome="applied",
        checkpoint=checkpoint,
    )
    assert reconciled.next_action_index == 1
    assert reconciled.in_flight_action_index is None
    assert reconciled.owner_token != original_owner

    with pytest.raises(RuntimeError, match="状态已变化"):
        ledger.begin_office_action(
            record.operation_id,
            expected_revision=record.revision,
            owner_token=original_owner,
            action_index=0,
            before_snapshot=before,
        )

    completed = ledger.finish_office_execution(
        record.operation_id,
        expected_revision=reconciled.revision,
        owner_token=reconciled.owner_token,
        result={"actionResults": [], "verification": {"status": "verified"}},
    )
    assert completed.status == "completed"


def test_office_execution_refuses_unstable_or_changed_workbook_identity(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    plan = {"id": "p", "actions": [{"type": "setFill"}]}
    with pytest.raises(ValueError, match="未保存"):
        ledger.register_office_execution(
            "office-execution-0002",
            "fingerprint",
            plan=plan,
            authorization={"authorizedSheetNames": ["Sheet1"]},
            workbook_identity="unsaved",
        )

    ledger.register_office_execution(
        "office-execution-0002",
        "fingerprint",
        plan=plan,
        authorization={"authorizedSheetNames": ["Sheet1"]},
        workbook_identity="url:c:/work/report.xlsx",
    )
    with pytest.raises(RuntimeError, match="另一个工作簿"):
        ledger.register_office_execution(
            "office-execution-0002",
            "fingerprint",
            plan=plan,
            authorization={"authorizedSheetNames": ["Sheet1"]},
            workbook_identity="url:c:/other/report.xlsx",
        )


def test_wps_execution_registers_with_path_identity_and_task_id(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    plan = {
        "id": "plan-wps",
        "title": "写入区域",
        "summary": "写入两个单元格",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {
                "type": "writeValues",
                "sheet": "Sheet1",
                "range": "A1:A1",
                "values": [["完成"]],
            }
        ],
    }
    record, created = ledger.register_office_execution(
        "wps-execution-0001",
        "fingerprint-wps",
        plan=plan,
        authorization={"authorizedSheetNames": ["Sheet1"]},
        workbook_identity="wps-path:c:/work/report.xlsx",
        identity_prefix="wps-path:",
        task_id="task-keep",
    )
    assert created is True
    assert record.task_id == "task-keep"
    assert record.workbook_identity.startswith("wps-path:")

    # 同 operationId 重放（身份、指纹、taskId 一致）→ 不新建，返回既有记录。
    replay, recreated = ledger.register_office_execution(
        "wps-execution-0001",
        "fingerprint-wps",
        plan=plan,
        authorization={"authorizedSheetNames": ["Sheet1"]},
        workbook_identity="wps-path:c:/work/report.xlsx",
        identity_prefix="wps-path:",
        task_id="task-keep",
    )
    assert recreated is False
    assert replay.operation_id == record.operation_id

    # 同 operationId 但绑定不同 taskId → 拒绝（幂等身份不得翻新）。
    with pytest.raises(RuntimeError, match="另一个任务"):
        ledger.register_office_execution(
            "wps-execution-0001",
            "fingerprint-wps",
            plan=plan,
            authorization={"authorizedSheetNames": ["Sheet1"]},
            workbook_identity="wps-path:c:/work/report.xlsx",
            identity_prefix="wps-path:",
            task_id="task-other",
        )


def test_wps_execution_refuses_unprefixed_or_office_identity(tmp_path) -> None:
    ledger = FolderOperationLedger(lambda: tmp_path / "state.sqlite3")
    plan = {"id": "p", "actions": [{"type": "writeValues"}]}

    # WPS 身份没有 wps-path: 前缀 → 拒绝（未保存 / 无法稳定识别）。
    with pytest.raises(ValueError, match="未保存"):
        ledger.register_office_execution(
            "wps-execution-0002",
            "fingerprint",
            plan=plan,
            authorization={"authorizedSheetNames": ["Sheet1"]},
            workbook_identity="c:/work/report.xlsx",
            identity_prefix="wps-path:",
        )

    # 同 operationId 先用 wps-path: 登记，再用 url: 重放 → 身份前缀不同，
    # 完整字符串比较必然命中"另一个工作簿"，Office 与 WPS 互不冒充。
    ledger.register_office_execution(
        "wps-execution-0003",
        "fingerprint",
        plan=plan,
        authorization={"authorizedSheetNames": ["Sheet1"]},
        workbook_identity="wps-path:c:/work/report.xlsx",
        identity_prefix="wps-path:",
    )
    with pytest.raises(RuntimeError, match="另一个工作簿"):
        ledger.register_office_execution(
            "wps-execution-0003",
            "fingerprint",
            plan=plan,
            authorization={"authorizedSheetNames": ["Sheet1"]},
            workbook_identity="url:c:/work/report.xlsx",
            identity_prefix="url:",
        )

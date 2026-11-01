from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import os
from pathlib import Path
import sqlite3
import subprocess
import sys
import time

import pytest

import server.app.task_ledger as task_ledger_module
from server.app.task_ledger import ExpiredOperationError, FolderOperationLedger


@pytest.mark.parametrize("stored_version", [9, 10])
def test_existing_ledger_without_recovery_deadline_migrates_without_rewriting_execution(tmp_path, stored_version):
    database = tmp_path / "state.sqlite3"
    first = FolderOperationLedger(lambda: database)
    original, _ = first.register_office_execution("old-execution", "unchanged-fingerprint",
        plan={"actions": []}, authorization={"authorizedSheetNames": ["Target"]},
        workbook_identity="url:existing.xlsx", task_id="existing-task")
    # Reproduce the shipped v9 database, including its existing execution row.
    # A version label alone must not conceal the missing column either.
    with sqlite3.connect(database) as connection:
        connection.execute("ALTER TABLE office_executions DROP COLUMN recovery_package_expires_at")
        connection.execute(f"PRAGMA user_version = {stored_version}")
    restarted = FolderOperationLedger(lambda: database)
    assert restarted.reserve_recovery_package_expiry("old-execution", "2035-01-01T00:00:00Z") == "2035-01-01T00:00:00Z"
    assert restarted.reserve_recovery_package_expiry("old-execution", "2036-01-01T00:00:00Z") == "2035-01-01T00:00:00Z"
    assert restarted.office_execution("old-execution") == original
    with sqlite3.connect(database) as connection:
        assert connection.execute("PRAGMA user_version").fetchone()[0] >= 10
        assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"


def test_migration_keeps_action_progress_and_checkpoint(tmp_path) -> None:
    """补列不许碰到执行本身：进度、检查点、计划、授权、修订号都得原样。

    真实故障里这条列是**出厂旧库缺的**，不是"某次迁移写坏了"。补列是纯 schema 动作，
    一旦它在 `INSERT ... SELECT` 式的重建里丢过一列，用户看到的会是"任务还在、进度归零"，
    那比 PUT 500 更难察觉。
    """

    database = tmp_path / "state.sqlite3"
    first = FolderOperationLedger(lambda: database)
    record, _ = first.register_office_execution(
        "in-progress",
        "unchanged-fingerprint",
        plan={"actions": [{"type": "writeValues", "sheet": "Target", "range": "A1"}]},
        authorization={"authorizedSheetNames": ["Target"]},
        workbook_identity="wps-path:c:\\data\\in-progress.xlsx",
        identity_prefix="wps-path:",
        task_id="task-in-progress",
    )
    begun = first.begin_office_action(
        "in-progress",
        expected_revision=record.revision,
        owner_token=record.owner_token,
        action_index=0,
        before_snapshot={"cells": [["old"]]},
    )
    done = first.complete_office_action(
        "in-progress",
        expected_revision=begun.revision,
        owner_token=record.owner_token,
        action_index=0,
        checkpoint={"cells": [["new"]]},
    )
    assert done.next_action_index == 1
    assert done.checkpoint == {"cells": [["new"]]}

    # 还原成出厂 v9 形态：版本号旧，而且列**真的**不存在（只改版本号骗不过检查）。
    with sqlite3.connect(database) as connection:
        connection.execute("ALTER TABLE office_executions DROP COLUMN recovery_package_expires_at")
        connection.execute("PRAGMA user_version = 9")

    restarted = FolderOperationLedger(lambda: database)
    migrated = restarted.office_execution("in-progress")
    assert migrated == done
    assert migrated is not None
    assert migrated.next_action_index == 1
    assert migrated.checkpoint == {"cells": [["new"]]}
    assert migrated.before_snapshot is None
    # 迁移只补列，不伪造期限：旧库从来没登记过，所以这里仍是"未登记"。
    assert restarted.recovery_package_expiry("in-progress") is None
    # 补列之后"首次写赢"照旧：第一次写入生效，第二次不许推后。
    assert restarted.reserve_recovery_package_expiry(
        "in-progress", "2036-01-01T00:00:00Z"
    ) == "2036-01-01T00:00:00Z"
    assert restarted.reserve_recovery_package_expiry(
        "in-progress", "2037-01-01T00:00:00Z"
    ) == "2036-01-01T00:00:00Z"
    with sqlite3.connect(database) as connection:
        assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"


def test_migration_keeps_an_existing_deadline_when_only_the_version_lags(tmp_path) -> None:
    """版本号落后但列已经在：只能升级版本号，不许重建表、也不许清掉已登记的期限。"""

    database = tmp_path / "state.sqlite3"
    ledger = FolderOperationLedger(lambda: database)
    ledger.register_office_execution(
        "version-leg",
        "fingerprint",
        plan={"actions": []},
        authorization={"authorizedSheetNames": ["Target"]},
        workbook_identity="url:version-leg.xlsx",
    )
    assert ledger.reserve_recovery_package_expiry(
        "version-leg", "2035-01-01T00:00:00Z"
    ) == "2035-01-01T00:00:00Z"
    with sqlite3.connect(database) as connection:
        connection.execute("PRAGMA user_version = 9")

    restarted = FolderOperationLedger(lambda: database)
    assert restarted.recovery_package_expiry("version-leg") == "2035-01-01T00:00:00Z"
    # 已经登记过的期限不会被重启后的任何一次 PUT 推开。
    assert restarted.reserve_recovery_package_expiry(
        "version-leg", "2036-01-01T00:00:00Z"
    ) == "2035-01-01T00:00:00Z"
    with sqlite3.connect(database) as connection:
        assert connection.execute("PRAGMA user_version").fetchone()[0] >= 10


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


def test_first_initialization_under_concurrency_never_reports_readonly(
    tmp_path,
) -> None:
    """首次并发落库必须在 Windows/WAL 下保持可写，不得抛出只读库错误。

    真实缺陷：旧实现在建表连接尚未关闭时就让其它连接打开同一文件，
    SQLite 在 Windows 上会把并发打开的连接持成只读句柄，
    报 "attempt to write a readonly database"。这里用多轮并发首建复现。
    """

    failures: list[str] = []
    for round_index in range(12):
        database = tmp_path / f"round-{round_index}" / "tasks" / "state.sqlite3"
        ledger = FolderOperationLedger(lambda: database)

        def register(index: int = round_index) -> bool:
            _record, created = ledger.register(
                f"operation-{index}", f"fingerprint-{index}"
            )
            return created

        try:
            with ThreadPoolExecutor(max_workers=8) as pool:
                results = list(pool.map(lambda _index: register(), range(16)))
        except sqlite3.OperationalError as error:  # pragma: no cover - 缺陷路径
            failures.append(f"round {round_index}: {error}")
            continue
        assert results.count(True) == 1

    assert failures == []


def test_read_is_not_serialized_behind_concurrent_write_lock(tmp_path) -> None:
    """第十批 F4：纯读取不得被另一连接的写事务串行化。

    主连接持有未提交写事务时，另一实例的 ``get()``（只读）必须**立即**返回，
    不能等写锁释放、也不能抛 "database is locked"。旧实现在每次 _connect 都用
    BEGIN IMMEDIATE 探测可写，导致纯读也去抢写锁。
    """
    import time

    database = tmp_path / "tasks" / "state.sqlite3"
    writer = FolderOperationLedger(lambda: database)
    writer.register("op-1", "fp-1")

    connection = None
    writer._ensure_schema(writer._path())
    connection = writer._open_writable_connection(writer._path())
    connection.execute("BEGIN IMMEDIATE")
    connection.execute(
        "INSERT INTO folder_operations(operation_id, request_fingerprint, status, "
        "created_at, updated_at, executor_id) VALUES ('op-held','fp-held','running',?,?,?)",
        (time.time(), time.time(), "holder"),
    )
    try:
        reader = FolderOperationLedger(lambda: database)
        started = time.monotonic()
        record = reader.get("op-1")
        elapsed = time.monotonic() - started
        assert record is not None and record.status == "running"
        assert elapsed < 0.5, f"只读被写锁串行化：{elapsed:.3f}s"
        # 新加的"首次期限"读口读的是同一张表的新列，同样不许绕道去抢写锁 ——
        # 否则"补了一列"就把这条既有保证重新打穿了。
        expiry_started = time.monotonic()
        assert reader.recovery_package_expiry("op-1") is None
        assert time.monotonic() - expiry_started < 0.5, "recovery_package_expiry 抢了写锁"
    finally:
        connection.rollback()
        connection.close()


def test_new_instance_on_existing_db_does_not_collide_with_open_writer(tmp_path) -> None:
    """第十批 F4：已有库上新建实例不得因重跑 bootstrap 与写锁冲突。

    旧实现每个实例都会重跑 ``_initialize_schema``（WAL 切换 / 迁移），
    在另一连接持写事务时直接抛 "database is locked"。
    """
    import time

    database = tmp_path / "tasks" / "state.sqlite3"
    first = FolderOperationLedger(lambda: database)
    first.register("op-1", "fp-1")

    connection = first._open_writable_connection(first._path())
    connection.execute("BEGIN IMMEDIATE")
    connection.execute(
        "INSERT INTO folder_operations(operation_id, request_fingerprint, status, "
        "created_at, updated_at, executor_id) VALUES ('op-held2','fp-held2','running',?,?,?)",
        (time.time(), time.time(), "holder"),
    )
    try:
        second = FolderOperationLedger(lambda: database)  # 新实例
        record = second.get("op-1")  # 应能读，不抛 database is locked
        assert record is not None
    finally:
        connection.rollback()
        connection.close()


def test_read_after_write_commit_reflects_new_state(tmp_path) -> None:
    """第十批 F4：写入提交后，只读连接能看到新状态（读写并发正确性）。"""
    database = tmp_path / "tasks" / "state.sqlite3"
    ledger = FolderOperationLedger(lambda: database)
    ledger.register("op-1", "fp-1")
    assert ledger.get("op-1").status == "running"
    ledger.finish("op-1", status="completed", result={"ok": True})
    record = ledger.get("op-1")
    assert record is not None and record.status == "completed"


def test_readonly_connection_release_after_exception(tmp_path) -> None:
    """第十批 F4：只读连接在异常后必须被释放，不泄漏句柄。"""
    database = tmp_path / "tasks" / "state.sqlite3"
    ledger = FolderOperationLedger(lambda: database)
    ledger.register("op-1", "fp-1")

    # 读取不存在的记录不抛异常，返回 None
    assert ledger.get("does-not-exist") is None
    # 连续多次读取仍成功（句柄正确释放）
    for _ in range(5):
        assert ledger.get("op-1") is not None


def test_writable_connection_retries_on_readonly_error_code(tmp_path, monkeypatch) -> None:
    """第十批 F4：可写连接重试须基于**准确错误类别**（sqlite_errorcode），非字符串。"""
    import server.app.task_ledger as module

    calls = {"n": 0}
    real_connect = sqlite3.connect

    class FakeReadonlyOnce:
        pass

    # 直接验证 _is_readonly_error 按 errorcode 判定
    err = sqlite3.OperationalError("attempt to write a readonly database")
    # 真实异常带 sqlite_errorcode；构造一个带 errorcode 的替身
    class Coded(sqlite3.OperationalError):
        sqlite_errorcode = 8
        sqlite_errorname = "SQLITE_READONLY"

    assert module._is_readonly_error(Coded("x")) is True
    # 非只读错误码不得被判为只读
    class Other(sqlite3.OperationalError):
        sqlite_errorcode = 5  # SQLITE_BUSY
        sqlite_errorname = "SQLITE_BUSY"

    assert module._is_readonly_error(Other("database is locked")) is False



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
    # 2026-10-10 时间策略改造：一次"继续"即**结清到总活跃执行时限**
    # （hard_wall_time_seconds），不再每次只加 extension_seconds。旧语义下初批
    # 预算等于提醒阈值（60 秒），按 60 秒递增会让用户被迫点几十次 —— 那正是
    # "被短时预算反复打断"的另一面。
    assert claimed.approved_wall_time_seconds == 300
    assert claimed.extension_count == 1

    # 同一 request_id 重复领取是幂等的：不重复加预算，也不重复计次数。
    duplicate, duplicate_should_run = restarted.claim_planning_continuation(
        record.turn_id,
        expected_revision=paused.revision,
        request_id="extension-request-0001",
        extension_seconds=60,
        maximum_extensions=4,
        hard_wall_time_seconds=300,
    )
    assert duplicate_should_run is True
    assert duplicate.approved_wall_time_seconds == 300
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


@pytest.mark.parametrize("stored_version", [9, 10, 11])
def test_existing_ledger_without_discovery_column_migrates_without_touching_the_execution(
    tmp_path, stored_version
):
    """发现处置列与恢复期限列同样受"版本号是意图、列是事实"约束。

    旧库缺 `discovery_dismissed_at` 时，发现查询会直接撞 `no such column`；而版本号
    已经被写成 11 的库若只看版本号就早退，补列会永久跳过。这里把三种出厂形态都还原
    一遍：迁移只加列，**执行本身（进度、在途动作、检查点、修订号、执行权令牌）原样**。
    """

    database = tmp_path / "state.sqlite3"
    first = FolderOperationLedger(lambda: database)
    record, _ = first.register_office_execution(
        "discovery-legacy",
        "unchanged-fingerprint",
        plan={"actions": [
            {"actionIndex": 0, "type": "writeValues", "sheet": "Target", "range": "A1:A1",
             "rowCount": 1, "columnCount": 1, "cellCount": 1,
             "previewFingerprint": "v2:before-0", "expectedFingerprint": "v2:after-0"},
        ]},
        authorization={"authorizedSheetNames": ["Target"]},
        workbook_identity="wps-path:c://data//legacy.xlsx",
        identity_prefix="wps-path:",
        task_id="task-legacy",
    )
    begun = first.begin_office_action(
        "discovery-legacy",
        expected_revision=record.revision,
        owner_token=record.owner_token,
        action_index=0,
        before_snapshot={"sheetName": "Target", "range": "A1:A1", "type": "writeValues",
                         "beforeFingerprint": "v2:before-0"},
    )
    suspended = first.suspend_office_execution(
        "discovery-legacy",
        expected_revision=begun.revision,
        owner_token=record.owner_token,
        error={"code": "WPS_WRITE_INTERRUPTED", "message": "interrupted"},
    )
    assert suspended.status == "needs_reconciliation"

    with sqlite3.connect(database) as connection:
        connection.execute("ALTER TABLE office_executions DROP COLUMN discovery_dismissed_at")
        connection.execute(f"PRAGMA user_version = {stored_version}")

    restarted = FolderOperationLedger(lambda: database)
    pending, truncated = restarted.pending_office_executions(
        "wps-path:c://data//legacy.xlsx", limit=5
    )
    assert truncated is False
    assert [item.operation_id for item in pending] == ["discovery-legacy"]
    assert restarted.office_execution("discovery-legacy") == suspended
    assert restarted.dismiss_office_execution_discovery("discovery-legacy") is True
    # 处置只写发现元数据：执行本身一个字段都没变。
    assert restarted.office_execution("discovery-legacy") == suspended
    assert restarted.pending_office_executions(
        "wps-path:c://data//legacy.xlsx", limit=5
    )[0] == []
    with sqlite3.connect(database) as connection:
        assert connection.execute("PRAGMA user_version").fetchone()[0] >= 11
        assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"

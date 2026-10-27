"""A1-E3 第四项：**并发恢复时只有一个执行者拿到执行权**。

为什么单独立一个文件：既有用例证明的是「顺序 CAS」——
`test_resume_claim_is_compare_and_swap_and_fences_old_executor`（账本）
与 `test_planning_continuation_rejects_stale_revision` 都是「先认领一次，
再拿旧 revision 认领一次」。那条证据不能代替并发：两个面板**同时**提交认领时，
两者读到的 revision 是一样的，判据必须由数据库的条件更新来裁定，不能靠调用顺序。

三条断言各自钉住一件事：

1. 成功数**恰好为 1**（不是"至少一个"——两个成功就等于两次写入权）；
2. 落败者的 owner token 已被换掉，用旧 token 开动作会被拒（没有第二次宿主写入）；
3. 赢家仍能按**自己拿到的 revision** 开一次动作（认领成功不是"认领后不能干活"）。

规划侧同理：八路并发只有一个追加预算成功，扩展记录只有一条。
"""
from __future__ import annotations

import sqlite3
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from server.app.task_ledger import FolderOperationLedger

PLAN = {
    "id": "plan-claim-race",
    "title": "写入一个单元格",
    "summary": "写入 A1",
    "assumptions": [],
    "warnings": [],
    "actions": [
        {"type": "writeValues", "sheet": "Sheet1", "range": "A1:A1", "values": [["完成"]]}
    ],
}

BEFORE_SNAPSHOT = {
    "actionIndex": 0,
    "sheet": "Sheet1",
    "range": "A1:A1",
    "formulas": [[None]],
    "numberFormat": [["General"]],
    "fillColor": "#FFFFFF",
    "fontBold": False,
    "fontColor": "#000000",
}


def _race(work, count: int = 8) -> tuple[list[object], list[BaseException]]:
    """让 ``count`` 个线程在同一道闸门后同时起跑，收齐成功与失败两侧。"""
    gate = threading.Barrier(count)

    def runner(index: int):
        gate.wait()
        return work(index)

    successes: list[object] = []
    failures: list[BaseException] = []
    with ThreadPoolExecutor(max_workers=count) as pool:
        futures = [pool.submit(runner, index) for index in range(count)]
        for future in futures:
            try:
                successes.append(future.result())
            except BaseException as error:  # noqa: BLE001 - 失败侧就是被测量对象
                failures.append(error)
    return successes, failures


def test_only_one_concurrent_office_resume_claim_wins(tmp_path: Path) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    opener = FolderOperationLedger(lambda: database, executor_id="pane-original")
    record, created = opener.register_office_execution(
        "office-execution-race-01",
        "fingerprint-race",
        plan=PLAN,
        authorization={"authorizedSheetNames": ["Sheet1"]},
        workbook_identity="url:c:/work/report.xlsx",
    )
    assert created is True
    assert record.owner_token

    def claim(index: int):
        # 每个线程一个独立账本实例（= 一个独立面板进程），都用**同一个** revision。
        pane = FolderOperationLedger(lambda: database, executor_id=f"pane-{index}")
        return pane.claim_office_resume(
            "office-execution-race-01", expected_revision=record.revision
        )

    winners, losers = _race(claim, 8)
    assert len(winners) == 1, f"并发认领出现了 {len(winners)} 个执行权"
    assert len(losers) == 7
    for error in losers:
        # 落败者必须是**被拒绝**，不是"也拿到一份记录"。
        assert isinstance(error, RuntimeError)

    winner = winners[0]
    assert winner.owner_token != record.owner_token
    assert winner.revision == record.revision + 1

    # ② 被替代的执行者没有任何后续写入路径：旧 token 开动作直接被拒。
    with pytest.raises(RuntimeError, match="已停止写入"):
        opener.begin_office_action(
            record.operation_id,
            expected_revision=winner.revision,
            owner_token=record.owner_token,
            action_index=0,
            before_snapshot=BEFORE_SNAPSHOT,
        )
    still_clean = opener.office_execution(record.operation_id)
    assert still_clean is not None
    assert still_clean.in_flight_action_index is None
    assert still_clean.next_action_index == 0

    # ③ 赢家能且只能开一次。
    started = opener.begin_office_action(
        record.operation_id,
        expected_revision=winner.revision,
        owner_token=winner.owner_token,
        action_index=0,
        before_snapshot=BEFORE_SNAPSHOT,
    )
    assert started.in_flight_action_index == 0
    with pytest.raises(RuntimeError, match="已停止写入"):
        opener.begin_office_action(
            record.operation_id,
            expected_revision=started.revision,
            owner_token=winner.owner_token,
            action_index=0,
            before_snapshot=BEFORE_SNAPSHOT,
        )


def test_only_one_concurrent_planning_continuation_wins(tmp_path: Path) -> None:
    database = tmp_path / "tasks" / "state.sqlite3"
    first = FolderOperationLedger(lambda: database, executor_id="planning-original")
    record, _created = first.register_planning_task(
        "turn-claim-race",
        "fingerprint-planning-race",
        {"turnId": "turn-claim-race", "prompt": "按最新数据统计"},
        approved_wall_time_seconds=60,
    )
    paused = first.pause_planning_task(
        record.turn_id,
        {
            "schemaVersion": 1,
            "kind": "planning",
            "loopState": {"activeElapsedSeconds": 60},
        },
        active_elapsed_seconds=60,
        reason="等待用户继续",
    )
    assert paused.status == "waiting_user"

    def claim(index: int):
        # 每次点击都有自己的 request_id：幂等只对**同一** request_id 成立，
        # 所以这里测的是"两个不同的继续请求同时到达"。
        pane = FolderOperationLedger(lambda: database, executor_id=f"planner-{index}")
        return pane.claim_planning_continuation(
            paused.turn_id,
            expected_revision=paused.revision,
            request_id=f"extension-race-{index}",
            extension_seconds=60,
            maximum_extensions=4,
            hard_wall_time_seconds=300,
        )

    winners, losers = _race(claim, 8)
    assert len(winners) == 1, f"并发继续出现了 {len(winners)} 个执行权"
    assert len(losers) == 7
    for error in losers:
        assert isinstance(error, RuntimeError)

    claimed, extended = winners[0]
    assert extended is True
    assert claimed.status == "running"
    assert claimed.revision == paused.revision + 1
    assert claimed.extension_count == 1

    # 预算只被追加一次：八路并发里七路被拒，账本里只能有一条扩展记录。
    with sqlite3.connect(database) as connection:
        rows = connection.execute(
            "SELECT request_id FROM planning_budget_extensions WHERE turn_id = ?",
            (paused.turn_id,),
        ).fetchall()
    assert len(rows) == 1

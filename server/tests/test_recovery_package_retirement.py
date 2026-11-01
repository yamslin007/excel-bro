"""删除前的恢复材料留痕（U5，2026-10-30）：**删的是哪一份**必须能回答。

缺口（`docs/A1_GATE_REMAINING_2026-10-28.md` §1 U5）：`retireWpsRecoveryPackage` 直接删，
删前不存证。于是材料消失之后，"当时那份材料是不是这一份"只能靠推断 —— 而终态回收
（原地撤销成功）恰恰是**唯一**合法删除点，验收里最想知道的正是这件事。

补的是一条**顺序严格**的留痕：`DELETE /api/execution-recovery/{id}` 先算磁盘上那堆字节的
sha256、写进账本、**再**删。三列 `recovery_package_digest` / `recovery_package_bytes` /
`recovery_package_retired_at` 同一次 UPDATE 一起写、首次写赢。

本文件钉七件事：

1. 留痕里的指纹就是**删除那一刻**文件的 sha256（不是 header 里那个"内容摘要"）；
2. **先落账再删**：账本写不进去就**不许删**（503 + 材料仍在），绝不出现"材料没了还没证据"；
3. 重复 DELETE 幂等，且**不改写**第一次的指纹与时刻；
4. 材料本来就不在 ⇒ 204、**不留痕**（`NULL` 是"没有留痕"，不是"材料没被回收过"）；
5. 没登记过的执行 ⇒ 409 拒绝，且不会因此**创建**一条账本记录；
6. 留痕里**没有单元格值**（恢复包含原值，账本里一个都不许有）；
7. 旧库（版本号已是当前、但没有这三列）必须被补列。

口径与 U4 对照（两条**不许互相照抄**）：

| | 处置留痕（U4） | 材料回收留痕（U5） |
| --- | --- | --- |
| 写在哪一步 | 宿主恢复**成功之后** | 删除**之前** |
| 为什么 | 没恢复就不许说"撤销了" | 删完就无从取证 |
| 失败怎么办 | 吞掉（best-effort，不影响用户结论） | **不删**（503，材料留着到期清理） |
"""

from __future__ import annotations

import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
from server.app import execution_recovery
from server.app import task_ledger as task_ledger_module
from server.app.main import app
from server.app.models import recovery_package_plan_fingerprint
from server.app.task_ledger import FolderOperationLedger
from server.app.turn_state import turn_registry

OPERATION = "wps-retirement-0001"
IDENTITY = "wps-path:c:/work/report.xlsx"


@pytest.fixture()
def ledger(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[FolderOperationLedger, Path]:
    """给这一组用例一条**自己知道路径**的账本：要直接查库，不能只靠对象。

    `conftest.isolated_a1_ledger` 保证本文件不会写到用户真实账本；这里再显式指名，
    是为了能 `sqlite3.connect(database)` 把整行 dump 出来（"账本里没有单元格值"这条要用）。
    """

    turn_registry.reset()
    main_module._folder_operations.clear()
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    database = tmp_path / "state.sqlite3"
    instance = FolderOperationLedger(lambda: database)
    monkeypatch.setattr(main_module, "_folder_operation_ledger", instance)
    return instance, database


def _descriptor(actions: int = 3) -> dict:
    """与 `test_execution_recovery_package.py` 同形状的恢复描述符。

    两份必须保持一致：那边钉的是"恢复包能存能读"，这边钉的是"删它之前留了什么"。
    任何一边改协议，另一边都会同时红 —— 这是刻意的重复，不是可以合并的重复。
    """

    return {
        "actions": [
            {
                "actionIndex": index,
                "type": "writeValues",
                "sheet": "目标",
                "range": f"A{index + 1}:A{index + 1}",
                "rowCount": 1,
                "columnCount": 1,
                "cellCount": 1,
                "previewFingerprint": f"v2:1:{index:08x}",
                "expectedFingerprint": f"v2:1:{(index + 10):08x}",
            }
            for index in range(actions)
        ],
        "resumeContext": {
            "version": 1,
            "hostVersion": "12.0",
            "sources": [
                {
                    "sheetName": "来源",
                    "range": "A1:A1",
                    "rowCount": 1,
                    "columnCount": 1,
                    "previewFingerprint": "v2:1:eeeeeeee",
                    "valueFingerprint": "v2:1:ffffffff",
                }
            ],
        },
    }


def _package(operation_id: str = OPERATION) -> dict:
    descriptor = _descriptor()
    rows = [
        {
            **action,
            "values": [[index + 1]],
            "beforeValues": [[f"old{index + 1}"]],
            "beforeFormulas": [[f"old{index + 1}"]],
        }
        for index, action in enumerate(descriptor["actions"])
    ]
    return {
        "schemaVersion": 1,
        "host": "wps",
        "operationId": operation_id,
        "taskId": "task-keep",
        "workbookIdentity": IDENTITY,
        "authorizedSheetNames": ["来源", "目标"],
        "planId": "plan-resume",
        "planTitle": "更新",
        "planFingerprint": recovery_package_plan_fingerprint(rows),
        "hostVersion": "12.0",
        "createdAt": "2026-10-20T00:00:00.000Z",
        "actions": rows,
        "sourceEvidence": descriptor["resumeContext"],
    }


def _register(client: TestClient, operation_id: str = OPERATION) -> None:
    response = client.post(
        "/api/wps/executions/register",
        json={
            "operationId": operation_id,
            "taskId": "task-keep",
            "confirmed": True,
            "workbookIdentity": IDENTITY,
            "authorizedSheetNames": ["来源", "目标"],
            "descriptor": _descriptor(),
        },
    )
    assert response.status_code == 200, response.text


def _put(client: TestClient, operation_id: str = OPERATION) -> None:
    text = json.dumps(_package(operation_id), ensure_ascii=False)
    response = client.put(
        f"/api/execution-recovery/{operation_id}",
        json={"packageJson": text, "size": len(text.encode("utf-8"))},
    )
    assert response.status_code == 200, response.text


def test_delete_records_the_digest_of_the_bytes_it_deleted(ledger, monkeypatch) -> None:
    """留痕的指纹 = **删除那一刻**磁盘上那堆字节的 sha256。"""

    instance, _database = ledger
    client = TestClient(app)
    _register(client)
    _put(client)

    before = execution_recovery.raw_digest(OPERATION)
    assert before is not None
    digest, size = before
    assert len(digest) == 64 and size > 0

    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204

    # 材料真的没了 —— 否则下面那条"留痕"证不到"删前"。
    assert execution_recovery.raw_digest(OPERATION) is None
    assert execution_recovery.status(OPERATION) == "missing"

    retired = instance.recovery_package_retirement(OPERATION)
    assert retired is not None
    assert retired["digest"] == digest
    assert retired["bytes"] == size
    # 时刻是可比的 epoch 秒，而且不是"未来"。
    assert retired["retiredAt"] <= datetime.now(timezone.utc).timestamp() + 1


def test_repeat_delete_keeps_the_first_trail(ledger) -> None:
    """删第二次（材料已不在）是幂等成功，**不改写**第一次的指纹与两个时刻。"""

    instance, _database = ledger
    client = TestClient(app)
    _register(client)
    _put(client)
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204
    first = instance.recovery_package_retirement(OPERATION)
    assert first is not None
    assert first["removedAt"] is not None

    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204
    second = instance.recovery_package_retirement(OPERATION)
    assert second == first


def test_first_digest_and_first_removal_survive_a_second_retirement(ledger) -> None:
    """重新落一份材料再删一次：**第一次的指纹与删除时刻都保留**（首次写赢）。

    为什么这份材料第二次的字节必然不同：文件头带 `savedAt`，重新保存一定变。
    所以这条同时是"首次写赢"的正向锚点 —— 若判据写成"最后一次"，`digest` 就会换成新的。
    """

    instance, _database = ledger
    client = TestClient(app)
    _register(client)
    _put(client)
    first_digest, _size = execution_recovery.raw_digest(OPERATION)
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204
    first = instance.recovery_package_retirement(OPERATION)
    assert first is not None

    # 同一执行重新落一份材料（正文相同、文件头时间不同 ⇒ 字节不同 ⇒ 摘要求必然不同）。
    _put(client)
    second_digest, _size2 = execution_recovery.raw_digest(OPERATION)
    assert second_digest != first_digest, "重新落盘后字节必须变了，否则这条用例证不到'首次保留'"
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204

    assert instance.recovery_package_retirement(OPERATION) == first


def test_registered_execution_without_material_leaves_no_trail(ledger) -> None:
    """登记过但**从没落过材料**：DELETE 仍然 204，且不留痕（不能瞎记一个指纹）。"""

    instance, _database = ledger
    client = TestClient(app)
    _register(client)
    assert execution_recovery.raw_digest(OPERATION) is None

    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204

    assert instance.office_execution(OPERATION) is not None
    assert instance.recovery_package_retirement(OPERATION) is None


def test_unknown_execution_is_refused_and_creates_no_record(ledger) -> None:
    """没登记过的执行：409 拒绝，且**不补建**账本记录（留痕不能造事实）。"""

    instance, _database = ledger
    client = TestClient(app)

    assert client.delete("/api/execution-recovery/wps-retirement-9999").status_code == 409

    assert instance.office_execution("wps-retirement-9999") is None
    assert instance.recovery_package_retirement("wps-retirement-9999") is None


def test_cross_namespace_id_is_refused_and_leaves_no_trail(ledger) -> None:
    """身份不符：Office 身份（`url:`）的记录在 WPS 的回收端点上就是"不存在"。

    这条钉的是"先核身份再动手"：跨命名空间既不能删材料，也不能留痕 ——
    否则一次误传 id 就会在别人的记录上写一条"材料已交出"。
    """

    instance, _database = ledger
    client = TestClient(app)
    instance.register_office_execution(
        "office-retirement-0001",
        "fingerprint-office",
        plan={"id": "office-plan", "title": "办公簿", "summary": "", "assumptions": [],
              "warnings": [], "actions": []},
        authorization={"authorizedSheetNames": ["汇总"], "allowedExternalFiles": []},
        workbook_identity="url:c:/work/office.xlsx",
    )

    assert client.delete("/api/execution-recovery/office-retirement-0001").status_code == 409

    assert instance.recovery_package_retirement("office-retirement-0001") is None
    # 记录本身没被删、也没被改。
    assert instance.office_execution("office-retirement-0001") is not None


def test_ledger_failure_refuses_to_delete_and_keeps_the_material(ledger, monkeypatch) -> None:
    """**先落账再删**：账本写不进去就不许删，材料必须还在（fail-closed）。

    这条是 U5 与 U4 的分界：处置留痕失败只是"少一条观测"；材料回收留痕失败意味着
    "要么没证据、要么别删"。这里选后者 —— 宁可材料留着到期清理，也不留下无证删除。
    """

    instance, _database = ledger
    client = TestClient(app)
    _register(client)
    _put(client)

    def boom(self, operation_id, digest, size):  # noqa: ANN001 - 替身签名跟随真实方法
        raise sqlite3.OperationalError("database is locked")

    original = FolderOperationLedger.record_recovery_package_retirement
    monkeypatch.setattr(FolderOperationLedger, "record_recovery_package_retirement", boom)

    response = client.delete(f"/api/execution-recovery/{OPERATION}")
    assert response.status_code == 503
    assert response.json()["detail"]["code"] == "RECOVERY_PACKAGE_RETIREMENT_UNRECORDED"

    # 材料一个字节都没少：读回还是 200，指纹还是删除前的那个。
    # **只还原这一条**（`monkeypatch.undo()` 会连 `EXCEL_BRO_CONFIG_DIR` 一起退回去，
    # 于是"材料还在"会被误判成"材料不见"—— 那正是这条用例要防的假绿）。
    monkeypatch.setattr(
        FolderOperationLedger, "record_recovery_package_retirement", original
    )
    assert execution_recovery.raw_digest(OPERATION) is not None
    assert client.get(f"/api/execution-recovery/{OPERATION}").status_code == 200
    assert instance.recovery_package_retirement(OPERATION) is None


def test_removal_failure_records_the_digest_but_no_deletion_confirmation(ledger, monkeypatch) -> None:
    """**落账之后删除失败**：账本必须分得开"记了删前摘要"与"文件确已删除"。

    这是这条留痕最容易读错的一格：`retired_at` 有值**不等于**删掉了。所以删除失败时
    `removed_at` 必须为空，而且材料确实还在（不是"删了但没记"）。
    """

    instance, _database = ledger
    client = TestClient(app)
    _register(client)
    _put(client)
    digest_before, _size = execution_recovery.raw_digest(OPERATION)

    def boom(operation_id: str) -> bool:
        raise execution_recovery.RecoveryPackagePersistenceError("恢复包无法从本机数据目录删除。")

    original_remove = execution_recovery.remove
    monkeypatch.setattr(execution_recovery, "remove", boom)

    response = client.delete(f"/api/execution-recovery/{OPERATION}")
    assert response.status_code == 507
    assert response.json()["detail"]["code"] == "RECOVERY_PACKAGE_PERSISTENCE"

    trail = instance.recovery_package_retirement(OPERATION)
    assert trail is not None
    assert trail["digest"] == digest_before
    assert trail["retiredAt"] is not None
    # ★ 关键：没有删除确认。读这一行的人不许据此宣称"材料已删"。
    assert trail["removedAt"] is None

    monkeypatch.setattr(execution_recovery, "remove", original_remove)
    assert execution_recovery.raw_digest(OPERATION) == (digest_before, _size)
    assert client.get(f"/api/execution-recovery/{OPERATION}").status_code == 200


def test_retry_after_removal_failure_confirms_deletion_and_keeps_the_first_trail(ledger, monkeypatch) -> None:
    """删除失败后重试：只重删同一份材料，**不改写**第一次的指纹与"交出回收"时刻。

    重试是幂等的：`retired_at` 与指纹都保留第一次的值，只补上 `removed_at`。
    这条同时钉住"重试不会变成第二次留痕"——否则账本会看起来像删了两份材料。
    """

    instance, _database = ledger
    client = TestClient(app)
    _register(client)
    _put(client)
    digest_before, _size = execution_recovery.raw_digest(OPERATION)

    def boom(operation_id: str) -> bool:
        raise execution_recovery.RecoveryPackagePersistenceError("恢复包无法从本机数据目录删除。")

    original_remove = execution_recovery.remove
    monkeypatch.setattr(execution_recovery, "remove", boom)
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 507
    failed = instance.recovery_package_retirement(OPERATION)
    assert failed is not None and failed["removedAt"] is None

    # 重试：只把"删除"这一步再来一次。
    monkeypatch.setattr(execution_recovery, "remove", original_remove)
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204

    retried = instance.recovery_package_retirement(OPERATION)
    assert retried is not None
    assert retried["digest"] == failed["digest"] == digest_before
    assert retried["retiredAt"] == failed["retiredAt"]  # 第一次的时刻，不是重试的时间
    assert retried["removedAt"] is not None  # 这一次才确认删掉了
    assert execution_recovery.raw_digest(OPERATION) is None


def test_the_trail_carries_no_cell_values(ledger) -> None:
    """账本里**没有单元格值**：留痕只有指纹/字节数/时刻，恢复包正文一个字符都不进库。"""

    instance, database = ledger
    client = TestClient(app)
    _register(client)
    _put(client)
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204

    with sqlite3.connect(database) as connection:
        connection.row_factory = sqlite3.Row
        row = connection.execute(
            "SELECT * FROM office_executions WHERE operation_id = ?", (OPERATION,)
        ).fetchone()
        writes = connection.execute(
            "SELECT * FROM office_host_writes WHERE operation_id = ?", (OPERATION,)
        ).fetchall()
    assert row is not None
    dump = json.dumps({key: row[key] for key in row.keys()}, ensure_ascii=False, default=str)
    dump += json.dumps([dict(entry) for entry in writes], ensure_ascii=False, default=str)
    # 夹具里的写前值是 old1/old2/old3，写后值是 1/2/3 —— 一个都不许出现在账本里。
    for token in ("old1", "old2", "old3", "beforeValues"):
        assert token not in dump
    # 正向锚点：留痕本身确实在（否则"没有值"可以因为整件事没发生而假绿）。
    assert instance.recovery_package_retirement(OPERATION) is not None


def test_legacy_ledger_is_migrated_with_the_three_retirement_columns(tmp_path: Path) -> None:
    """旧库（没有这三列）必须被补列 —— `_REQUIRED_OFFICE_COLUMNS` 少一列就永久跳过。

    刻意把 `user_version` 调到**当前**版本：版本落后时 bootstrap 反正会重跑，列会被补回来，
    判据就等于没在测那份"必需列"名单。真实故障（2026-10-27）正是"版本号看着是新的、
    列却没补"——那时唯一能救的就是逐列比对。
    """

    database = tmp_path / "legacy.sqlite3"
    first = FolderOperationLedger(lambda: database)
    first.register_office_execution(
        "wps-retirement-legacy",
        "fingerprint-legacy",
        plan={"id": "legacy-plan", "title": "旧库", "summary": "", "assumptions": [],
              "warnings": [], "actions": []},
        authorization={"authorizedSheetNames": ["目标"], "allowedExternalFiles": []},
        workbook_identity=IDENTITY,
        identity_prefix="wps-path:",
    )
    with sqlite3.connect(database) as connection:
        connection.execute("ALTER TABLE office_executions DROP COLUMN recovery_package_digest")
        connection.execute("ALTER TABLE office_executions DROP COLUMN recovery_package_bytes")
        connection.execute("ALTER TABLE office_executions DROP COLUMN recovery_package_retired_at")
        connection.execute(
            f"PRAGMA user_version = {task_ledger_module._SCHEMA_VERSION}"
        )

    restarted = FolderOperationLedger(lambda: database)
    assert restarted.recovery_package_retirement("wps-retirement-legacy") is None
    exists, recorded = restarted.record_recovery_package_retirement(
        "wps-retirement-legacy", "a" * 64, 1234
    )
    assert exists is True
    assert recorded is not None and recorded["bytes"] == 1234
    assert restarted.recovery_package_retirement("wps-retirement-legacy") == recorded


def test_ledger_refuses_a_bad_digest_and_reports_missing_record(ledger) -> None:
    """指纹不是 64 位十六进制就报错（不猜、不截断）；记录不存在返回 `(False, None)`。"""

    instance, _database = ledger
    with pytest.raises(ValueError):
        instance.record_recovery_package_retirement(OPERATION, "not-a-digest", 10)
    with pytest.raises(ValueError):
        instance.record_recovery_package_retirement(OPERATION, "a" * 64, -1)
    assert instance.record_recovery_package_retirement("wps-retirement-absent", "b" * 64, 1) == (
        False,
        None,
    )

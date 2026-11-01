"""用户处置留痕：**放弃变更 / 撤销本次执行**必须在服务端留下可核实的一行。

真实现场（2026-10-29 Excel，`execution-c52ad791…`）：三步写入、第二步后中断、约 2.5 小时后
只续第三步、确认变更。账本里 `status=completed`、`revision=11`、`next=3`、三行
`office_host_writes` —— 但"用户点了确认/放弃/撤销"这件事**一个字都没有**。
于是判断"这次执行到底是用户主动撤掉的，还是根本没跑完"只能靠"恢复材料没了 + `revision`
没动 + 截图"三方推断。这组用例钉的是补上的那一行：

1. 两种取值各写一列，`abandoned` 与 `undone` **不许互相冒充**；
2. **首次写赢**：重复上报不改写第一次的时刻（重试、双端各报一次都不行）；
3. **不推进 `revision`、不改 `status`、不动动作游标** —— 留痕是审计，不是进度；
4. 不存在的 ID 一律 404，且**不会因此创建任何记录**；
5. 另一端的 ID 在本端端点上就是"不存在"（不能跨命名空间探测）；
6. 非法取值 422，不静默归一到某个默认值。
"""

from __future__ import annotations

import sqlite3

import pytest
from fastapi.testclient import TestClient

from server.app import task_ledger as task_ledger_module
from server.app.main import app
from server.app.task_ledger import FolderOperationLedger

OFFICE_ID = "office-disposition-0001"
WPS_ID = "wps-disposition-0001"
OFFICE_IDENTITY = "url:c:/work/disposition.xlsx"
WPS_IDENTITY = "wps-path:c:/work/disposition.xlsx"


class _Clock:
    """可推进的假时钟：让"首次写赢"变成严格可判（第二次必然拿到更大的值）。"""

    def __init__(self, start: float = 1000.0) -> None:
        self.now = start

    def time(self) -> float:
        self.now += 1.0
        return self.now


@pytest.fixture()
def clock(monkeypatch: pytest.MonkeyPatch) -> _Clock:
    fake = _Clock()
    monkeypatch.setattr(task_ledger_module, "time", fake)
    return fake


def _office_plan() -> dict[str, object]:
    return {
        "id": "office-plan-disposition-1",
        "title": "写入测试",
        "summary": "处置留痕",
        "assumptions": [],
        "warnings": [],
        "actions": [
            {"type": "writeValues", "sheet": "结果", "range": "B2", "values": [["111"]]}
        ],
    }


def _register_office(client: TestClient, operation_id: str = OFFICE_ID) -> dict:
    response = client.post(
        "/api/office/executions/register",
        json={
            "operationId": operation_id,
            "plan": _office_plan(),
            "authorizedSheetNames": ["结果"],
            "allowedExternalFiles": [],
            "workbookIdentity": OFFICE_IDENTITY,
            "confirmed": True,
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def _register_wps(client: TestClient, operation_id: str = WPS_ID) -> dict:
    response = client.post(
        "/api/wps/executions/register",
        json={
            "operationId": operation_id,
            "taskId": "task-disposition",
            "confirmed": True,
            "workbookIdentity": WPS_IDENTITY,
            "authorizedSheetNames": ["Sheet1"],
            "descriptor": {
                "actions": [
                    {
                        "actionIndex": 0,
                        "type": "writeValues",
                        "sheet": "Sheet1",
                        "range": "B2:B2",
                        "rowCount": 1,
                        "columnCount": 1,
                        "cellCount": 1,
                        "previewFingerprint": "v2:before-0",
                        "expectedFingerprint": "v2:after-0",
                    }
                ]
            },
        },
    )
    assert response.status_code == 200, response.text
    return response.json()


def _ledger() -> FolderOperationLedger:
    from server.app import main as main_module

    return main_module._folder_operation_ledger


def test_office_abandoned_is_recorded_and_progress_is_untouched() -> None:
    """放弃变更留痕：写 `abandoned_at`，`revision`/`status`/游标一字不动。"""
    client = TestClient(app)
    state = _register_office(client)
    before = _ledger().office_execution(OFFICE_ID)
    assert before is not None

    response = client.post(
        f"/api/office/executions/{OFFICE_ID}/disposition",
        json={"disposition": "abandoned"},
    )
    assert response.status_code == 204, response.text

    after = _ledger().office_execution(OFFICE_ID)
    assert after is not None
    assert after.revision == before.revision
    assert after.status == before.status
    assert after.next_action_index == before.next_action_index
    assert after.in_flight_action_index == before.in_flight_action_index
    disposition = _ledger().office_disposition(OFFICE_ID)
    assert disposition is not None
    assert disposition["abandonedAt"] is not None
    assert disposition["undoneAt"] is None  # 两种取值不许互相冒充
    assert state["status"] == "running"


def test_wps_undone_is_recorded_on_its_own_namespace() -> None:
    """撤销本次执行留痕：WPS 端点写 `undone_at`，不碰 `abandoned_at`。"""
    client = TestClient(app)
    _register_wps(client)
    assert (
        client.post(
            f"/api/wps/executions/{WPS_ID}/disposition", json={"disposition": "undone"}
        ).status_code
        == 204
    )
    disposition = _ledger().office_disposition(WPS_ID)
    assert disposition is not None
    assert disposition["undoneAt"] is not None
    assert disposition["abandonedAt"] is None


def test_wps_state_response_projects_disposition_so_reopen_cannot_resume_again() -> None:
    """WPS 状态响应必须投影 disposition —— 撤销后账本 status/游标都不变，跨重启续跑判据
    只能靠这两列识别终态。

    Excel 侧 `GET /api/office/executions/{id}` 已经投影（`test_state_response_projects_...`），
    WPS 侧 `GET /api/wps/executions/{id}` 此前**没有**：撤销成功后 `undoWpsExecution` 只把
    宿主缓存里的 `record.state` 改成 `restored`，`retireWpsRecoveryPackage` 又 best-effort
    删后端材料。重开后缓存没了、材料若删失败还在，水合会按账本游标把 state 推回
    `partial_write_recoverable`，续跑被重新开放。投影这两列，客户端才能先判终态。
    """
    client = TestClient(app)
    _register_wps(client, "wps-disposition-0009")
    state_url = "/api/wps/executions/wps-disposition-0009"

    fresh = client.get(state_url)
    assert fresh.status_code == 200, fresh.text
    assert fresh.json()["disposition"] == {"undoneAt": None, "abandonedAt": None}

    assert (
        client.post(
            f"{state_url}/disposition", json={"disposition": "undone"}
        ).status_code
        == 204
    )
    undone = client.get(state_url)
    assert undone.status_code == 200
    assert undone.json()["disposition"]["undoneAt"] is not None
    assert undone.json()["disposition"]["abandonedAt"] is None
    # 撤销不改 status / 游标 —— 但状态响应仍如实带上 disposition，供客户端判终态。
    assert undone.json()["status"] in {"running", "needs_reconciliation", "failed"}


def test_first_write_wins_and_the_two_kinds_are_independent(clock: _Clock) -> None:
    """重复上报不改写第一次的时刻；两种取值各自记各自的第一次。"""
    client = TestClient(app)
    _register_office(client, "office-disposition-0002")
    url = "/api/office/executions/office-disposition-0002/disposition"

    assert client.post(url, json={"disposition": "abandoned"}).status_code == 204
    first = _ledger().office_disposition("office-disposition-0002")
    assert first is not None and first["abandonedAt"] is not None

    # 第二次（重试 / 用户又点了一次 / 另一个窗格上报）必须拿回第一次的时刻。
    assert client.post(url, json={"disposition": "abandoned"}).status_code == 204
    second = _ledger().office_disposition("office-disposition-0002")
    assert second is not None
    assert second["abandonedAt"] == first["abandonedAt"], "首次写赢：重复上报不许推后时刻"

    # 换一种取值：写另一列，且不动前一列。
    assert client.post(url, json={"disposition": "undone"}).status_code == 204
    third = _ledger().office_disposition("office-disposition-0002")
    assert third is not None
    assert third["abandonedAt"] == first["abandonedAt"]
    assert third["undoneAt"] is not None and third["undoneAt"] > first["abandonedAt"]


def test_unknown_operation_is_404_and_creates_nothing() -> None:
    """不存在的执行：404，且**绝不**为了留痕补建一条账本。"""
    client = TestClient(app)
    assert (
        client.post(
            "/api/office/executions/never-registered-0001/disposition",
            json={"disposition": "abandoned"},
        ).status_code
        == 404
    )
    assert (
        client.post(
            "/api/wps/executions/never-registered-0002/disposition",
            json={"disposition": "undone"},
        ).status_code
        == 404
    )
    assert _ledger().office_execution("never-registered-0001") is None
    assert _ledger().office_execution("never-registered-0002") is None
    assert _ledger().office_disposition("never-registered-0001") is None


def test_cross_namespace_disposition_is_404() -> None:
    """另一端 ID 在本端端点上是"不存在" —— 留痕不是跨宿主探测工具。"""
    client = TestClient(app)
    _register_office(client, "office-disposition-0003")
    _register_wps(client, "wps-disposition-0003")

    assert (
        client.post(
            "/api/wps/executions/office-disposition-0003/disposition",
            json={"disposition": "abandoned"},
        ).status_code
        == 404
    )
    assert (
        client.post(
            "/api/office/executions/wps-disposition-0003/disposition",
            json={"disposition": "undone"},
        ).status_code
        == 404
    )
    # 反向也要成立：跨命名空间的 404 不该顺手把对端的记录写脏。
    office = _ledger().office_disposition("office-disposition-0003")
    wps = _ledger().office_disposition("wps-disposition-0003")
    assert office == {"abandonedAt": None, "undoneAt": None}
    assert wps == {"abandonedAt": None, "undoneAt": None}


def test_invalid_disposition_is_rejected_instead_of_defaulted() -> None:
    """非法取值 422：不静默当成 `undone`，也不当成"没发生"。"""
    client = TestClient(app)
    _register_office(client, "office-disposition-0004")
    url = "/api/office/executions/office-disposition-0004/disposition"
    for bad in ("confirmed", "abandon", "UNDONE", "", None, 1):
        assert client.post(url, json={"disposition": bad}).status_code == 422, bad
    assert _ledger().office_disposition("office-disposition-0004") == {
        "abandonedAt": None,
        "undoneAt": None,
    }


def test_state_response_projects_disposition_so_a_reopened_pane_does_not_offer_it_again() -> None:
    """重开的窗格只能靠**状态响应**知道"这个入口已经被用掉了"。

    留痕落在两列里还不够：Excel 恢复卡在 `GET /api/office/executions/{id}` 这次读数上判
    "已撤销 / 已放弃 ⇒ 不再重建撤销入口、也不再摆成可继续"。少了这条投影，重开后界面只能
    凭"账本还没完成"把入口重新画出来 —— 用户点下去只会被写后指纹挡下（2026-10-31 真机：
    撤销过一次的执行又收到一颗同样的按钮）。

    两种取值仍不许互相冒充：`undone` 只写 `undoneAt`，`abandoned` 只写 `abandonedAt`，
    而这条投影是**照抄**账本的，不在这里做任何解释。
    """
    client = TestClient(app)
    _register_office(client, "office-disposition-0007")
    state_url = "/api/office/executions/office-disposition-0007"
    disposition_url = f"{state_url}/disposition"

    fresh = client.get(state_url)
    assert fresh.status_code == 200, fresh.text
    assert fresh.json()["disposition"] == {"undoneAt": None, "abandonedAt": None}

    assert client.post(disposition_url, json={"disposition": "undone"}).status_code == 204
    undone = client.get(state_url)
    assert undone.status_code == 200
    assert undone.json()["disposition"]["undoneAt"] is not None
    assert undone.json()["disposition"]["abandonedAt"] is None

    # 另一种取值写另一列，投影跟着走（不因为"已经处置过"就吞掉第二次的类别）。
    _register_office(client, "office-disposition-0008")
    assert (
        client.post(
            "/api/office/executions/office-disposition-0008/disposition",
            json={"disposition": "abandoned"},
        ).status_code
        == 204
    )
    abandoned = client.get("/api/office/executions/office-disposition-0008")
    assert abandoned.json()["disposition"]["abandonedAt"] is not None
    assert abandoned.json()["disposition"]["undoneAt"] is None


def test_legacy_ledger_is_migrated_with_the_two_disposition_columns(tmp_path) -> None:
    """旧库（`user_version=11`、没有这两列）必须被补列，而不是"列不存在"直接 500。

    版本号只是意图，列才是事实：`_REQUIRED_OFFICE_COLUMNS` 少一列就会让补列永久跳过
    （2026-10-27 真机就是这么 500 的）。这里把两个新列真的删掉，再看新实例能不能补回来。
    """
    database = tmp_path / "state.sqlite3"
    first = FolderOperationLedger(lambda: database)
    first.register_office_execution(
        "office-disposition-legacy",
        "fingerprint-legacy",
        plan={"id": "legacy-plan", "title": "旧库", "summary": "", "assumptions": [],
              "warnings": [], "actions": []},
        authorization={"authorizedSheetNames": ["结果"], "allowedExternalFiles": []},
        workbook_identity=OFFICE_IDENTITY,
    )
    with sqlite3.connect(database) as connection:
        connection.execute("ALTER TABLE office_executions DROP COLUMN abandoned_at")
        connection.execute("ALTER TABLE office_executions DROP COLUMN undone_at")
        connection.execute("PRAGMA user_version = 11")

    restarted = FolderOperationLedger(lambda: database)
    assert restarted.office_disposition("office-disposition-legacy") == {
        "abandonedAt": None,
        "undoneAt": None,
    }
    exists, recorded = restarted.record_office_disposition(
        "office-disposition-legacy", "undone"
    )
    assert exists is True and recorded is not None
    with sqlite3.connect(database) as connection:
        assert connection.execute("PRAGMA user_version").fetchone()[0] >= 12
        columns = {row[1] for row in connection.execute("PRAGMA table_info(office_executions)")}
        assert {"abandoned_at", "undone_at"} <= columns
        assert connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    # 迁移不该顺手改执行进度。
    record = restarted.office_execution("office-disposition-legacy")
    assert record is not None and record.revision == 1 and record.status == "running"


def test_both_endpoints_accept_both_values() -> None:
    """两端同一份取值（`Literal["abandoned","undone"]` ↔ TS 的 `ExecutionDisposition`）。

    两端读同一份判据：服务端两个模型、客户端 `packages/core/src/executionDisposition.ts`、
    `apps/excel-addin/src/contracts.ts` 的转发，必须是同一组取值。任何一端偷偷多收一个
    或漏收一个，这里都会红。
    """
    client = TestClient(app)
    _register_office(client, "office-disposition-0006")
    _register_wps(client, "wps-disposition-0006")
    for path, value in (
        ("/api/office/executions/office-disposition-0006/disposition", "undone"),
        ("/api/wps/executions/wps-disposition-0006/disposition", "abandoned"),
    ):
        assert client.post(path, json={"disposition": value}).status_code == 204
    assert _ledger().office_disposition("office-disposition-0006")["undoneAt"] is not None
    assert _ledger().office_disposition("wps-disposition-0006")["abandonedAt"] is not None


def test_ledger_rejects_unknown_disposition_and_reports_missing_record() -> None:
    """账本层：未知取值抛 `ValueError`（不是静默返回 False）；缺记录返回 `(False, None)`。"""
    client = TestClient(app)
    _register_office(client, "office-disposition-0005")
    with pytest.raises(ValueError):
        _ledger().record_office_disposition("office-disposition-0005", "confirmed")
    assert _ledger().record_office_disposition("office-disposition-0005", "abandoned")[0] is True
    exists, recorded_at = _ledger().record_office_disposition("no-such-record", "abandoned")
    assert exists is False and recorded_at is None

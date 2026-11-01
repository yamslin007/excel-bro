"""未完成任务恢复包：本机数据目录的原子存储、期限、配额与 HTTP 契约。

缺陷来源（2026-10-27 真实 WPS 验收）：完整退出 WPS 后 `PluginStorage` 读不到原执行
记录，于是"已核对第二步、仍可继续第三步"在真实机器上断链。修复方向是**本机后端
应用数据目录**持久保存不可变恢复材料，本文件钉住它的边界：

- 落盘后能跨"后端重启"（新 TestClient + 新存储实例）读回**逐字节相同**的正文；
- 摘要不符、过期、配额不足各有独立状态，不许互相冒充；
- 配额不足时**不淘汰尚未结束任务的包**；
- 恢复包内含单元格原值，但**不进入 SQLite 账本**、不进入任何对外事件。

存储测试直接调 `execution_recovery`，接口测试走真实 `TestClient`；两者都用
`EXCEL_BRO_CONFIG_DIR` 指向 `tmp_path`，不碰用户真实账本。
"""

from __future__ import annotations

import json
import hashlib
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
from server.app import execution_recovery
from server.app.capabilities import capability_int
from server.app.execution_recovery import (
    RecoveryPackageConflict,
    RecoveryPackageCorrupt,
    RecoveryPackageExpired,
    RecoveryPackageError,
    RecoveryPackageMissing,
    RecoveryPackagePersistenceError,
    RecoveryPackageQuota,
)
from server.app.main import app
from server.app.models import ExecutionRecoveryPackage, recovery_package_plan_fingerprint
from server.app.turn_state import turn_registry

REPO_ROOT = Path(__file__).resolve().parents[2]
OPERATION = "wps-execution-0001"


@pytest.mark.parametrize("case", json.loads((REPO_ROOT / "config/execution_recovery_package_cases.json").read_text(encoding="utf-8"))["formulaActionCases"], ids=lambda case: case["name"])
def test_formula_action_cases_share_core_contract(case):
    """公式动作的目标矩阵必须按类型二选一，与 TS 侧读同一份样例。

    "计算值顶替公式"是本文件最要紧的一条否定判据：两条公式能算出同一个数，
    把值写进 `formulas` 或把公式写进 `values` 都必须被拒，否则续跑会把公式写成常量。
    计划指纹在替换动作后重算，避免用例因为指纹不符而"假通过"。
    """

    fixture = json.loads((REPO_ROOT / "config/execution_recovery_package_cases.json").read_text(encoding="utf-8"))["valid"][0]
    fixture["actions"] = [json.loads(json.dumps(case["action"]))]
    fixture["planFingerprint"] = recovery_package_plan_fingerprint(fixture["actions"])
    if case["valid"]:
        ExecutionRecoveryPackage.model_validate(fixture)
    else:
        with pytest.raises(ValueError):
            ExecutionRecoveryPackage.model_validate(fixture)


@pytest.mark.parametrize("case", json.loads((REPO_ROOT / "config/execution_recovery_package_cases.json").read_text(encoding="utf-8"))["sourceEvidenceCases"])
def test_source_evidence_versions_share_core_contract(case):
    fixture = json.loads((REPO_ROOT / "config/execution_recovery_package_cases.json").read_text(encoding="utf-8"))["valid"][0]
    fixture["sourceEvidence"]["version"] = case["version"]
    if case["empty"]:
        fixture["sourceEvidence"]["sources"] = []
    if case["valid"]:
        ExecutionRecoveryPackage.model_validate(fixture)
    else:
        with pytest.raises(ValueError):
            ExecutionRecoveryPackage.model_validate(fixture)


def test_literal_resume_descriptor_does_not_allow_non_value_actions():
    from server.app.models import WpsRecoveryDescriptor
    descriptor = _descriptor()
    descriptor["resumeContext"] = {"version": 2, "hostVersion": "12.0", "sources": []}
    WpsRecoveryDescriptor.model_validate(descriptor)
    descriptor["actions"][0]["type"] = "writeFormulas"
    with pytest.raises(ValueError):
        WpsRecoveryDescriptor.model_validate(descriptor)


@pytest.fixture(autouse=True)
def isolate(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    turn_registry.reset()
    main_module._folder_operations.clear()
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))


def _descriptor(actions: int = 3) -> dict:
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


def _package(operation_id: str = OPERATION, actions: int = 3) -> dict:
    descriptor = _descriptor(actions)
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
        "workbookIdentity": "wps-path:c:/work/report.xlsx",
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
            "workbookIdentity": "wps-path:c:/work/report.xlsx",
            "authorizedSheetNames": ["来源", "目标"],
            "descriptor": _descriptor(),
        },
    )
    assert response.status_code == 200, response.text


def _put(client: TestClient, operation_id: str = OPERATION, package: dict | None = None):
    text = json.dumps(package or _package(operation_id), ensure_ascii=False)
    return client.put(
        f"/api/execution-recovery/{operation_id}",
        json={"packageJson": text, "size": len(text.encode("utf-8"))},
    )


def test_store_roundtrip_survives_backend_restart(tmp_path: Path) -> None:
    """跨"重启"读回：文件在应用数据目录，不在 Temp，也不在工作簿旁边。"""

    package_json = json.dumps(_package(), ensure_ascii=False)
    saved = execution_recovery.save(OPERATION, package_json)
    root = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]).resolve()
    folder = (root / "execution-recovery" / OPERATION).resolve()
    assert folder.is_relative_to(root), "恢复包必须落在既有应用数据根目录下"
    assert folder.relative_to(root).as_posix() == f"execution-recovery/{OPERATION}"
    assert (folder / "material.dat").is_file()
    assert len(saved["digest"]) == 64

    # 新实例 = 新进程：没有任何内存状态可复用，只能从磁盘读。
    loaded = execution_recovery.load(OPERATION)
    assert loaded["packageJson"] == package_json
    assert loaded["digest"] == saved["digest"]
    created = datetime.fromisoformat(saved["savedAt"].replace("Z", "+00:00"))
    expires = datetime.fromisoformat(loaded["expiresAt"].replace("Z", "+00:00"))
    assert expires - created == timedelta(days=capability_int("taskRecovery", "unfinishedRecoveryPackageRetentionDays"))
    assert execution_recovery.status(OPERATION) == "ok"


def test_tampered_material_is_corrupt_not_recovered(tmp_path: Path) -> None:
    execution_recovery.save(OPERATION, json.dumps(_package(), ensure_ascii=False))
    path = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION / "material.dat"
    data = path.read_bytes()
    path.write_bytes(data.replace("old2".encode("utf-8"), b"hax2"))
    with pytest.raises(RecoveryPackageCorrupt):
        execution_recovery.load(OPERATION)
    assert execution_recovery.status(OPERATION) == "corrupt"


def test_expired_material_reports_its_own_state(tmp_path: Path) -> None:
    execution_recovery.save(OPERATION, json.dumps(_package(), ensure_ascii=False))
    later = datetime.now(timezone.utc) + timedelta(days=8)
    with pytest.raises(RecoveryPackageExpired):
        execution_recovery.load(OPERATION, now=later)
    assert execution_recovery.status(OPERATION, now=later) == "expired"


def test_missing_material_is_explicit(tmp_path: Path) -> None:
    with pytest.raises(RecoveryPackageMissing):
        execution_recovery.load(OPERATION)
    assert execution_recovery.status(OPERATION) == "missing"


def test_identical_material_is_idempotent_but_a_different_plan_is_rejected(tmp_path: Path) -> None:
    package_json = json.dumps(_package(), ensure_ascii=False)
    first = execution_recovery.save(OPERATION, package_json)
    second = execution_recovery.save(OPERATION, package_json)
    assert second["digest"] == first["digest"] and second["savedAt"] == first["savedAt"]
    other = _package()
    other["planTitle"] = "另一次执行"
    with pytest.raises(RecoveryPackageConflict):
        execution_recovery.save(OPERATION, json.dumps(other, ensure_ascii=False))


def test_quota_refuses_new_material_and_keeps_unfinished_packages(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """配额不足时**不淘汰尚未结束任务**的包，而是明确拒绝新增。"""

    original = capability_int

    def limits(section: str, key: str) -> int:
        if (section, key) == ("taskRecovery", "recoveryPackageMaxOperations"):
            return 1
        return original(section, key)

    monkeypatch.setattr(execution_recovery, "capability_int", limits)
    first = "wps-execution-0001"
    second = "wps-execution-0002"
    execution_recovery.save(first, json.dumps(_package(first), ensure_ascii=False))
    with pytest.raises(RecoveryPackageQuota):
        execution_recovery.save(second, json.dumps(_package(second), ensure_ascii=False))
    # 第一条材料必须还在：不能为了腾地方淘汰未结束任务的恢复点。
    assert execution_recovery.status(first) == "ok"


def test_expired_packages_are_pruned_to_make_room(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    original = capability_int

    def limits(section: str, key: str) -> int:
        if (section, key) == ("taskRecovery", "recoveryPackageMaxOperations"):
            return 1
        return original(section, key)

    monkeypatch.setattr(execution_recovery, "capability_int", limits)
    stale = "wps-execution-0001"
    execution_recovery.save(stale, json.dumps(_package(stale), ensure_ascii=False))
    # 让第一份过期：改元数据里的到期时间（正文与摘要不动，仍是"完整但过期"的材料）。
    path = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / stale / "material.dat"
    header, body = path.read_bytes().split(b"\n", 1)
    record = json.loads(header)
    record["expiresAt"] = (datetime.now(timezone.utc) - timedelta(days=1)).isoformat().replace("+00:00", "Z")
    path.write_bytes(json.dumps(record, ensure_ascii=False).encode("utf-8") + b"\n" + body)

    fresh = "wps-execution-0002"
    execution_recovery.save(fresh, json.dumps(_package(fresh), ensure_ascii=False))
    assert execution_recovery.status(stale) == "missing"
    assert execution_recovery.status(fresh) == "ok"


def test_write_failure_leaves_no_material_and_no_partial_file(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """落盘失败必须抛出，绝不能返回"已保存"让调用方继续写原表。"""

    original_replace = Path.replace

    def deny(source: Path, target: Path) -> Path:
        if Path(target).name == "material.dat":
            raise PermissionError(13, "synthetic lock")
        return original_replace(source, target)

    monkeypatch.setattr(Path, "replace", deny)
    monkeypatch.setattr(execution_recovery.time, "sleep", lambda _: None)
    with pytest.raises(RecoveryPackagePersistenceError):
        execution_recovery.save(OPERATION, json.dumps(_package(), ensure_ascii=False))
    folder = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION
    assert not (folder / "material.dat").exists()
    assert not list(folder.glob("*.tmp"))


def test_invalid_material_is_refused_before_touching_disk(tmp_path: Path) -> None:
    broken = _package()
    broken["actions"][0]["values"] = [[1, 2]]
    with pytest.raises(RecoveryPackageError):
        execution_recovery.save(OPERATION, json.dumps(broken, ensure_ascii=False))
    assert execution_recovery.status(OPERATION) == "missing"


def test_http_roundtrip_binds_package_to_registered_execution() -> None:
    client = TestClient(app)
    _register(client)
    response = _put(client)
    assert response.status_code == 200, response.text
    saved = response.json()
    assert saved["packageJson"] == json.dumps(_package(), ensure_ascii=False)

    # 新客户端 = 重启后的后端：仍能读回同一份正文。
    restarted = TestClient(app)
    loaded = restarted.get(f"/api/execution-recovery/{OPERATION}")
    assert loaded.status_code == 200, loaded.text
    assert loaded.json()["packageJson"] == saved["packageJson"]


def test_expiry_reservation_failure_returns_readable_cors_error_without_material(monkeypatch):
    import sqlite3
    client = TestClient(app, raise_server_exceptions=False)
    _register(client)
    def broken_reservation(*_args):
        raise sqlite3.OperationalError("no such column: recovery_package_expires_at")
    monkeypatch.setattr(main_module._folder_operation_ledger, "reserve_recovery_package_expiry", broken_reservation)
    body = json.dumps(_package(), ensure_ascii=False)
    response = client.put(f"/api/execution-recovery/{OPERATION}",
        json={"packageJson": body, "size": len(body.encode("utf-8"))},
        headers={"Origin": "http://127.0.0.1:3890"})
    assert response.status_code == 503
    assert response.json()["detail"]["code"] == "TASK_STATE_UNAVAILABLE"
    assert response.headers["access-control-allow-origin"] == "http://127.0.0.1:3890"
    assert "no such column" not in response.text
    assert execution_recovery.status(OPERATION) == "missing"


def test_http_refuses_package_from_another_plan() -> None:
    client = TestClient(app)
    _register(client)
    other = _package()
    other["actions"][1]["previewFingerprint"] = "v2:1:99999999"
    other["planFingerprint"] = recovery_package_plan_fingerprint(other["actions"])
    response = _put(client, package=other)
    assert response.status_code == 422
    assert response.json()["detail"]["code"] == "RECOVERY_PACKAGE_INVALID"
    assert client.get(f"/api/execution-recovery/{OPERATION}").status_code == 404


def test_http_validation_error_never_echoes_cell_content() -> None:
    """Pydantic 的默认错误文本会回显被校验输入，这里必须换成固定文案。"""

    client = TestClient(app)
    _register(client)
    sentinel = "SENTINEL-ECHO-b51d0e"
    broken = _package()
    broken["actions"][0]["beforeValues"] = [[sentinel, "extra"]]
    response = _put(client, package=broken)
    assert response.status_code == 422
    assert sentinel not in response.text
    assert "beforeValues" not in response.text


def test_http_refuses_package_identity_mismatch() -> None:
    client = TestClient(app)
    _register(client)
    other = _package()
    other["taskId"] = "task-other"
    assert _put(client, package=other).status_code == 422
    other = _package()
    other["authorizedSheetNames"] = ["来源", "目标", "别的表"]
    assert _put(client, package=other).status_code == 422
    other = _package()
    other["workbookIdentity"] = "wps-path:c:/work/other.xlsx"
    assert _put(client, package=other).status_code == 422


def test_http_requires_a_registered_execution() -> None:
    client = TestClient(app)
    response = _put(client, "wps-execution-9999")
    assert response.status_code == 409
    assert client.get("/api/execution-recovery/wps-execution-9999").status_code == 409


def test_http_reports_expired_and_missing_separately() -> None:
    client = TestClient(app)
    _register(client)
    assert client.get(f"/api/execution-recovery/{OPERATION}").json()["detail"]["code"] == "RECOVERY_PACKAGE_MISSING"
    _put(client)
    path = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION / "material.dat"
    header, body = path.read_bytes().split(b"\n", 1)
    record = json.loads(header)
    record["expiresAt"] = (datetime.now(timezone.utc) - timedelta(minutes=1)).isoformat().replace("+00:00", "Z")
    path.write_bytes(json.dumps(record, ensure_ascii=False).encode("utf-8") + b"\n" + body)
    expired = client.get(f"/api/execution-recovery/{OPERATION}")
    assert expired.status_code == 410
    assert expired.json()["detail"]["code"] == "RECOVERY_PACKAGE_EXPIRED"


def test_terminal_removal_is_idempotent_and_leaves_no_folder() -> None:
    """终态回收：删掉之后目录也消失，重复删不报错（幂等）。"""

    execution_recovery.save(OPERATION, json.dumps(_package(), ensure_ascii=False))
    assert execution_recovery.remove(OPERATION) is True
    folder = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION
    assert not folder.exists()
    assert execution_recovery.status(OPERATION) == "missing"
    assert execution_recovery.remove(OPERATION) is False


def test_http_delete_retires_material_only_for_a_registered_execution() -> None:
    client = TestClient(app)
    _register(client)
    _put(client)
    assert client.get(f"/api/execution-recovery/{OPERATION}").status_code == 200
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204
    assert client.get(f"/api/execution-recovery/{OPERATION}").json()["detail"]["code"] == "RECOVERY_PACKAGE_MISSING"
    # 幂等：再删一次仍然成功（"已经不在了"不能变成错误）。
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204
    # 没登记过的执行不允许删别人的材料（先核身份，再动手）。
    assert client.delete("/api/execution-recovery/wps-execution-9999").status_code == 409


def test_recovery_material_never_enters_the_task_ledger() -> None:
    """恢复包含单元格原值：它只能待在本机恢复目录，不能渗进账本/诊断。

    `plan_json` 是 WPS 出站闸的产物（只含指纹），这里同时验证"恢复包确实带着原值"
    （锚点）与"账本里读不到它"。
    """

    client = TestClient(app)
    _register(client)
    sentinel = "SENTINEL-RECOVERY-9c31f7"
    package = _package()
    package["actions"][0]["beforeValues"] = [[sentinel]]
    package["actions"][0]["beforeFormulas"] = [[sentinel]]
    assert _put(client, package=package).status_code == 200
    material = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION / "material.dat"
    assert sentinel in material.read_text(encoding="utf-8")

    database = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "tasks" / "state.sqlite3"
    raw = database.read_bytes()
    assert sentinel.encode("utf-8") not in raw
    # 诊断事件也不许带正文。
    diagnostics = client.get("/api/diagnostics")
    assert sentinel not in diagnostics.text


def test_shared_fixture_is_accepted_by_server_model() -> None:
    """与 `packages/core/src/executionRecoveryPackage.test.ts` 读同一份样例。

    TS 与 Python 各有一份实现，这条用例把"两份协议是否还一致"钉在数据上：
    只要有一侧放宽/收紧，样例就会在一侧通过、另一侧失败。
    """

    cases = json.loads((REPO_ROOT / "config" / "execution_recovery_package_cases.json").read_text(encoding="utf-8"))
    for valid in cases["valid"]:
        ExecutionRecoveryPackage.model_validate(valid)
    for item in cases["invalid"]:
        with pytest.raises(Exception):
            ExecutionRecoveryPackage.model_validate(item["value"])


# ---------------------------------------------------------------------------
# 2026-10-27 第二轮：同一个 operationId 的**首次**有效期不可续期
#
# 缺陷：`save()` 只在"同正文且 deadline > now"时直接返回；同正文但已过期时会走到底，
# 重新算一个 `savedAt/expiresAt` 覆盖文件 —— 于是重复 PUT 变成免费的续期按钮。
# 这些用例用**显式时间**与**账本登记值**制造"已经过期"，不靠真实等待。
# ---------------------------------------------------------------------------


def _rewrite_expiry(operation_id: str, expires_at: datetime) -> None:
    """把磁盘上那份材料的到期时间改成指定值（模拟时间流逝，不删文件）。"""

    path = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / operation_id / "material.dat"
    header, body = path.read_bytes().split(b"\n", 1)
    record = json.loads(header)
    record["expiresAt"] = expires_at.isoformat().replace("+00:00", "Z")
    path.write_bytes(json.dumps(record, ensure_ascii=False).encode("utf-8") + b"\n" + body)


def test_identical_save_is_idempotent_and_never_renews(tmp_path: Path) -> None:
    """同正文重复保存：首次 `savedAt`/`expiresAt` 与磁盘字节都不许变。"""

    package_json = json.dumps(_package(), ensure_ascii=False)
    first = execution_recovery.save(OPERATION, package_json)
    path = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION / "material.dat"
    digest_before = hashlib.sha256(path.read_bytes()).hexdigest()
    for _ in range(3):
        again = execution_recovery.save(OPERATION, package_json)
        assert again["savedAt"] == first["savedAt"], "重复保存不能刷新 savedAt"
        assert again["expiresAt"] == first["expiresAt"], "重复保存不能顺延保留期"
        assert again["digest"] == first["digest"]
    assert hashlib.sha256(path.read_bytes()).hexdigest() == digest_before, "幂等保存不应重写文件"


def test_expired_material_cannot_be_re_saved_into_a_new_window(tmp_path: Path) -> None:
    """已过期的同正文再保存 ⇒ 明确拒绝，且磁盘上那份仍是原来的期限。"""

    package_json = json.dumps(_package(), ensure_ascii=False)
    execution_recovery.save(OPERATION, package_json)
    expired = datetime.now(timezone.utc) - timedelta(minutes=1)
    _rewrite_expiry(OPERATION, expired)
    with pytest.raises(RecoveryPackageExpired):
        execution_recovery.save(OPERATION, package_json)
    path = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION / "material.dat"
    header = json.loads(path.read_bytes().split(b"\n", 1)[0])
    assert header["expiresAt"] == expired.isoformat().replace("+00:00", "Z"), "拒绝保存不能顺手改期限"
    assert execution_recovery.status(OPERATION) == "expired"


def test_pruned_material_cannot_be_re_saved_past_the_first_expiry(tmp_path: Path) -> None:
    """保留期清理把文件删掉之后，同一 operationId 也不能白拿一个新的 7 天。

    HTTP 层会把账本登记的**首次**到期时间传进来；这里直接用它模拟那条链路：
    文件没了 + 首次期限已过 ⇒ 拒绝，且不新建任何文件。
    """

    package_json = json.dumps(_package(), ensure_ascii=False)
    execution_recovery.save(OPERATION, package_json)
    expired = datetime.now(timezone.utc) - timedelta(minutes=1)
    _rewrite_expiry(OPERATION, expired)
    assert execution_recovery.prune_expired()["removed"] == 1
    folder = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION
    assert not folder.exists()

    with pytest.raises(RecoveryPackageExpired):
        execution_recovery.save(OPERATION, package_json, expired.isoformat().replace("+00:00", "Z"))
    assert not folder.exists(), "被拒绝的保存不许留下新材料"


def test_corrupt_material_is_not_overwritten_with_a_fresh_window(tmp_path: Path) -> None:
    package_json = json.dumps(_package(), ensure_ascii=False)
    execution_recovery.save(OPERATION, package_json)
    path = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION / "material.dat"
    data = path.read_bytes().replace(b"old2", b"hax2")
    path.write_bytes(data)
    with pytest.raises(RecoveryPackageCorrupt):
        execution_recovery.save(OPERATION, package_json)
    assert path.read_bytes() == data, "损坏材料既不能当有效旧记录，也不能被顺手覆盖"


def test_expiry_reservation_is_write_once() -> None:
    """账本只记**第一次**到期时间，后来的调用只会读回它。"""

    from server.app.task_ledger import FolderOperationLedger

    client = TestClient(app)
    _register(client)
    ledger = FolderOperationLedger()
    first = ledger.reserve_recovery_package_expiry(OPERATION, "2026-11-03T00:00:00.000Z")
    assert first == "2026-11-03T00:00:00.000Z"
    again = ledger.reserve_recovery_package_expiry(OPERATION, "2027-01-01T00:00:00.000Z")
    assert again == first, "第二次调用不许把期限往后推"
    assert ledger.recovery_package_expiry(OPERATION) == first
    assert ledger.reserve_recovery_package_expiry("wps-execution-9999", "2027-01-01T00:00:00.000Z") is None


def test_http_repeat_put_keeps_the_first_expiry() -> None:
    """真实 HTTP：重复 PUT 不续期；材料被回收后再 PUT 仍沿用首次期限。"""

    client = TestClient(app)
    _register(client)
    first = _put(client)
    assert first.status_code == 200
    expires = first.json()["expiresAt"]
    saved = first.json()["savedAt"]
    for _ in range(2):
        again = _put(client)
        assert again.status_code == 200
        assert again.json()["expiresAt"] == expires
        assert again.json()["savedAt"] == saved
    # 文件被回收（终态/清理）后仍然只有首次那个期限 —— 不许从"现在"重新起算 7 天。
    assert client.delete(f"/api/execution-recovery/{OPERATION}").status_code == 204
    revived = _put(client)
    assert revived.status_code == 200
    assert revived.json()["expiresAt"] == expires, "文件不在也不能重新获得一个新的保留期"


def test_http_refuses_put_when_the_first_expiry_has_passed(monkeypatch: pytest.MonkeyPatch) -> None:
    """账本登记的首次期限已经过去 ⇒ 任何后续 PUT 都是 410，而且不落盘。

    这里把"首次期限"直接设成过去（`default_expiry` 只在**第一次**登记时被调用），
    不靠真实等待，也不改动已经跑起来的用户服务。
    """

    past = (datetime.now(timezone.utc) - timedelta(days=8)).isoformat().replace("+00:00", "Z")
    real_default_expiry = execution_recovery.default_expiry
    monkeypatch.setattr(execution_recovery, "default_expiry", lambda moment=None: past)
    client = TestClient(app)
    _register(client)
    refused = _put(client)
    assert refused.status_code == 410
    assert refused.json()["detail"]["code"] == "RECOVERY_PACKAGE_EXPIRED"
    folder = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / OPERATION
    assert not folder.exists(), "被拒绝的 PUT 不许落盘"

    # "名额"已经被第一次登记占住了：即便之后期限算法恢复正常，也拿不到新期限。
    # 这里只还原这一个属性 —— `monkeypatch.undo()` 会把隔离用的环境变量一起撤掉。
    monkeypatch.setattr(execution_recovery, "default_expiry", real_default_expiry)
    still = _put(client)
    assert still.status_code == 410
    assert not folder.exists()

# ---------------------------------------------------------------------------
# 2026-10-27 第三轮：无时区到期时间 → TypeError；单条损坏材料会中断整批清理
#
# 缺陷（独立临时目录可复现）：把首行元数据的 `expiresAt` 改成不含时区的
# `2035-01-01T00:00:00`，正文/摘要/长度一字不动，随后
#   - `load()`、同正文 `save()`、`prune_expired()` 全都抛
#     `TypeError: can't compare offset-naive and offset-aware datetimes`；
#   - 也就是说"读得出但这个值没法用"被当成了**可比较的时刻**，而不是损坏材料。
# 正常生成的 UTC 时间不受影响 —— 这里钉的是那一条边缘形态，不是"恢复全坏"。
# ---------------------------------------------------------------------------


def _material_path(operation_id: str = OPERATION) -> Path:
    return Path(os.environ["EXCEL_BRO_CONFIG_DIR"]) / "execution-recovery" / operation_id / "material.dat"


def _rewrite_header(operation_id: str, **fields: object) -> bytes:
    """只改首行元数据的指定字段：正文、摘要、字节数一字不动。返回改后的完整字节。"""

    path = _material_path(operation_id)
    header, body = path.read_bytes().split(b"\n", 1)
    record = json.loads(header)
    record.update(fields)
    data = json.dumps(record, ensure_ascii=False).encode("utf-8") + b"\n" + body
    path.write_bytes(data)
    return data


def test_naive_expiry_reads_as_corrupt_not_a_type_error(tmp_path: Path) -> None:
    """无时区的期限 = **无法确定**，按损坏处理；不许按本机时区猜，也不许漏 TypeError。"""

    package_json = json.dumps(_package(), ensure_ascii=False)
    execution_recovery.save(OPERATION, package_json)
    before = _rewrite_header(OPERATION, expiresAt="2035-01-01T00:00:00")

    with pytest.raises(RecoveryPackageCorrupt):
        execution_recovery.load(OPERATION)
    assert execution_recovery.status(OPERATION) == "corrupt"
    assert _material_path().read_bytes() == before, "读回不能顺手改材料"


def test_naive_expiry_refuses_resave_and_leaves_bytes_identical(tmp_path: Path) -> None:
    """同正文重存遇到"期限无法确定" ⇒ 拒绝覆盖，也不给它一个新期限。"""

    package_json = json.dumps(_package(), ensure_ascii=False)
    execution_recovery.save(OPERATION, package_json)
    before = _rewrite_header(OPERATION, expiresAt="2035-01-01T00:00:00")

    with pytest.raises(RecoveryPackageCorrupt):
        execution_recovery.save(OPERATION, package_json)
    assert _material_path().read_bytes() == before, "损坏材料不许被顺手覆盖"


def test_prune_clears_only_confirmed_expiry_and_keeps_undecidable_material(tmp_path: Path) -> None:
    """同目录三份材料：确认过期（清）、未过期（留）、期限读不出（留 + 计入未确认）。

    单条材料读不出期限时，清理既不能整批中断，也不能把它当成"已过期"删掉 ——
    删掉就失去排查依据；也不能为它重算一个期限。
    """

    stale, fresh, broken = "wps-execution-0001", "wps-execution-0002", "wps-execution-0003"
    for operation_id in (stale, fresh, broken):
        execution_recovery.save(operation_id, json.dumps(_package(operation_id), ensure_ascii=False))
    _rewrite_expiry(stale, datetime.now(timezone.utc) - timedelta(days=1))
    broken_bytes = _rewrite_header(broken, expiresAt="2035-01-01T00:00:00")

    result = execution_recovery.prune_expired()

    assert result["removed"] == 1, "只有确认过期的那一份可以清"
    assert result["failed"] == 0
    assert result["unconfirmed"] == 1, "无法确定期限既不算清理成功，也不算清理失败"
    assert execution_recovery.status(stale) == "missing"
    assert execution_recovery.status(fresh) == "ok"
    assert _material_path(broken).read_bytes() == broken_bytes, "损坏材料必须原样保留"


def test_expiry_offsets_compare_as_the_same_instant(tmp_path: Path) -> None:
    """`Z`、`+00:00`、`+08:00` 指向同一时刻时，判定必须一致。"""

    execution_recovery.save(OPERATION, json.dumps(_package(), ensure_ascii=False))
    moment = datetime(2035, 1, 1, tzinfo=timezone.utc)

    for text in ("2035-01-01T00:00:00Z", "2035-01-01T00:00:00+00:00", "2035-01-01T08:00:00+08:00"):
        _rewrite_header(OPERATION, expiresAt=text)
        assert execution_recovery.load(OPERATION, now=moment - timedelta(seconds=1))["digest"]
        with pytest.raises(RecoveryPackageExpired):
            execution_recovery.load(OPERATION, now=moment)


def test_exact_expiry_instant_is_already_expired(tmp_path: Path) -> None:
    """`now == expiresAt` 即过期，与服务端其它入口同一口径。"""

    execution_recovery.save(OPERATION, json.dumps(_package(), ensure_ascii=False))
    moment = datetime(2035, 1, 1, tzinfo=timezone.utc)
    _rewrite_header(OPERATION, expiresAt="2035-01-01T00:00:00Z")

    with pytest.raises(RecoveryPackageExpired):
        execution_recovery.load(OPERATION, now=moment)
    assert execution_recovery.status(OPERATION, now=moment) == "expired"


@pytest.mark.parametrize(
    "value",
    [None, 20350101, "2035-01-01T00:00:00", "not-a-time", "", {"iso": "2035-01-01T00:00:00Z"}, [2035]],
)
def test_unusable_expiry_values_are_corrupt_and_never_pruned(tmp_path: Path, value: object) -> None:
    """缺失 / null / 错误类型 / 非法文本 / 无时区：一律"损坏"，且清理时原样保留。"""

    package_json = json.dumps(_package(), ensure_ascii=False)
    execution_recovery.save(OPERATION, package_json)
    before = _rewrite_header(OPERATION, expiresAt=value)

    with pytest.raises(RecoveryPackageCorrupt):
        execution_recovery.load(OPERATION)
    with pytest.raises(RecoveryPackageCorrupt):
        execution_recovery.save(OPERATION, package_json)
    result = execution_recovery.prune_expired()
    assert result["removed"] == 0, "期限读不出来不等于已经过期"
    assert result["unconfirmed"] == 1
    assert _material_path().read_bytes() == before


def test_registered_expiry_must_carry_an_explicit_timezone(tmp_path: Path) -> None:
    """账本传进来的首次期限同样要带时区：无时区直接拒绝，不猜、不落盘。"""

    package_json = json.dumps(_package(), ensure_ascii=False)
    with pytest.raises(RecoveryPackageError):
        execution_recovery.save(OPERATION, package_json, "2035-01-01T00:00:00")
    assert not _material_path().parent.exists(), "被拒绝的保存不许留下目录"

    with pytest.raises(RecoveryPackageError):
        execution_recovery.default_expiry(datetime(2035, 1, 1))


def test_naive_now_is_refused_instead_of_guessed(tmp_path: Path) -> None:
    """内部 `now` 也守调用约定：naive 一律拒绝，避免"只修文件读取"漏掉另一个比较入口。"""

    execution_recovery.save(OPERATION, json.dumps(_package(), ensure_ascii=False))
    naive = datetime(2035, 1, 1)

    with pytest.raises(RecoveryPackageError):
        execution_recovery.load(OPERATION, now=naive)
    with pytest.raises(RecoveryPackageError):
        execution_recovery.prune_expired(now=naive)


def test_http_naive_expiry_maps_to_corrupt_without_echoing_material() -> None:
    """接口层：损坏材料映射到既有 409 明确响应，不能变成未处理的 500，也不能回显正文。"""

    client = TestClient(app)
    _register(client)
    sentinel = "SENTINEL-RECOVERY-na1ve"
    package = _package()
    package["actions"][0]["beforeValues"] = [[sentinel]]
    package["actions"][0]["beforeFormulas"] = [[sentinel]]
    assert _put(client, package=package).status_code == 200

    before = _rewrite_header(OPERATION, expiresAt="2035-01-01T00:00:00")

    read = client.get(f"/api/execution-recovery/{OPERATION}")
    assert read.status_code == 409
    assert read.json()["detail"]["code"] == "RECOVERY_PACKAGE_CORRUPT"
    assert sentinel not in read.text
    written = _put(client, package=package)
    assert written.status_code == 409
    assert written.json()["detail"]["code"] == "RECOVERY_PACKAGE_CORRUPT"
    assert sentinel not in written.text
    assert _material_path().read_bytes() == before, "损坏材料不许被重建"

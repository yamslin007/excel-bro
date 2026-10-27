"""S1 acceptance CLI for the explicit isolated native-test launcher.

Use a1e3_native.ps1 prepare/frontend/status/arm/evidence/stop. Direct Python
calls require --manifest. The test-only audit keeps the exact bound/consumed
turn ID after the production fault clears its in-memory binding. Evidence is
read-only, scoped to that task and host, and never substitutes for native UI
and workbook observations. Missing material is insufficient, never zero writes.
"""

from __future__ import annotations

import argparse
import json
import sqlite3
import sys
import uuid
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_BASE = "http://127.0.0.1:8877"
DEFAULT_TTL = 120
_SESSION: dict | None = None


# ---------------------------------------------------------------------------
# HTTP 助手（只读 GET + 显式 POST，永不静默重试，不吞错误）
# ---------------------------------------------------------------------------
class ServiceError(RuntimeError):
    """测试服务端返回非预期状态，或网络不可达。"""


def _request(base: str, path: str, method: str = "GET", body: dict | None = None) -> dict:
    url = f"{base.rstrip('/')}{path}"
    data = None
    headers = {"Accept": "application/json"}
    if _SESSION and base == _SESSION["base"]:
        headers["X-Native-Acceptance-Token"] = _SESSION["token"]
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        headers["Content-Type"] = "application/json"
    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=5) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", "replace")
        raise ServiceError(
            f"{method} {path} → HTTP {error.code}：{detail}"
        ) from error
    except urllib.error.URLError as error:
        raise ServiceError(f"{method} {path} 不可达（{error.reason}）：服务端没在 {base} 上。") from error
    if not raw.strip():
        return {}
    try:
        return json.loads(raw)
    except (ValueError, TypeError) as error:
        raise ServiceError(f"{path}: invalid JSON response") from error


def _status(base: str) -> dict:
    return _request(base, "/api/dev/source-changed-fault")


def _health(base: str) -> bool:
    """健康检查用能力端点（无副作用 GET）。"""
    try:
        _request(base, "/api/capabilities")
        return True
    except ServiceError:
        return False


# ---------------------------------------------------------------------------
# 账本只读定位（只发现 turnId / executionId，不参与判据）
# ---------------------------------------------------------------------------
def _ledger_path(config_dir: str | None) -> Path | None:
    if config_dir:
        return Path(config_dir) / "tasks" / "state.sqlite3"
    return None


def _readonly_rows(db_path: Path, sql: str, params: tuple = ()) -> list[sqlite3.Row]:
    uri = f"file:{db_path.as_posix()}?mode=ro"
    connection = sqlite3.connect(uri, uri=True, timeout=3)
    connection.row_factory = sqlite3.Row
    try:
        return list(connection.execute(sql, params))
    finally:
        connection.close()


# ---------------------------------------------------------------------------
# 证据导出（判据仍由服务端 observation 给，这里只呈现）
# ---------------------------------------------------------------------------
def _describe_observation(observation: dict) -> str:
    """把服务端的 observation 折成一句话；**不做零写入判断**，只复述服务端结论。

    ``complete + zero`` 才显示"明确零调用"，其余如实标"未采集/覆盖不全/未知/非零"。
    """
    observation_state = observation.get("observation")
    evidence = observation.get("callEvidence")
    unit = observation.get("unit", "")
    unit_label = "刷新批次" if unit == "excel_sync_batches" else "写 API 调用"
    if observation_state == "not_collected":
        attempted = observation.get("attemptedActions", 0)
        return (
            f"未采集（账本显示 {attempted} 个动作可能写过）" if attempted
            else "未采集（账本也没有动作被认领）"
        )
    if evidence == "nonzero":
        missing = observation.get("missingActions", 0)
        return (
            f"已观测到非零{unit_label}，但还缺 {missing} 个动作的观测" if missing
            else f"已观测到非零{unit_label}"
        )
    if evidence == "zero" and observation_state == "complete" and unit in ("excel_sync_batches", "wps_api_calls"):
        return f"已观测到零{unit_label}（覆盖 {observation.get('attemptedActions', 0)} 个动作）"
    if evidence == "unknown":
        missing = observation.get("missingActions", 0)
        return (
            f"观测不完整（缺 {missing} 个动作），结果未知" if missing
            else "观测结果未知（有动作没拿到确定的宿主结果）"
        )
    return "未采集到宿主观测"


def _session(base: str) -> dict:
    if not _SESSION:
        raise ServiceError("Missing run manifest; evidence cannot be associated with this run")
    state = _request(base, "/api/dev/native-acceptance/session")
    if state.get("runId") != _SESSION["runId"] or state.get("auditError"):
        raise ServiceError("Wrong test instance or incomplete audit")
    return state


def _collect_evidence(base: str, config_dir: str | None, host_id: str | None, out_dir: Path) -> dict:
    db_path = _ledger_path(config_dir)
    if db_path is None or not db_path.is_file():
        raise ServiceError("Evidence insufficient: isolated ledger is missing")
    session = _session(base)
    attempt = session.get("attempt", {})
    task_id = attempt.get("boundTaskId")
    if not task_id or attempt.get("consumedTaskId") != task_id or attempt.get("cancelled"):
        raise ServiceError("Evidence insufficient: this attempt was not bound and consumed")
    target = attempt["target"]
    if host_id and host_id != target["hostId"]:
        raise ServiceError("Host filter does not match this attempt")
    namespace = {"excel-office": "office", "wps-et": "wps"}.get(target["hostId"])
    if not namespace:
        raise ServiceError("Unknown host namespace")
    rows = _readonly_rows(db_path, "SELECT turn_id,status,last_sequence,request_json,error_json,created_at,updated_at FROM planning_tasks WHERE turn_id=?", (task_id,))
    if len(rows) != 1:
        raise ServiceError("Evidence insufficient: bound task is absent from isolated ledger")
    row = dict(rows[0])
    request = json.loads(row.pop("request_json"))
    workbook = request.get("workbook", {})
    if request.get("hostId") != target["hostId"] or workbook.get("name") != target["workbookName"] or workbook.get("workbookIdentity") != target["workbookIdentity"]:
        raise ServiceError("Bound task target disagrees with isolated ledger")
    # Ledger IDs must agree with the trusted bridge audit; no latest-record heuristics.
    if row["created_at"] < attempt["armedAt"] - 1:
        raise ServiceError("Bound task predates this attempt")
    error = json.loads(row.pop("error_json") or "null")
    executions = [dict(item) for item in _readonly_rows(db_path,
        "SELECT operation_id,task_id,status,next_action_index,created_at,updated_at FROM office_executions WHERE task_id=? ORDER BY created_at", (task_id,))]
    summaries = []
    for execution in executions:
        summary = _request(base, f"/api/{namespace}/executions/{execution['operation_id']}/host-writes")
        summary["_namespace"] = namespace
        summary["_observationText"] = _describe_observation(summary.get("observation", {}))
        summaries.append(summary)
    stopped = row["status"] == "failed" and isinstance(error, dict) and error.get("code") == "SOURCE_CHANGED"
    result = {"collectedAt": datetime.now(timezone.utc).isoformat(), "base": base,
              "runId": session["runId"], "attempt": attempt, "fault": _status(base),
              "hostId": target["hostId"], "turn": row, "errorCode": error.get("code") if isinstance(error, dict) else None,
              "executions": executions, "hostWriteSummaries": summaries,
              "planningStopConfirmed": stopped,
              "executionEvidence": "未进入执行" if stopped and not executions else "需核对执行观测",
              "nativeUiAndWorkbook": "pending"}
    # Every export gets a new directory; previous evidence is immutable.
    export_dir = out_dir / (datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S") + "-" + uuid.uuid4().hex[:8])
    export_dir.mkdir(parents=True, exist_ok=False)
    result["evidencePath"] = str(export_dir / "evidence.json")
    (export_dir / "evidence.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return result


# ---------------------------------------------------------------------------
# 子命令
# ---------------------------------------------------------------------------
def cmd_status(base: str) -> int:
    if not _health(base):
        print(f"[停下] 测试服务端不可达（{base}）。先启动测试服务端。")
        return 2
    status = _status(base)
    print(json.dumps(status, ensure_ascii=False, indent=2))
    return 0


def cmd_arm(
    base: str,
    host: str | None,
    name: str | None,
    identity: str | None,
    ttl: int,
) -> int:
    if not _health(base):
        print(f"[停下] 测试服务端不可达（{base}）。先启动测试服务端。")
        return 2

    status = _status(base)
    seen = status.get("seenTargets", [])
    if not seen:
        print("[停下] 服务端还没见过任何工作簿（seenTargets 为空）。")
        print("  先在真实窗格里发一条只读需求（例如「来源表有多少行」），让桥立起来，再重跑本命令。")
        return 3

    if host is None and name is None and identity is None:
        for index, target in enumerate(seen, start=1):
            print(f"[{index}] {target['hostId']} | {target['workbookName']} | {target['workbookIdentity']}")
        if not sys.stdin.isatty():
            print("[停下] 非交互调用必须显式指定 --host / --name / --identity。尚未启用。")
            return 3
        try:
            selected = int(input("选择要验收的工作簿序号（输入 0 取消）："))
        except (ValueError, EOFError):
            return 3
        if not 1 <= selected <= len(seen):
            return 3
        chosen = seen[selected - 1]
        host, name, identity = chosen["hostId"], chosen["workbookName"], chosen["workbookIdentity"]

    # 按 host/name 过滤候选。
    candidates = [
        target
        for target in seen
        if (host is None or target["hostId"] == host)
        and (name is None or target["workbookName"] == name)
    ]
    if identity is not None:
        candidates = [
            target for target in candidates if target["workbookIdentity"] == identity
        ]
        if not candidates:
            print(f"[停下] 给的 identity 不在服务端见过的目标里：{identity}")
            print("  可用目标：")
            for target in seen:
                print(f"    host={target['hostId']} name={target['workbookName']} identity={target['workbookIdentity']}")
            return 3

    if len(candidates) == 0:
        print("[停下] 没有匹配的目标（host/name 组合不对，或服务端还没见过这本簿）。")
        return 3
    if len(candidates) > 1:
        print(f"[停下] 有 {len(candidates)} 个候选，不能默认取第一条：")
        for target in candidates:
            print(f"  host={target['hostId']} name={target['workbookName']} identity={target['workbookIdentity']}")
        print("  请用 --identity 显式指定其一（值抄上面的 identity 字段）。")
        return 3

    target = candidates[0]
    payload = {
        "hostId": target["hostId"],
        "workbookName": target["workbookName"],
        "workbookIdentity": target["workbookIdentity"],
    }
    print("将启用故障注入，绑定到：")
    print(f"  host     = {payload['hostId']}")
    print(f"  name     = {payload['workbookName']}")
    print(f"  identity = {payload['workbookIdentity']}")

    if _SESSION:
        observed = _session(base)["bridgeTargets"]
        if payload not in observed:
            raise ServiceError("Frontend has not opened a real bridge on this test instance for this target")
    armed = _request(base, "/api/dev/source-changed-fault/arm", "POST", payload)
    print(json.dumps(armed, ensure_ascii=False, indent=2))
    phase = armed.get("phase")
    identity_bound = armed.get("identityBound")
    if phase != "armed" or not identity_bound:
        print(f"[停下] 启用未成功：phase={phase} identityBound={identity_bound}")
        return 3
    expires = armed.get("expiresInSeconds", 0)
    if expires <= 0:
        print(f"[停下] 启用已过期（expiresInSeconds={expires}）。重新 arm。")
        return 3
    print(f"[就绪] 已启用（armed / identityBound），有效期 {expires} 秒。")
    print("  请在有效期内发送 S1 需求：")
    print("  「读取「来源」B2:B4，将每个值乘以 2，写入「目标」B2:B4，来源保持不变。」")
    return 0


def cmd_cancel(base: str) -> int:
    if not _health(base):
        print(f"[停下] 测试服务端不可达（{base}）。")
        return 2
    print(json.dumps(_request(base, "/api/dev/source-changed-fault/cancel", "POST", {}), ensure_ascii=False, indent=2))
    return 0


def cmd_evidence(
    base: str,
    config_dir: str | None,
    host: str | None,
    out: str,
) -> int:
    if not _health(base):
        print(f"[停下] 测试服务端不可达（{base}）。")
        return 2
    try:
        evidence = _collect_evidence(base, config_dir, host, Path(out))
    except (ServiceError, sqlite3.Error, OSError, ValueError) as error:
        directory = Path(out) / ("insufficient-" + uuid.uuid4().hex)
        directory.mkdir(parents=True, exist_ok=False)
        (directory / "evidence.json").write_text(json.dumps({
            "collectedAt": datetime.now(timezone.utc).isoformat(),
            "runId": (_SESSION or {}).get("runId"), "status": "insufficient",
            "error": str(error), "nativeUiAndWorkbook": "pending"
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        raise
    print(f"[证据已导出] {evidence['evidencePath']}")

    print(json.dumps(evidence, ensure_ascii=False, indent=2))
    if not evidence["planningStopConfirmed"]:
        print("[未完成] 故障已触发，但本任务尚无 SOURCE_CHANGED 失败收尾证据。")
        return 4
    print(evidence["executionEvidence"] + "；真实界面与单元格结果另行核对。")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="A1-E3 S1 原生验收 CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    status_p = sub.add_parser("status", help="查故障入口状态")
    status_p.add_argument("--base", default=None)

    arm_p = sub.add_parser("arm", help="显式启用一次故障注入")
    arm_p.add_argument("--base", default=None)
    arm_p.add_argument("--host", default=None, help="宿主 ID（excel-office / wps-et）")
    arm_p.add_argument("--name", default=None, help="工作簿显示名")
    arm_p.add_argument("--identity", default=None, help="工作簿身份（抄自 seenTargets，不手写）")
    arm_p.add_argument("--ttl", type=int, choices=[DEFAULT_TTL], default=DEFAULT_TTL, help="服务端固定有效期，当前为120秒")

    cancel_p = sub.add_parser("cancel", help="取消故障注入")
    cancel_p.add_argument("--base", default=None)

    ev_p = sub.add_parser("evidence", help="导出验收证据")
    ev_p.add_argument("--base", default=None)
    ev_p.add_argument("--config-dir", default=None, help="测试服务端配置目录（含 tasks/state.sqlite3）")
    ev_p.add_argument("--host", default=None)
    ev_p.add_argument("--out", default="outputs/a1e3-native/evidence")

    for command_parser in (status_p, arm_p, cancel_p, ev_p):
        command_parser.add_argument("--manifest", required=True, help="本次 prepare 生成的 manifest.json")
    return parser


def main(argv: list[str] | None = None) -> int:
    global _SESSION
    args = build_parser().parse_args(argv)
    try:
        manifest_path = Path(args.manifest).resolve()
        _SESSION = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
        if args.base and args.base != _SESSION["base"]:
            raise ServiceError("--base disagrees with run manifest")
        args.base = _SESSION["base"]
        _session(args.base)
        if args.command == "status":
            return cmd_status(args.base)
        if args.command == "arm":
            return cmd_arm(args.base, args.host, args.name, args.identity, args.ttl)
        if args.command == "cancel":
            return cmd_cancel(args.base)
        if args.command == "evidence":
            expected_config = str(manifest_path.parent / "server-config")
            if args.config_dir and Path(args.config_dir).resolve() != Path(expected_config).resolve():
                raise ServiceError("--config-dir disagrees with manifest")
            return cmd_evidence(args.base, expected_config, args.host, args.out)
    except (ServiceError, OSError, ValueError, KeyError, sqlite3.Error) as error:
        print(f"[停下] {error}", file=sys.stderr)
        return 2
    return 2


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    raise SystemExit(main())

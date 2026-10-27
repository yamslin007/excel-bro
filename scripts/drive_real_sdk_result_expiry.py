"""A1-E2：真实 SDK 的**结果引用过期恢复全链**驱动（2026-10-25）。

这条链与本仓库既有真实验收脚本的分工：

- `scripts/drive_real_wps_sdk.py`：真实模型 + 真实服务端，跑到"计划预览"为止。
- **本文件**：在同一条链上先**精确制造一次引用过期**（A1-E1 的开发入口），
  再看真实模型是否**重新读表重算**、拿到**新引用**、并由服务端**编译出新计划**。

链路（每一步都是生产代码，唯一替身是宿主读表）：

    面板等价请求 POST /api/turn/stream（真实模型规划）
      → 桥轮询 /api/sdk/tool-bridge/poll → 本脚本按指针真读宿主替身
      → **生产共享判据复核授权/身份/来源**（`packages/core/src/sdkToolBridge.ts`）
      → 回执 /receipt → 服务端登记数据集
      → 模型调 run_python_analysis（真实 Docker 隔离执行）→ 结果登记 R1
      → 模型用 R1 提交结果/计划 → **A1-E1 故障在此消费**：R1 删除 + 固定码拒绝
      → 模型重新 prepare + run_python_analysis → 结果登记 R2
      → 模型用 R2 提交计划 → 服务端 compile_result_plan 编译白名单计划
      → 本脚本做引用探针（R1 必须 404、R2 必须可读）并落工件

**不做**的事：
- 不替模型规划、不预制结果矩阵/计划；
- 不写任何宿主单元格（写回由两端面板在消费侧用例里做，本脚本统计为 0）；
- 不改生产 TTL、不清空结果；失效只来自 A1-E1 那个被显式启用的单次入口；
- 不记录密钥或额外业务原始数据（工件只留引用、动作序号、计数、结论）。

授权/来源复核**不在这里另写一份**：面板那一侧调的是共享核心
`verifyBridgeAuthorization`，本脚本通过 `scripts/bridge_authorization_gate.mjs` 直接
跑同一份 TS 判据（Node 22.18+ 原生支持 type stripping）。早先这里用 Python 复刻了
"指纹不等就回 source_changed"，那是第二份行为相似的判据，生产改一次就会漂移。

场景：
- `recovery`：A1-E1 让旧引用失效 → 模型重算 → 新引用 → 新计划（主场景）。
- `source-change`：首次交值后由**确定性屏障**改一个来源格子，模型**没有**再次读取
  （A 路）。
- `source-change-reread`：同一个屏障，但测试需求里写明了"提交写入预览前再读取一次
  核对"，于是模型**再次调用**受控读取 → 生产判据回 `source_changed`（B 路）。
  该场景**不启用** A1-E1 故障：它要单独隔离"来源变化"这一个因素。
- `pause`：在故障被消费那一刻发 `/pause`。

用法（recovery / source-change / pause 需服务端带 `EXCEL_BRO_RESULT_REF_FAULT=1`）：

    python scripts/drive_real_sdk_result_expiry.py --scenario recovery --out outputs/a1-e2-chain.json
    python scripts/drive_real_sdk_result_expiry.py --scenario source-change
    python scripts/drive_real_sdk_result_expiry.py --scenario source-change-reread
    python scripts/drive_real_sdk_result_expiry.py --scenario pause

工件缺失时消费侧用例跳过；`scripts/run_a1_e2_chain.mjs` 把"跳过"判成未完成并非零退出。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sqlite3
import subprocess
import sys
import threading
import time
import urllib.error
from pathlib import Path
from typing import Any

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))

import drive_real_wps_sdk as base  # noqa: E402  （同目录脚本，复用其宿主替身与 HTTP 封装）

OPERATIONS = "operations"
MEMBERS = base.MEMBERS
ORDERS = base.ORDERS
TARGET = base.TARGET
WORKBOOK_NAME = base.WORKBOOK_NAME

# 宿主与工作簿身份默认取 WPS 端（与 `drive_real_wps_sdk.py` 同一份替身数据）。
# 两端要分别留证据：Excel 端用 `--host excel-office --identity url:...` 再跑一轮，
# 各出各的工件 —— 一端的证据不能替另一端作证。
HOST_ID = base.HOST_ID
WORKBOOK_IDENTITY = base.WORKBOOK_IDENTITY


def configure(*, host_id: str, identity: str) -> None:
    """切换宿主身份（只影响本脚本发出的请求/回执，不碰服务端状态）。"""
    global HOST_ID, WORKBOOK_IDENTITY
    HOST_ID = host_id
    WORKBOOK_IDENTITY = identity

DEFAULT_MODEL_ID = "claude-agent-sdk:model-bad94777bbf6478c98cb0fa9025185a2"

# 这一轮绑定的是哪几张来源（与提示词里"这两张已授权来源表"一致）。
SOURCE_REGIONS = ((MEMBERS, base.MEMBER_RANGE), (ORDERS, base.ORDER_RANGE))

# 来源变化分支改成哪个值。
#
# 为什么是这个数：订单表 B2 原值 100（会员 7）。改成 250 之后，"按等级汇总"的
# 金卡从 400 变成 550、总额从 1000 变成 1150 —— 差值是**可判定**的，
# 面板测试也按同一对数字核对"这份计划写的是旧那一版"。
SOURCE_CHANGE_TO = 250

# B 路（模型确实再次读取）的测试需求。
#
# 只在 `base.PROMPT` 后面**加了一段用户要求**，不动前半句、不换来源表、不改授权范围：
# 用户要的是"提交写入预览前再核对一次来源"。两次读取读的是**同一批授权表**
# （订单表 A1:B5 + 会员表 A1:C6），于是两次读取之间唯一的变量就是**数据本身**
# —— 这正是 B 路要隔离的因素。
#
# 为什么必须写"新的读取，不要复用之前登记的数据集"：模型手里已经有第一次的
# `dataset://` 句柄，不写清就可能直接复用（A 路就是这么发生的）。这不是"诱导工具
# 顺序"—— 用户的需求本身就要求再读一遍源；读不读、读几次仍然由模型自己决定。
REREAD_NOTE = (
    "在提交写入预览之前，请再重新读取一次这两张来源表核对金额：这是一次新的读取，"
    "不要复用之前登记的数据集，也不要拿记忆里的数值代替。核对一致就按这个结果生成"
    "汇总表的写入预览；若第二次读到的金额与第一次不同，以第二次读取的为准重新汇总，"
    "并在结果里注明两次不一致。如果这次重新读取被拒绝或读不到数据，不要反复重试读取；"
    "直接按第一次读到的数据生成汇总表的写入预览，并在结果里说明"
    "「来源核对未完成，写入的是第一次读取的结果」。"
)


def prompt_for(scenario: str) -> tuple[str, str]:
    """这一轮发给模型的测试需求。第二个返回值只用于工件里标注用了哪一版需求。"""
    if scenario == "source-change-reread":
        return f"{base.PROMPT}{REREAD_NOTE}", "two-read-verification"
    return base.PROMPT, "single-read"


GATE_SCRIPT = Path(__file__).resolve().parent / "bridge_authorization_gate.mjs"
# 桥绑定的授权集合：与本轮请求里的 `worksheets` 一致（三张表都在这轮的授权范围内）。
AUTHORIZED_SHEETS = [MEMBERS, ORDERS, TARGET]


def authorization_verdict(binding: dict[str, Any], call: dict[str, Any], live: dict[str, Any]) -> dict[str, Any]:
    """跑**生产**共享判据得到这一次桥调用的裁决。

    入口 `scripts/bridge_authorization_gate.mjs` 直接 import
    `packages/core/src/sdkToolBridge.ts` 的 `verifyBridgeAuthorization` —— 与真实面板
    执行前复核同一份代码。这里不做任何"等价改写"。

    入口本身出错（Node 缺失/版本太老/输入形状不对）一律**中止实验**：那说明驱动写错了，
    不是"拒绝了一次读取"，绝不能被当成 source_changed 混进工件。
    """
    payload = json.dumps({"binding": binding, "call": call, "live": live}, ensure_ascii=False)
    node = os.getenv("EB_NODE", "node")
    try:
        done = subprocess.run(
            [node, str(GATE_SCRIPT)],
            input=payload,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=60,
        )
    except FileNotFoundError:
        raise SystemExit("找不到 node：授权/来源复核要跑生产共享判据（Node 22.18+）。")
    if done.returncode != 0:
        raise SystemExit(
            "授权/来源复核入口执行失败"
            f"（exit={done.returncode}）：{(done.stdout or done.stderr or '')[:400]}"
            "　—— 共享判据需要 Node 22.18+ 直接 import .ts。"
        )
    lines = [line for line in (done.stdout or "").splitlines() if line.strip()]
    try:
        verdict = json.loads(lines[-1])
    except (IndexError, ValueError):
        raise SystemExit(f"授权/来源复核入口输出无法解析：{(done.stdout or '')[:400]}")
    if "error" in verdict:
        raise SystemExit(f"授权/来源复核入口拒绝了这次输入：{verdict['error']}")
    return verdict



def plan_source_basis(
    served: list[dict[str, Any]] | None, source_change: dict[str, Any] | None
) -> str:
    """计划据以算出的输入，是**变更前**那一版还是**变更后**那一版。

    只从"桥实际交过什么值"推，不猜模型怎么算：

    - 没有任何一次成功的 `prepare_python_data`（交值）⇒ `no-read`（算不出来源）；
    - 没有来源变更记录 ⇒ `unrecorded`（不是本轮这个场景）；
    - 变更之后**再没有**成功的交值 ⇒ `old`：模型手里只有旧那一版数据，
      计划里的固定值只可能来自它（这正是 A 路的判据）；
    - 变更之后有过成功交值 ⇒ `new`：模型读到了新现场（B 路）。

    "被拒的调用"不算交值：`refusals` 里那几次只回了 `source_changed`，一个值都没给。
    """
    if not served:
        return "no-read"
    successful_reads = [
        item for item in served
        if item.get("tool") == "prepare_python_data" and not item.get("refused")
    ]
    if not successful_reads:
        return "no-read"
    if not source_change:
        return "unrecorded"
    changed_at = int(source_change.get("afterServedCall") or 0)
    later = [item for item in successful_reads if int(item.get("index") or 0) > changed_at]
    return "new" if later else "old"


def log(message: str) -> None:
    base.log(message)


def second_read_evidence(
    served: list[dict[str, Any]] | None, source_change: dict[str, Any] | None
) -> dict[str, Any]:
    """来源变化**之后**模型有没有再次调用受控读取工具，结果如何。

    "模型再次读取"只认一个事实：变更之后又出现了一次 `prepare_python_data` 调用
    （每次调用有自己的 `operationId`/`stepId`），并且它的回执是生产判据给的固定码。
    这里不做任何"它可能想读"的推断：没调用就是没调用。
    """
    if not served or not source_change:
        return {"branch": "none", "attempted": False, "attempts": []}
    changed_at = int(source_change.get("afterServedCall") or 0)
    attempts = [
        item for item in served
        if int(item.get("index") or 0) > changed_at and item.get("tool") == "prepare_python_data"
    ]
    if not attempts:
        # A 路：模型没有再次读取，只能靠写前复核拦下旧结果。
        return {"branch": "A", "attempted": False, "attempts": []}
    refused = all(bool(item.get("refused")) for item in attempts)
    first = attempts[0]
    first_served = next(
        (item for item in served
         if item.get("tool") == "prepare_python_data" and not item.get("refused")),
        None,
    )
    return {
        # B 路：模型确实再读了一次，而生产复核把它拒了。
        "branch": "B" if refused else "B-unrefused",
        "attempted": True,
        "attempts": [
            {
                "index": item.get("index"),
                "operationId": item.get("operationId"),
                "stepId": item.get("stepId"),
                "served": not item.get("refused"),
                "code": item.get("code"),
                "sources": item.get("sources"),
                "atSeconds": item.get("atSeconds"),
            }
            for item in attempts
        ],
        "refused": refused,
        "code": first.get("code"),
        "operationId": first.get("operationId"),
        "stepId": first.get("stepId"),
        # 同一来源范围：第二次读的表与第一次成功交值时完全相同 ⇒ 两次之间只有"数据变了"。
        "sameScope": first.get("sources") == (first_served or {}).get("sources"),
    }


def source_fingerprint() -> str:
    """按**当前**宿主替身内容算来源指纹。

    真实面板用 `wpsAuthorizedSourceFingerprint()`（同一批授权表的内容哈希）：
    绑定那一刻算一次，之后每次执行前**重算一次**再比对。这里保持同一语义，
    只是值由替身读出 —— 指纹本身仍是"这份内容的哈希"，不含业务值明文。
    """
    payload = [[sheet, base.slice_prepared_values(sheet, rect)] for sheet, rect in SOURCE_REGIONS]
    digest = hashlib.sha256(
        json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    ).hexdigest()
    return f"eb-src:{digest[:32]}"


# ---------------------------------------------------------------------------
# 开发故障入口（A1-E1）与结果引用探针
# ---------------------------------------------------------------------------


def fault_status(base_url: str) -> dict[str, Any]:
    return base.http_json(f"{base_url}/api/dev/result-ref-fault", None, timeout=15.0)


def arm_fault(base_url: str) -> dict[str, Any]:
    """显式启用一次。请求体只有宿主与工作簿 —— 与面板按钮同一条路径。"""
    status = base.http_json(
        f"{base_url}/api/dev/result-ref-fault/arm",
        {
            "hostId": HOST_ID,
            "workbookName": WORKBOOK_NAME,
            "workbookIdentity": WORKBOOK_IDENTITY,
        },
        timeout=15.0,
    )
    if status.get("phase") != "armed":
        raise SystemExit(f"启用后的状态不是 armed：{status}")
    return status


def probe_result(base_url: str, task_id: str, ref: str) -> dict[str, Any]:
    """按引用取回结果值。**故意用 GET 语义的只读端点**，不触发任何写入。"""
    body = {
        "taskId": task_id,
        "hostId": HOST_ID,
        "workbookName": WORKBOOK_NAME,
        "workbookIdentity": WORKBOOK_IDENTITY,
        "sourceFingerprint": None,
        "resultRef": ref,
    }
    try:
        payload = base.http_json(f"{base_url}/api/sdk/tool-bridge/result", body, timeout=30.0)
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        try:
            parsed = json.loads(detail)
        except ValueError:
            parsed = {"raw": detail[:400]}
        return {"ok": False, "status": error.code, "detail": parsed}
    return {
        "ok": True,
        "status": 200,
        "resultRef": payload.get("resultRef"),
        "headers": payload.get("headers"),
        "rows": payload.get("rows"),
        "rowCount": len(payload.get("rows") or []),
        "sourceSheets": payload.get("sourceSheets"),
    }


# ---------------------------------------------------------------------------
# 任务状态库（只读）：取回本次任务**按登记顺序**产生的全部结果引用
# ---------------------------------------------------------------------------


def task_state_db() -> Path:
    configured = os.getenv("EXCEL_BRO_CONFIG_DIR", "").strip()
    if configured:
        return Path(configured).resolve() / "tasks" / "state.sqlite3"
    local = os.getenv("LOCALAPPDATA", "").strip()
    if not local:
        raise SystemExit("Windows 环境缺少 LOCALAPPDATA，读不到任务状态库")
    return Path(local) / "Excel Bro" / "tasks" / "state.sqlite3"


def checkpoint_of(turn_id: str) -> dict[str, Any] | None:
    path = task_state_db()
    if not path.exists():
        return None
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=10.0)
    try:
        row = con.execute(
            "select checkpoint_json, status from planning_tasks where turn_id = ?",
            (turn_id,),
        ).fetchone()
    finally:
        con.close()
    if not row or not row[0]:
        return None
    try:
        checkpoint = json.loads(row[0])
    except ValueError:
        return None
    checkpoint["_taskStatus"] = row[1]
    return checkpoint


def registered_refs(turn_id: str) -> list[str]:
    """该任务在检查点里记下的结果引用（**按登记顺序**，与 `_note_step` 同源）。"""
    checkpoint = checkpoint_of(turn_id) or {}
    refs: list[str] = []
    for item in checkpoint.get("resultRefs") or []:
        ref = item.get("resultRef") if isinstance(item, dict) else None
        if isinstance(ref, str) and ref not in refs:
            refs.append(ref)
    return refs


def ledger_events(turn_id: str) -> list[dict[str, Any]]:
    """服务端**已落盘**的步骤序列（只取 stepId / status，不含代码与值）。

    为什么不靠 SSE 帧：流式帧在“暂停/中断/客户端先退出”时有丢失，而账本是服务端
    每一步 `_note_step` 之后写的 —— 它是"这条任务到底做了什么"的唯一持久事实。
    这里只读账本，不调模型、不发需求、不写任何东西。
    """
    path = task_state_db()
    if not path.exists():
        return []
    con = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=10.0)
    try:
        rows = con.execute(
            "select sequence, payload_json from planning_task_events where turn_id = ? order by sequence",
            (turn_id,),
        ).fetchall()
    finally:
        con.close()
    events: list[dict[str, Any]] = []
    for sequence, payload in rows:
        try:
            item = json.loads(payload)
        except ValueError:
            continue
        events.append({
            "sequence": sequence,
            "stepId": item.get("stepId"),
            "status": item.get("status"),
            "tool": (item.get("bridgeTool") or {}).get("tool"),
            "resultRef": (item.get("pythonResult") or {}).get("resultRef"),
        })
    return events


# ---------------------------------------------------------------------------
# 桥（面板等价）：读宿主替身 → 回执。可挂"第一次登记数据之后"的钩子。
# ---------------------------------------------------------------------------


def bridge_body(task_id: str, fingerprint: str) -> dict[str, Any]:
    """与面板 `bridgeBody()` 同形：身份 + 绑定那一刻的来源指纹。

    **必须**带 `workbookIdentity`：服务端按 `host + 名字 + 身份` 三者全等判归属
    （`BridgeBinding.matches`）。少了它就会回落到 `name:<名字>`，于是轮询永远返回
    空 `calls`、回执永远 `stale_binding` —— 表现是每个桥调用都按 `timeout` 收场，
    而界面上看不出原因（本轮就踩过：三次 `prepare_python_data` 全 warning）。
    """
    return {
        "taskId": task_id,
        "hostId": HOST_ID,
        "workbookName": WORKBOOK_NAME,
        "workbookIdentity": WORKBOOK_IDENTITY,
        "sourceFingerprint": fingerprint,
    }


def serve_bridge(
    base_url: str,
    task_id: str,
    *,
    fingerprint: str,
    stop: threading.Event | None = None,
    after_first_prepare: Any = None,
    idle_seconds: float = 240.0,
) -> dict[str, Any]:
    """面板等价循环：轮询 → 授权/来源复核（生产判据）→ 读宿主替身 → 回执。

    与真实面板的差别只有一处、且是**更严格**的方向：面板在复核不过时根本不执行读表；
    这里同样先复核，不通过就**一个值都不读**、只回固定错误码。复核本身来自共享核心
    `verifyBridgeAuthorization`，不是本文件的实现。
    """
    body = bridge_body(task_id, fingerprint)
    binding = {
        "taskId": task_id,
        "hostId": HOST_ID,
        "workbookName": WORKBOOK_NAME,
        "workbookIdentity": WORKBOOK_IDENTITY,
        "sourceFingerprint": fingerprint,
        "authorizedSheetNames": list(AUTHORIZED_SHEETS),
    }
    seen: list[str] = []
    served: list[dict[str, Any]] = []
    prepare_attempts: list[dict[str, Any]] = []
    reads_at_prepare: list[int] = []
    refusals: list[dict[str, Any]] = []
    started = time.time()
    idle_deadline = time.time() + idle_seconds
    while time.time() < idle_deadline and not (stop is not None and stop.is_set()):
        try:
            polled = base.http_json(f"{base_url}/api/sdk/tool-bridge/poll", body, timeout=30.0)
        except Exception:  # noqa: BLE001 - 桥是旁路，单次失败不该打断主轮
            time.sleep(1.0)
            continue
        calls = polled.get("calls") or []
        for call in calls:
            operation_id = call["operationId"]
            if operation_id in seen:
                continue
            seen.append(operation_id)
            tool = call.get("tool")
            args = call.get("arguments") or {}
            index = len(served) + 1
            step_id = f"bridge.{tool}:{str(operation_id).removeprefix('op_')[:12]}"
            log(f"桥调用：{tool} (#{index} {operation_id})")
            live = {
                "workbookIdentity": WORKBOOK_IDENTITY,
                # 现场指纹：每次调用前按**当前**宿主替身内容重算，与面板 `liveContext()` 同义。
                "sourceFingerprint": source_fingerprint(),
                "authorizedSheetNames": list(AUTHORIZED_SHEETS),
                "activeWorksheet": MEMBERS,
            }
            verdict = authorization_verdict(
                binding,
                {
                    "operationId": operation_id,
                    "tool": tool,
                    "arguments": args,
                    "requestFingerprint": call.get("requestFingerprint") or "",
                    "expiresAt": call.get("expiresAt") or 0,
                },
                live,
            )
            if not verdict.get("ok"):
                # 与面板同一条判据：授权/来源复核不过，就**不交值**，只回固定码。
                code = str(verdict.get("errorCode") or "internal")
                receipt = {"errorCode": code}
                refusals.append({
                    "index": index,
                    "tool": tool,
                    "code": code,
                    "operationId": operation_id,
                    "stepId": step_id,
                    "boundFingerprint": fingerprint,
                    "liveFingerprint": live["sourceFingerprint"],
                    "atSeconds": round(time.time() - started, 2),
                })
                log(f"  生产复核拒绝（{code}）：绑定 {fingerprint[:20]}… ≠ 现场 {live['sourceFingerprint'][:20]}…")
            else:
                try:
                    if tool == "query_table":
                        values = base.read_range(args.get("sheet", ""), args.get("range", "A1"))
                        receipt = {"result": {"headers": values[0], "rows": values[1:]}}
                    elif tool == "prepare_python_data":
                        receipt = {"datasets": base.build_datasets(args)}
                    else:
                        receipt = {"errorCode": "unsupported_tool"}
                except (ValueError, KeyError, TypeError) as error:
                    log(f"  桥执行被拒：{error}")
                    receipt = {"errorCode": "invalid_arguments"}
            base.http_json(
                f"{base_url}/api/sdk/tool-bridge/receipt",
                {
                    **body,
                    "operationId": operation_id,
                    "requestFingerprint": call.get("requestFingerprint"),
                    "receipt": receipt,
                },
                timeout=30.0,
            )
            refused = bool(receipt.get("errorCode"))
            served.append({
                "index": index,
                "operationId": operation_id,
                "stepId": step_id,
                "tool": tool,
                "sources": args.get("sources"),
                "refused": refused,
                "code": receipt.get("errorCode"),
                # 生产判据的裁决原文（`ok` / 固定错误码），不是本文件的推断。
                "gate": verdict,
                "boundFingerprint": fingerprint,
                "atSeconds": round(time.time() - started, 2),
            })
            if tool == "prepare_python_data":
                prepare_attempts.append({
                    "index": index,
                    "operationId": operation_id,
                    "stepId": step_id,
                    "served": not refused,
                    "code": receipt.get("errorCode"),
                    "sources": args.get("sources"),
                })
            if tool == "prepare_python_data" and not refused:
                # `len(served)` 就是这次调用在桥调用序列里的**序号**（1 起）。
                # 以前这里记的是"第几次成功 prepare 交值"（一个计数），而
                # `plan_source_basis` / `second_read_evidence` 拿它跟 `index` 比 ——
                # 只要前面还有一次 `query_table`，计数与序号就不相等，
                # "变更之后的交值"会被误判成变更之前（或反过来）。传序号。
                reads_at_prepare.append(len(served))
                if after_first_prepare is not None:
                    hook, after_first_prepare = after_first_prepare, None
                    hook(len(served))
            idle_deadline = time.time() + idle_seconds
        time.sleep(0.5)
    return {
        "served": served,
        "prepareCount": len(reads_at_prepare),
        "prepareAttempts": prepare_attempts,
        "refusals": refusals,
        "boundFingerprint": fingerprint,
        "finalFingerprint": source_fingerprint(),
    }


def watch_fault(
    base_url: str,
    stop: threading.Event,
    out: list[dict[str, Any]],
    on_consumed: Any = None,
) -> None:
    """记录故障状态的**时序**：armed → bound → consumed 是"确实在那一刻失效"的证据。

    `on_consumed` 在**第一次**看到 `consumed` 时回调一次 —— "用户在看到引用失效后
    按暂停"就发生在这一刻，早于/晚于它都不是那个场景。
    """
    started = time.time()
    last = None
    while not stop.is_set():
        try:
            status = fault_status(base_url)
        except Exception:  # noqa: BLE001
            time.sleep(0.5)
            continue
        phase = status.get("phase")
        if phase != last:
            out.append({"atSeconds": round(time.time() - started, 2), "phase": phase,
                        "taskBound": status.get("taskBound"), "refBound": status.get("refBound")})
            log(f"故障状态：{phase}（taskBound={status.get('taskBound')} refBound={status.get('refBound')}）")
            last = phase
        if phase == "consumed":
            if on_consumed is not None:
                callback, on_consumed = on_consumed, None
                callback()
            return
        time.sleep(0.3)


def pause_task(base_url: str, task_id: str) -> dict[str, Any]:
    return base.http_json(f"{base_url}/api/turn/{task_id}/pause", {}, timeout=30.0)


# ---------------------------------------------------------------------------
# 一次真实模型轮次
# ---------------------------------------------------------------------------


def build_request(turn_id: str, model_id: str, fingerprint: str, prompt: str) -> dict[str, Any]:
    return {
        "turnId": turn_id,
        "prompt": prompt,
        "modelId": model_id,
        "hostId": HOST_ID,
        "hostVersion": base.HOST_VERSION,
        "planningReadPolicy": "none",
        "selfReportedActions": base.SELF_REPORTED_ACTIONS,
        "workbook": {
            "name": WORKBOOK_NAME,
            "capturedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "activeWorksheet": MEMBERS,
            # 身份与来源指纹都是**绑定**的一部分：身份决定桥的归属（也是 A1-E1 故障
            # 绑定的匹配项），指纹决定"这一轮是在什么源版本上开始的"。
            "workbookIdentity": WORKBOOK_IDENTITY,
            "sourceFingerprint": fingerprint,
            "sourceFingerprintSheets": [MEMBERS, ORDERS],
            "worksheets": [
                {"name": MEMBERS, "usedRange": base.MEMBER_RANGE, "rowCount": 6,
                 "columnCount": 3, "headers": ["会员编号", "会员", "等级"]},
                {"name": ORDERS, "usedRange": base.ORDER_RANGE, "rowCount": 5,
                 "columnCount": 2, "headers": ["会员编号", "金额"]},
                {"name": TARGET, "usedRange": "A1", "rowCount": 1,
                 "columnCount": 1, "headers": []},
            ],
        },
    }


def run_turn(
    base_url: str, request: dict[str, Any], timeout: float
) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    response: dict[str, Any] = {}
    events: list[dict[str, Any]] = []
    extra: dict[str, Any] = {}
    try:
        for event, data in base.http_stream(f"{base_url}/api/turn/stream", request, timeout=timeout):
            if event in {"progress", "trace"} and isinstance(data, dict):
                events.append({
                    "stepId": data.get("stepId"),
                    "status": data.get("status"),
                    "summary": str(data.get("summary") or "")[:160],
                })
            elif event == "result" and isinstance(data, dict):
                response = data
            elif event == "paused" and isinstance(data, dict):
                extra["paused"] = data
                log(f"任务暂停：{data.get('pauseType')}")
            elif event == "error" and isinstance(data, dict):
                extra["streamError"] = data
                log(f"流式错误：{data.get('code')} {data.get('message')}")
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        extra["httpError"] = {"code": error.code, "detail": detail[:2000]}
        log(f"规划流失败 HTTP {error.code}")
    except Exception as error:  # noqa: BLE001 - 驱动要如实记录失败
        extra["error"] = f"{type(error).__name__}: {error}"
        log(f"规划流异常：{type(error).__name__}: {error}")
    return response, events, extra


# ---------------------------------------------------------------------------
# 场景
# ---------------------------------------------------------------------------


def plan_numbers(value: Any, out: set[float] | None = None) -> list[float]:
    """计划里出现过的数值（去重、排序）。

    只用来让"这份计划写的是旧那一版"在工件里可读：1000/400 是旧来源的汇总值，
    1150/550 才会是改后来源的汇总值。它**不是**判据本身 —— 判据是
    `plan_source_basis()`（桥到底交过哪一版值）。
    """
    sink = out if out is not None else set()
    if isinstance(value, bool):
        return sorted(sink)
    if isinstance(value, (int, float)):
        sink.add(float(value))
    elif isinstance(value, dict):
        for item in value.values():
            plan_numbers(item, sink)
    elif isinstance(value, list):
        for item in value:
            plan_numbers(item, sink)
    return sorted(sink)


def scenario(
    base_url: str,
    name: str,
    *,
    model_id: str,
    timeout: float,
    out: Path,
) -> dict[str, Any]:
    started = time.time()
    turn_id = f"turn-a1e2-{name}-{int(started)}"
    log(f"=== 场景 {name} turnId={turn_id} ===")
    # 绑定那一刻的来源指纹：与面板 `liveContext()` 同义 —— 之后每次执行前重算比对。
    fingerprint = source_fingerprint()
    log(f"绑定来源指纹 {fingerprint[:20]}…")
    prompt, prompt_label = prompt_for(name)

    # B 路要**单独隔离**"来源变化"这一个因素：不启用 A1-E1 的引用过期，
    # 于是旧结果引用仍然有效、计划仍可能被编译出来 —— 那样才谈得上
    # "迟到的旧计划会不会被写前复核拦下"。
    arms_fault = name in {"recovery", "source-change", "pause"}
    if arms_fault:
        armed = arm_fault(base_url)
        log(f"已启用一次性引用过期：{armed}")
    else:
        log("本场景不启用引用过期（只隔离来源变化这一个因素）")

    stop = threading.Event()
    timeline: list[dict[str, Any]] = []
    pause_sent: dict[str, Any] = {}
    source_change: dict[str, Any] | None = None

    def press_pause() -> None:
        """用户按下暂停：就在"引用刚失效、模型要重算"那一刻。"""
        try:
            pause_sent.update(pause_task(base_url, turn_id))
            log(f"已发送暂停请求：{pause_sent.get('status')}")
        except Exception as error:  # noqa: BLE001
            pause_sent["error"] = f"{type(error).__name__}: {error}"
            log(f"暂停请求失败：{error}")

    watcher = threading.Thread(
        target=watch_fault,
        args=(base_url, stop, timeline),
        kwargs={"on_consumed": press_pause if name == "pause" else None},
        daemon=True,
    )
    watcher.start()

    def after_first_prepare(index: int) -> None:
        """
        来源变化分支的**确定性屏障**：第一次 `prepare_python_data` 已经交值成功之后
        （模型手里已经是旧那一版数据），这里改一个来源单元格。

        不靠 sleep 抢时序、也不换来源表去逼模型重读：屏障只决定"什么时候改"，
        "模型要不要再读一次"完全交给模型自己。读到的现场与绑定指纹不同，
        之后任何一次桥调用都要先过生产复核 —— 过不了就只有 `source_changed`。
        """
        nonlocal source_change
        if name not in {"source-change", "source-change-reread"}:
            return
        before = base.SHEET_ROWS[ORDERS][1][1]
        before_fingerprint = source_fingerprint()
        base.SHEET_ROWS[ORDERS][1][1] = SOURCE_CHANGE_TO
        source_change = {
            # 改在"第几次桥调用之后"：这里是那次**成功交值**的桥调用序号（1 起）。
            "afterServedCall": index,
            "sheet": ORDERS,
            "cell": "B2",
            "before": before,
            "after": SOURCE_CHANGE_TO,
            "fingerprintBefore": before_fingerprint,
            "fingerprintAfter": source_fingerprint(),
        }
        log(
            f"来源已变更（首次读取完成之后）：{ORDERS}!B2 {before} → {SOURCE_CHANGE_TO}"
            f"（指纹 {before_fingerprint[:20]}… → {source_change['fingerprintAfter'][:20]}…）"
        )

    bridge_result: dict[str, Any] = {}
    bridge_thread = threading.Thread(
        target=lambda: bridge_result.update(
            serve_bridge(
                base_url,
                turn_id,
                fingerprint=fingerprint,
                stop=stop,
                after_first_prepare=after_first_prepare,
            )
        ),
        daemon=True,
    )
    bridge_thread.start()
    time.sleep(0.5)

    response, events, extra = run_turn(
        base_url, build_request(turn_id, model_id, fingerprint, prompt), timeout
    )
    # 先让桥收尾再写工件：否则 `bridge_result` 还是空的，工件里"桥做了什么"就变成 null
    # ——那种取证缺口最容易被误读成"没有桥调用"。
    stop.set()
    bridge_thread.join(timeout=10)

    elapsed = round(time.time() - started, 2)
    try:
        final_fault = fault_status(base_url)
    except Exception as error:  # noqa: BLE001 - 未启用故障入口时这个端点就是 403
        # B 路不需要故障入口；读不到状态只影响"故障"这一节的记录，不影响本场景结论。
        final_fault = {"enabled": None, "error": f"{type(error).__name__}: {error}"}
    refs = registered_refs(turn_id)
    checkpoint = checkpoint_of(turn_id)
    plan = response.get("plan") if isinstance(response.get("plan"), dict) else None
    plan_ref = plan.get("localResultRef") if plan else None
    log(f"规划返回 kind={response.get('kind')} 用时 {elapsed}s；登记引用 {len(refs)} 个")

    probes = []
    for index, ref in enumerate(refs):
        probe = probe_result(base_url, turn_id, ref)
        probes.append({
            "index": index,
            "resultRef": ref,
            "isPlanRef": ref == plan_ref,
            "ok": probe.get("ok"),
            "status": probe.get("status"),
            "rowCount": probe.get("rowCount"),
            "detailCode": (
                (probe.get("detail") or {}).get("detail", {}).get("code")
                if isinstance(probe.get("detail"), dict) else None
            ),
        })
        log(f"  引用 #{index + 1} {ref} → {'可读' if probe.get('ok') else probe.get('status')}")
    # 重复使用旧引用：必须仍然无效（不因缓存/重试复活）。
    # 只在**真的启用过**引用过期的轮次里验：B 路不启用故障，两个引用都还有效，
    # 那时去"再取一次"读到的是一份正常结果，不是"旧引用复活"。
    repeat_old = None
    if arms_fault and len(refs) >= 2:
        again = probe_result(base_url, turn_id, refs[0])
        repeat_old = {"ok": again.get("ok"), "status": again.get("status")}
        log(f"  旧引用再取一次 → {'可读（异常）' if again.get('ok') else again.get('status')}")

    actions = (plan or {}).get("actions") or []
    reread = second_read_evidence(bridge_result.get("served"), source_change)
    artifact = {
        "scenario": name,
        "turnId": turn_id,
        "modelId": model_id,
        # 这一轮用的**测试需求**（原样留档）：B 路靠的是"用户要求提交前再核对一次来源"，
        # 不是替模型编排工具顺序；工件里必须能看出用了哪一版需求。
        "promptLabel": prompt_label,
        "prompt": prompt,
        "workbook": {
            "name": WORKBOOK_NAME,
            "identity": WORKBOOK_IDENTITY,
            "hostId": HOST_ID,
        },
        "startedAt": started,
        "elapsedSeconds": elapsed,
        "fault": {
            "armed": arms_fault,
            "timeline": timeline,
            "final": final_fault,
        },
        "bridge": {
            "served": bridge_result.get("served"),
            "prepareCount": bridge_result.get("prepareCount"),
            "prepareAttempts": bridge_result.get("prepareAttempts"),
            "refusals": bridge_result.get("refusals"),
            "boundFingerprint": bridge_result.get("boundFingerprint"),
            "finalFingerprint": bridge_result.get("finalFingerprint"),
        },
        "registeredResultRefs": refs,
        # 来源变化分支的证据（其它场景为 null）：改的是哪一格、改在"第几次桥调用"
        # 之后、改动前后的指纹。`planSourceBasis` 只从"桥交过哪一版值"推出来。
        "sourceChange": source_change,
        # 变更之后模型有没有**真的再读一次**：`branch=B` 才算覆盖了 B 路；
        # `A` = 没再读（只能靠写前复核）；`B-unrefused` = 再读了但没被拦住（缺陷）。
        "secondRead": reread,
        "branch": reread["branch"],
        "planSourceBasis": plan_source_basis(bridge_result.get("served"), source_change),
        "planNumbers": plan_numbers(plan) if plan else [],
        # 检查点是"暂停/中断后可继续"的唯一凭据来源：只有 `checkpoint` 能说"继续"。
        # 这里只抄四个判据字段，不抄检查点正文（正文里有任务提示词等业务内容）。
        "checkpoint": {
            "stage": (checkpoint or {}).get("stage"),
            "resumability": (checkpoint or {}).get("resumability"),
            "taskStatus": (checkpoint or {}).get("_taskStatus"),
            "resultRefCount": len(refs),
        },
        "probes": probes,
        "repeatOldRef": repeat_old,
        "pauseRequest": pause_sent or None,
        "kind": response.get("kind"),
        "plan": plan,
        "planActionCount": len(actions),
        "planTargets": [
            {"sheet": action.get("sheet"), "range": action.get("range")}
            for action in actions
        ],
        # 规划期本驱动**从不写宿主**（桥只承载 query_table / prepare_python_data
        # 两个读工具，写回只能来自编译后的计划动作，且要用户确认）。所以这里恒为 0，
        # 它证明的是"旧/迟到结果被拒时没有任何写入"，不是"我数了 0 次"。
        "hostWritesDuringPlanning": 0,
        # 同一个计数器的显式别名：消费侧用例与链路校验都按这个键读"拒绝前后宿主写入数"。
        "hostWriteCalls": 0,
        # 步骤序列取自**服务端账本**（落盘事实）；SSE 帧只作旁证，缺帧不算没做事。
        "events": ledger_events(turn_id) or events,
        "streamEvents": events,
        **extra,
        "responseKeys": sorted(response.keys()),
    }
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(artifact, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"工件已写 {out}")
    return artifact


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", default="http://127.0.0.1:8791")
    parser.add_argument("--scenario", default="recovery",
                        choices=["recovery", "source-change", "source-change-reread",
                                 "pause", "no-fault"])
    parser.add_argument("--model", default=DEFAULT_MODEL_ID)
    parser.add_argument("--timeout", type=float, default=420.0)
    parser.add_argument("--out", default="")
    parser.add_argument("--host", default=base.HOST_ID,
                        help="宿主身份（wps-et / excel-office）；两端分别留证据时各跑一轮")
    parser.add_argument("--identity", default=base.WORKBOOK_IDENTITY,
                        help="工作簿身份（WPS = wps-path:…，Excel = url:…）")
    args = parser.parse_args()
    configure(host_id=args.host, identity=args.identity)
    log(f"宿主 {HOST_ID} · 工作簿身份 {WORKBOOK_IDENTITY}")

    if args.scenario in {"recovery", "source-change", "pause"} and not fault_status(args.base).get("enabled"):
        raise SystemExit("服务端没有开启 EXCEL_BRO_RESULT_REF_FAULT=1，拒绝用替身假装过期。")

    out = Path(args.out) if args.out else REPO / "outputs" / f"a1-e2-{args.scenario}.json"
    artifact = scenario(
        args.base, args.scenario, model_id=args.model, timeout=args.timeout, out=out
    )
    # 退出码只表达"驱动跑完了"，业务结论由消费侧用例与报告判定。
    log(
        f"场景 {args.scenario} 结束：kind={artifact['kind']} "
        f"refs={len(artifact['registeredResultRefs'])} actions={artifact['planActionCount']} "
        f"branch={artifact.get('branch')}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

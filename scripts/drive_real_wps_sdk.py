"""真实 SDK 后端的**面板等价驱动**：把 WPS 面板那次调用链在服务端外侧重放一遍。

用途：实机验收里"WPS 面板 + 真实 SDK 模型"这一组合，需要在**真实模型自主规划**下
跑完整条链（宿主读表 → 隔离容器 → 白名单计划 → 预览）。面板本身跑在 WPS 里，脚本
无法点击；本驱动逐字模仿 `apps/excel-addin/src/api.ts` 的请求体与
`useSdkToolBridge` 的轮询/回执循环，**唯一**替身是宿主侧读表（用真实 Docker 隔离
容器算，值矩阵由本脚本按真实语义读出并回执）。

明确**不做**的事：
- 不调用任何服务端内部函数绕过 `/api/turn`、桥轮询、回执、结果端点；
- 不预制结果矩阵、计划或异常说明（全部来自服务端真实执行）；
- 不替用户确认写回 —— 计划编译出来即停在预览处，写回由用户在面板上点。

用法：
    python scripts/drive_real_wps_sdk.py --out outputs/real-sdk-acceptance.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

REPO = Path(__file__).resolve().parents[1]

MEMBERS = "会员表"
ORDERS = "订单表"
TARGET = "汇总表"
MEMBER_RANGE = "A1:C6"
ORDER_RANGE = "A1:B5"
TARGET_RANGE = "A1:H8"

# 宿主侧的真实单元格内容（与面板里那份工作簿一致）。
MEMBER_ROWS: list[list[Any]] = [
    ["会员编号", "会员", "等级"],
    ["  7  ", "带空格会员", "金卡"],
    ["7", "重复键会员", "银卡"],
    ["007", "前导零会员", "银卡"],
    ["", "空键会员", "铜卡"],
    ["9", "正常会员", "金卡"],
]
ORDER_ROWS: list[list[Any]] = [
    ["会员编号", "金额"],
    ["7", 100],
    ["007", 200],
    ["9", 300],
    ["404", 400],
]

PROMPT = (
    "用隔离 Python 分析这两张已授权来源表：以订单表为左表，按会员编号左连接会员表。"
    "编号只去除首尾空格，保留前导零；空编号不参与匹配；会员表重复编号保留首次出现，"
    "后续重复记录单独说明。按关联得到的等级汇总订单金额，未匹配订单归入'未匹配'。"
    "展示汇总、金额总计，以及可定位的空编号、重复编号、未匹配明细。"
    "将汇总和异常明细安排到汇总表，使用定值矩阵，先展示写入预览，等我确认后再执行。"
    "不修改会员表、订单表，也不把原始数据行发送给模型。"
)

HOST_ID = "wps-et"
HOST_VERSION = "12.0"
_RANGE = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$")


def _column_index(column: str) -> int:
    value = 0
    for char in column:
        value = value * 26 + (ord(char) - 64)
    return value - 1


# 客户端自报集合（§4.3 收窄）：WPS 端在 `capabilities` 里自报为 supported 的动作。
# 不上报 ⇒ 服务端上界 ∩ 空集 = 空集 ⇒ `submit_plan` 的 `actions.items.anyOf` 变成
# `[]`，而网关会因此**整轮 400**（见下面 `_assert_tool_schemas_are_gateway_safe`）。
SELF_REPORTED_ACTIONS = ["writeValues", "writeFormulas", "clearRange:contents"]
WORKBOOK_NAME = "会员.xlsx"
WORKBOOK_IDENTITY = "wps-path:e:\\codex\\excel-bro\\outputs\\manual-acceptance\\会员.xlsx"


def http_json(url: str, payload: dict[str, Any] | None, timeout: float = 120.0):
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST" if data is not None else "GET",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        raw = response.read().decode("utf-8")
    return json.loads(raw) if raw.strip() else {}


def http_stream(url: str, payload: dict[str, Any], timeout: float):
    """读 SSE：逐帧产出 (event, data)。面板走的就是这条流式端点。"""
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "Accept": "text/event-stream"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        event = None
        for raw in response:
            line = raw.decode("utf-8").rstrip("\r\n")
            if line.startswith("event:"):
                event = line.split(":", 1)[1].strip()
            elif line.startswith("data:"):
                body = line.split(":", 1)[1].strip()
                try:
                    yield event, json.loads(body)
                except ValueError:
                    yield event, body
                event = None


def log(message: str) -> None:
    stamp = time.strftime("%H:%M:%S")
    print(f"[{stamp}] {message}", flush=True)


def bridge_body(task_id: str) -> dict[str, Any]:
    return {
        "taskId": task_id,
        "hostId": HOST_ID,
        "workbookName": WORKBOOK_NAME,
        "workbookIdentity": WORKBOOK_IDENTITY,
        "sourceFingerprint": None,
    }


SHEET_ROWS: dict[str, list[list[Any]]] = {
    MEMBERS: MEMBER_ROWS,
    ORDERS: ORDER_ROWS,
    TARGET: [[""]],
}


def read_range(sheet: str, rect: str) -> list[list[Any]]:
    """宿主替身：按矩形读出真实值矩阵（绝对坐标，与面板 `query_table` 同语义）。"""
    rows = SHEET_ROWS[sheet]
    columns = "ABCDEFGH".index(rect.split(":")[-1][0]) + 1
    return [list(row[:columns]) for row in rows]


def read_whole_sheet(sheet: str) -> list[list[Any]]:
    """读**整张已用区域**（`UsedRange` 的相对矩阵）。

    数据登记与确定性查询的区别就在这里：登记的 `range` 是 A1 **绝对**坐标，
    而宿主读回来的是**相对**矩阵。要按绝对区域切片，就必须先拿到
    `UsedRange` 的左上角（本替身里各表都从 A1 起，所以 origin = (1, 1)）。
    """
    return [list(row) for row in SHEET_ROWS[sheet]]


def slice_prepared_values(sheet: str, rect: str) -> list[list[Any]]:
    """按 A1 绝对区域切片，缺行/短行/越界**一律拒绝**，绝不补空白。

    与生产两端**同一份判据**（`@excel-bro/core/pythonPrepare` 的
    `slicePythonPrepareRegion`）：补空白会把"读了但没读到"变成"读到了空值"，
    那是伪造。
    """
    match = _RANGE.match(rect.strip().upper())
    if not match:
        raise ValueError("数据登记区域无效")
    c1 = _column_index(match.group(1))
    r1 = int(match.group(2)) - 1
    c2 = _column_index(match.group(3) or match.group(1))
    r2 = int(match.group(4) or match.group(2)) - 1
    if r2 < r1 or c2 < c1:
        raise ValueError("数据登记区域无效")
    # 本替身的 UsedRange 一律从 A1 起 ⇒ origin = (1, 1) ⇒ 矩阵下标 = 绝对下标。
    matrix = read_whole_sheet(sheet)
    if r2 >= len(matrix):
        raise ValueError("数据登记区域超出已读取范围")
    values: list[list[Any]] = []
    for row in matrix[r1 : r2 + 1]:
        if len(row) <= c2:
            raise ValueError("数据登记区域超出已读取范围")
        values.append(row[c1 : c2 + 1])
    return values


def build_datasets(args: dict[str, Any]) -> list[dict[str, Any]]:
    """`prepare_python_data` 回执体：逐块 `{sourceIndex, values}`。

    `sourceIndex` 按**声明顺序**给（0..n-1），与生产端
    `sliceWpsPythonSources` 一致 —— 服务端有一道完整性闸：声明 N 块就必须
    交回 `sourceIndex` 完整且互不重复的一集合，少一块直接判 `invalid_arguments`。
    """
    sources = args.get("sources")
    if not isinstance(sources, list) or not sources:
        raise ValueError("数据登记来源无效")
    blocks: list[dict[str, Any]] = []
    for source_index, source in enumerate(sources):
        sheet = source.get("sheet") if isinstance(source, dict) else None
        rect = source.get("range") if isinstance(source, dict) else None
        if not isinstance(sheet, str) or not isinstance(rect, str):
            raise ValueError("数据登记来源无效")
        blocks.append({"sourceIndex": source_index, "values": slice_prepared_values(sheet, rect)})
    return blocks


def serve_bridge(base: str, task_id: str) -> dict[str, Any]:
    """面板等价循环：轮询 → 在宿主上真实读表 → 回执。

    轮询节奏与面板一致（约 0.5 秒一次）。**不要空转紧循环** —— 每次 `urlopen` 都
    新建连接，Windows 的临时端口会被迅速耗尽（WinError 10048）。
    """
    body = bridge_body(task_id)
    seen: list[str] = []
    served: list[dict[str, Any]] = []
    idle_deadline = time.time() + 180
    while time.time() < idle_deadline:
        try:
            polled = http_json(f"{base}/api/sdk/tool-bridge/poll", body, timeout=30.0)
        except Exception:  # noqa: BLE001 - 桥是旁路，失败不该打断主轮
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
            log(f"桥调用：{tool} ({operation_id})")
            # 桥**只承载两个**工具（`BridgeToolName` = query_table | prepare_python_data）。
            # 其余一律 `unsupported_tool` —— 与生产端 `isBridgeToolName` 同一集合。
            try:
                if tool == "query_table":
                    values = read_range(args.get("sheet", ""), args.get("range", "A1"))
                    receipt = {"result": {"headers": values[0], "rows": values[1:]}}
                elif tool == "prepare_python_data":
                    receipt = {"datasets": build_datasets(args)}
                else:
                    receipt = {"errorCode": "unsupported_tool"}
            except (ValueError, KeyError, TypeError) as error:
                # 固定错误码，不把细节/值回传给模型 —— 与生产端一致。
                log(f"  桥执行被拒：{error}")
                receipt = {"errorCode": "invalid_arguments"}
            http_json(
                f"{base}/api/sdk/tool-bridge/receipt",
                {
                    **body,
                    "operationId": operation_id,
                    "requestFingerprint": call.get("requestFingerprint"),
                    "receipt": receipt,
                },
                timeout=30.0,
            )
            served.append({"tool": tool, "operationId": operation_id, "args": args})
            idle_deadline = time.time() + 180
        time.sleep(0.5 if calls else 0.5)
    return {"served": served}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", default="http://127.0.0.1:8765")
    parser.add_argument("--out", default="outputs/real-sdk-acceptance.json")
    parser.add_argument("--timeout", type=float, default=420.0)
    args = parser.parse_args()

    started = time.time()
    turn_id = f"turn-realsdk-{int(started)}"
    request = {
        "turnId": turn_id,
        "prompt": PROMPT,
        "modelId": "claude-agent-sdk:model-bad94777bbf6478c98cb0fa9025185a2",
        "hostId": HOST_ID,
        "hostVersion": HOST_VERSION,
        # SDK 后端是"表结构规划"：数据只经本机工具桥进来，HTTP 请求里不带值矩阵。
        "planningReadPolicy": "none",
        "selfReportedActions": SELF_REPORTED_ACTIONS,
        "workbook": {
            "name": WORKBOOK_NAME,
            "capturedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
            "activeWorksheet": MEMBERS,
            "worksheets": [
                {
                    "name": MEMBERS,
                    "usedRange": MEMBER_RANGE,
                    "rowCount": 6,
                    "columnCount": 3,
                    "headers": ["会员编号", "会员", "等级"],
                },
                {
                    "name": ORDERS,
                    "usedRange": ORDER_RANGE,
                    "rowCount": 5,
                    "columnCount": 2,
                    "headers": ["会员编号", "金额"],
                },
                {
                    "name": TARGET,
                    "usedRange": "A1",
                    "rowCount": 1,
                    "columnCount": 1,
                    "headers": [],
                },
            ],
        },
    }

    log(f"发出规划请求 turnId={turn_id}（Agent SDK · DS）")
    result: dict[str, Any] = {"turnId": turn_id, "startedAt": started}

    import threading

    bridge_thread = threading.Thread(
        target=lambda: result.setdefault("bridge", serve_bridge(args.base, turn_id)),
        daemon=True,
    )
    bridge_thread.start()

    # 面板的真实顺序：先意图判断（生成 turnId），再走流式规划端点。
    try:
        intent = http_json(
            f"{args.base}/api/intent/message",
            {
                "prompt": PROMPT,
                "modelId": request["modelId"],
                "imageCount": 0,
                "conversation": [],
            },
            timeout=120.0,
        )
        result["intent"] = {
            key: intent.get(key)
            for key in ("kind", "summary", "confirmedPrompt", "turnId")
        }
        log(f"意图判断 kind={intent.get('kind')}")
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        log(f"意图判断失败 HTTP {error.code}: {detail[:500]}")
        result["intentError"] = {"code": error.code, "detail": detail[:1200]}
        intent = {}

    events: list[dict[str, Any]] = []
    response: dict[str, Any] = {}
    try:
        for event, data in http_stream(
            f"{args.base}/api/turn/stream", request, timeout=args.timeout
        ):
            if event in {"progress", "trace"}:
                step = (data or {}).get("stepId") if isinstance(data, dict) else None
                if step:
                    events.append(
                        {
                            "stepId": step,
                            "status": (data or {}).get("status"),
                            "summary": str((data or {}).get("summary") or "")[:160],
                        }
                    )
            elif event == "result" and isinstance(data, dict):
                response = data
            elif event == "paused" and isinstance(data, dict):
                result["paused"] = data
                log(f"任务在预算/中断边界暂停：{data.get('pauseType')}")
            elif event == "error" and isinstance(data, dict):
                result["streamError"] = data
                log(f"流式错误：{data.get('code')} {data.get('message')}")
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        log(f"规划流失败 HTTP {error.code}: {detail[:800]}")
        result["httpError"] = {"code": error.code, "detail": detail[:2000]}
    except Exception as error:  # noqa: BLE001 - 驱动脚本要如实记录失败
        log(f"规划流异常：{type(error).__name__}: {error}")
        result["error"] = f"{type(error).__name__}: {error}"

    elapsed = time.time() - started
    result["elapsedSeconds"] = round(elapsed, 2)
    result["events"] = events
    result["response"] = response
    result["kind"] = response.get("kind")
    log(f"规划返回 kind={response.get('kind')} 用时 {elapsed:.1f} 秒（事件 {len(events)} 条）")
    for item in events:
        log(f"  step={item['stepId']} status={item['status']}")

    if response.get("kind") == "plan":
        plan = response.get("plan") or {}
        actions = plan.get("actions") or []
        result["planSummary"] = {
            "summary": plan.get("summary"),
            "actionCount": len(actions),
            "targets": [
                {"sheet": action.get("sheet"), "range": action.get("range")}
                for action in actions
            ],
            "localResultRef": plan.get("localResultRef"),
        }
        log(f"计划动作 {len(actions)} 个；进入预览（不代用户确认）")
    elif response.get("kind") == "answer":
        card = response.get("localResult") or {}
        result["answerCard"] = card
        log(
            "只拿到只读结果卡（未编译写入计划）"
            f" ref={card.get('resultRef')} message={str(response.get('message'))[:120]}"
        )
    elif response.get("kind") == "clarification":
        result["clarification"] = response.get("clarification")
        log(f"服务端要求澄清：{str(response.get('clarification'))[:200]}")

    bridge_thread.join(timeout=3)
    out = REPO / args.out
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log(f"已写 {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

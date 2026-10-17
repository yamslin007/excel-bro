"""跨前后端**同一次任务的数据流**集成测试（真实 ASGI 服务端 + 真实 Docker 隔离执行器）。

与既有三段测试的关系（用户点名要求 6：不得再用三个独立通过的测试合称完成）：

- ``test_python_continuous_chain_real.py``：只调**服务端内部函数**（不经 HTTP、
  不含端上面板）。
- ``apps/{excel,wps}-addin/src/*ContinuousPythonChain.test.tsx``：面板是真的，但
  **服务端的分析结果由 fetch mock 预制返回**（`EXPECTED_SUMMARY` 直接写在返回值里）。
- **本文件**把这两半接起来：一次真实调用里同时跑

    有状态宿主替身读表
      → **真实** ``POST /api/sdk/tool-bridge/poll``（服务端下发 prepare 指针）
      → 本机按指针真读宿主替身 → **真实** ``POST /api/sdk/tool-bridge/receipt``
      → 服务端 ``session.deliver`` → 协调器绑定 + 数据句柄登记
      → **真实** ``run_python_analysis``（`execute_python_analysis_async`）
      → **真实 Docker 隔离容器**执行清洗/关联/汇总
      → 结果登记进 ``local_results``
      → **真实** ``POST /api/sdk/tool-bridge/result`` 按引用取回矩阵
      → **真实** ``compile_result_plan`` 编译白名单 ``AnalysisPlan``

  端上的预览/确认/写回/回读/撤销由 TypeScript 侧用例覆盖（它们消费的就是本文件
  产出的**同一形状**：`BridgeResultResponse` 矩阵 + `AnalysisPlan`）。

**生产/消费闭环（本次新增）**：本文件把这次真实执行的产物导出到
`config/cross_process_chain.json`（结果引用 + 真实矩阵 + 异常明细 + 服务端编译的
`AnalysisPlan`），`apps/excel-addin/src/crossProcessChain.test.tsx` **读这份工件**
跑预览/确认/写回/回读/撤销 —— 端上拿到的矩阵与计划**不再**是测试里预制的常量。

**纪律（用户点名）**：
- 回执**不**由 mock 返回固定成功：请求真的走 ASGI 应用（``httpx.ASGITransport``），
  服务端真的处理；分析结果真的来自容器内执行，本文件**不注入任何预制矩阵**。
- 允许固定工具调用序列：本用例自己扮演"按服务端下发的指针干活的本机桥"。
- 期望常量（``EXPECTED_*``）**只用于断言**，不参与构造任何返回值。
- 因果验证：改的是**宿主替身里的会员表数据**（`7` 的等级 金卡 → 银卡），
  重跑同一条链路得到新结果；**不修改任何模拟响应**。

引擎/镜像不可用 ⇒ 整体 skip，如实记环境阻断，不用 mock 通过替代。
"""

from __future__ import annotations

import asyncio
import copy
import json
from pathlib import Path
from typing import Any

import pytest

from server.app.main import app
from server.app.models import PlanRequest
from server.app.result_plan import compile_result_plan

from server.tests.test_python_docker_real import _docker_available

pytestmark = pytest.mark.skipif(
    not _docker_available(),
    reason="本机无可用的 Docker Linux 容器引擎或隔离镜像（环境阻断，非测试通过）",
)

# 本用例产出的**真实工件**（服务端真实执行取回的矩阵 + 服务端真实编译的计划）。
# 端上用例（`apps/excel-addin/src/crossProcessChain.test.tsx`）**读这份工件**，
# 而不是在返回值里预制矩阵/计划 —— 这是"接成同一次任务数据流"的落点：
# 消费方拿到的结果引用、矩阵、异常明细、AnalysisPlan 全部来自这里的一次执行。
# 位置与既有跨语言共享用例一致（`config/*.json`，两端直接 import）。
FIXTURE_PATH = (
    Path(__file__).resolve().parents[2] / "config" / "cross_process_chain.json"
)

# ---------------------------------------------------------------------------
# 有状态宿主替身：两张来源表 + 一个空写入目标。改数据只能改这里。
# ---------------------------------------------------------------------------

MEMBERS = "会员表"
ORDERS = "订单表"
TARGET = "汇总表"

HOST_MEMBER_ROWS = [
    ["会员编号", "会员", "等级"],
    ["  7  ", "带空格会员", "金卡"],
    ["7", "重复键会员", "银卡"],
    ["007", "前导零会员", "银卡"],
    ["", "空键会员", "铜卡"],
    ["9", "正常会员", "金卡"],
]
HOST_ORDER_ROWS = [
    ["会员编号", "金额"],
    ["7", 100],
    ["007", 200],
    ["9", 300],
    ["404", 400],
]

# 手算期望（只用于断言）：
#   金卡  = 100（键 7 → 带空格那行，金卡）+ 300（键 9）  = 400
#   银卡  = 200（键 007）                                = 200
#   未匹配 = 400（键 404 在会员表查不到）                 = 400
EXPECTED_LEVEL_TOTALS_BASELINE = {"金卡": 400, "银卡": 200, "未匹配": 400}
EXPECTED_TOTAL_BASELINE = 1000

# 因果验证：把会员表**第一次出现的键 7** 的等级从金卡改成银卡。
#   金卡  = 300（键 9）                                  = 300
#   银卡  = 100（键 7）+ 200（键 007）                    = 300
#   未匹配 = 400（键 404）                                = 400
EXPECTED_LEVEL_TOTALS_AFTER_EDIT = {"金卡": 300, "银卡": 300, "未匹配": 400}
EXPECTED_TOTAL_AFTER_EDIT = 1000

WORKBOOK_NAME = "会员.xlsx"
WORKBOOK_IDENTITY = "url:https://contoso.sharepoint.com/membership.xlsx"
SOURCE_FINGERPRINT = "fp-source-1"
TASK_ID = "cross-process-chain-task"

# 容器内分析代码（"模型生成分析逻辑"的等价物）。自己读输入、自己清洗、自己
# 关联、自己汇总，**不含任何预制结果**。
ANALYSIS_CODE = r"""
import json

members = input["datasets"][0]["values"]
orders = input["datasets"][1]["values"]

m_header, *m_body = members
m_key, m_level = m_header.index("会员编号"), m_header.index("等级")
clean = {}
empty_keys, duplicate_keys = [], []
for offset, row in enumerate(m_body):
    raw = "" if row[m_key] is None else str(row[m_key])
    key = raw.strip()
    if key == "":
        empty_keys.append(offset + 2)
        continue
    if key in clean:
        duplicate_keys.append(key)
        continue
    clean[key] = row[m_level]

o_header, *o_body = orders
o_key, o_amount = o_header.index("会员编号"), o_header.index("金额")
totals = {}
unmatched = []
for row in o_body:
    raw = "" if row[o_key] is None else str(row[o_key])
    key = raw.strip()
    if key == "":
        continue
    amount = 0 if row[o_amount] is None else row[o_amount]
    if key in clean:
        level = clean[key]
    else:
        level = "未匹配"
        unmatched.append(key)
    totals[level] = totals.get(level, 0) + amount

rows = sorted([[level, total] for level, total in totals.items()])
detail_rows = [
    ["空键数", len(empty_keys), ""],
    ["重复键数", len(duplicate_keys), ""],
    ["未匹配键", len(unmatched), ",".join(sorted(unmatched))],
]
result = {
    "kind": "matrix",
    "values": [["会员等级", "金额合计", ""]] + [r + [""] for r in rows] + detail_rows,
}
"""


class HostDouble:
    """有状态宿主替身：`read`/`write` 落在同一份矩阵上。"""

    def __init__(self, sheets: dict[str, list[list[Any]]]) -> None:
        self._sheets = {name: copy.deepcopy(rows) for name, rows in sheets.items()}

    def read(self, sheet: str, row_start: int, row_end: int, col_start: int, col_end: int):
        grid = self._sheets[sheet]
        out: list[list[Any]] = []
        for r in range(row_start, row_end + 1):
            row = grid[r - 1] if 0 <= r - 1 < len(grid) else []
            out.append(
                [
                    row[c - 1] if 0 <= c - 1 < len(row) else None
                    for c in range(col_start, col_end + 1)
                ]
            )
        return out

    def set_level(self, sheet: str, key: str, level: str) -> None:
        """只改宿主数据：找到该键**第一次出现**的行，改它的等级列。"""
        grid = self._sheets[sheet]
        header = grid[0]
        key_column = header.index("会员编号")
        level_column = header.index("等级")
        for row in grid[1:]:
            if str(row[key_column]).strip() == key:
                row[level_column] = level
                return
        raise AssertionError(f"宿主替身里找不到键 {key!r}")

    def matrix(self, sheet: str) -> list[list[Any]]:
        return copy.deepcopy(self._sheets[sheet])


def _local_client():
    import httpx

    return httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app), base_url="http://localhost"
    )


def _bridge_body(**extra: Any) -> dict[str, Any]:
    return {
        "taskId": TASK_ID,
        "hostId": "excel-office",
        "workbookName": WORKBOOK_NAME,
        "workbookIdentity": WORKBOOK_IDENTITY,
        "sourceFingerprint": SOURCE_FINGERPRINT,
        **extra,
    }


def _plan_request(result_rows: int = 8, result_columns: int = 3) -> PlanRequest:
    """面板规划请求（形状与端上一致；写入目标唯一授权表 = 汇总表）。

    `usedRange` 必须覆盖真实结果的形状：`compile_result_plan` 用
    `result_matrix(item, include_header=True)` = **服务端通用表头行 + 结果行**
    与目标区域逐项核对行列数，不一致就拒绝（`size_mismatch`）—— 这是产品闸，
    不是测试便利。本用例的矩阵 = 1 行服务端表头 + 7 行分析产出 = 8 行 × 3 列。
    """
    return PlanRequest.model_validate(
        {
            "hostId": "excel-office",
            "prompt": "按会员等级汇总订单金额，结果写到汇总表",
            "workbook": {
                "name": WORKBOOK_NAME,
                "capturedAt": "2026-10-10T00:00:00Z",
                "workbookIdentity": WORKBOOK_IDENTITY,
                "activeWorksheet": TARGET,
                "sourceFingerprint": SOURCE_FINGERPRINT,
                "worksheets": [
                    {
                        "name": TARGET,
                        "sourceSheet": TARGET,
                        "usedRange": _a1(result_rows, result_columns),
                        "rowCount": 0,
                        "columnCount": 0,
                        "headers": [],
                        "dataRows": [],
                        "truncated": False,
                    }
                ],
            },
        }
    )


def _column_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _a1(rows: int, columns: int) -> str:
    return f"A1:{_column_letter(columns)}{rows}"


def _parse_range(range_text: str) -> tuple[int, int, int, int]:
    start, end = range_text.split(":")
    col_start, row_start = _split_cell(start)
    col_end, row_end = _split_cell(end)
    return row_start, col_start, row_end, col_end


def _split_cell(cell: str) -> tuple[int, int]:
    letters = "".join(ch for ch in cell if ch.isalpha())
    digits = "".join(ch for ch in cell if ch.isdigit())
    column = 0
    for ch in letters:
        column = column * 26 + (ord(ch.upper()) - ord("A") + 1)
    return column, int(digits)


class _Turn:
    """最小 request 对象：`turnId` 就是任务 id（协调器按它查绑定）。"""

    def __init__(self, turn_id: str) -> None:
        self.turnId = turn_id


async def _run_cross_process_chain(host: HostDouble) -> dict[str, Any]:
    """一次连续调用：桥读表 → 服务端真实分析（真实容器）→ 取回 → 编译计划。

    只**开功能开关**（`python_execution_enabled`，与仓库既有真实用例同一做法）；
    `run_isolated_analysis_internal` / `run_registered_analysis_internal` /
    `execute_python_analysis_async` 一律走真实实现，**不**被顶替。
    """
    from server.app import local_results, python_analysis, python_coordinator, task_bridge

    # 干净起点：清掉可能残留的会话/绑定/归属，避免跨用例串台。
    task_bridge.task_tool_bridge.close(TASK_ID)
    task_bridge.task_tool_bridge.release(TASK_ID)
    local_results.local_results.discard_task(TASK_ID)
    python_coordinator.coordinator.release(TASK_ID)

    plan_request = _plan_request()
    binding = task_bridge.BridgeBinding.from_request(plan_request, task_id=TASK_ID)
    session = task_bridge.task_tool_bridge.open(binding)

    original_enabled = python_analysis.python_execution_enabled
    python_analysis.python_execution_enabled = lambda: True
    try:
        async with _local_client() as client:
            # ① 服务端（真实 handler）下发 prepare_python_data 指针。
            request_task = asyncio.create_task(
                session.request(
                    "prepare_python_data",
                    {
                        "sources": [
                            {"sheet": MEMBERS, "range": "A1:C6"},
                            {"sheet": ORDERS, "range": "A1:B5"},
                        ]
                    },
                )
            )
            for _ in range(1000):
                if session.pending_calls:
                    break
                await asyncio.sleep(0)

            # ② 本机桥：真实 HTTP 轮询拿到指针。
            poll = await client.post("/api/sdk/tool-bridge/poll", json=_bridge_body())
            assert poll.status_code == 200, poll.text
            calls = poll.json()["calls"]
            assert calls, "服务端没有下发任何桥指针"
            call = calls[0]
            assert call["tool"] == "prepare_python_data", call

            # ③ 本机桥：按指针**真读宿主替身**，拼回执。
            blocks = []
            for source in call["arguments"]["sources"]:
                row_start, col_start, row_end, col_end = _parse_range(source["range"])
                blocks.append(
                    {
                        "sourceIndex": len(blocks),
                        "values": host.read(
                            source["sheet"], row_start, row_end, col_start, col_end
                        ),
                    }
                )

            # ④ 回执交给**真实服务端**：session.deliver → 协调器绑定 + 句柄登记。
            receipt = await client.post(
                "/api/sdk/tool-bridge/receipt",
                json=_bridge_body(
                    operationId=call["operationId"],
                    requestFingerprint=call["requestFingerprint"],
                    receipt={"datasets": blocks},
                ),
            )
            assert receipt.status_code == 200, receipt.text
            assert receipt.json() == {"status": "accepted", "code": None}, receipt.text
            # 服务端回执：`{status:"ok", datasetHandles:[...]}`。句柄是数据**引用**，
            # 值本身留在服务端易失登记里（`dataset://` 不是授权凭据）。
            bridge_receipt = await request_task
            assert bridge_receipt["status"] == "ok", bridge_receipt
            handles = bridge_receipt["datasetHandles"]
            assert handles, "服务端没有从回执里得到数据句柄"
            assert all(str(h).startswith("dataset://") for h in handles), handles

            # ⑤ 真实 `run_python_analysis`：句柄 → **真实 Docker 隔离执行器**。
            response = await python_coordinator.execute_python_analysis_async(
                _Turn(TASK_ID), {"code": ANALYSIS_CODE, "datasetHandles": handles}
            )
            assert response["status"] == "completed", response
            # `run_python_analysis` 登记进 `local_results`，句柄是 `res_<32hex>`
            # （端上面板按它走 `/api/sdk/tool-bridge/result` 取回）。
            result_ref = response["resultHandle"]
            assert result_ref.startswith("res_"), result_ref

            # ⑥ 端上按引用取回：真实 HTTP 端点（local_results.resolve_owned）。
            fetched = await client.post(
                "/api/sdk/tool-bridge/result", json=_bridge_body(resultRef=result_ref)
            )
            assert fetched.status_code == 200, fetched.text
            assert fetched.headers.get("cache-control") == "no-store"
            payload = fetched.json()
    finally:
        python_analysis.python_execution_enabled = original_enabled

    # ⑦ 白名单计划编译：真实 `compile_result_plan`（不预制 AnalysisPlan）。
    item = local_results.local_results.resolve(TASK_ID, result_ref)
    assert item is not None, "结果未登记进 local_results"
    plan = compile_result_plan(
        item, plan_request, target_sheet=TARGET, target_range="A1:C8"
    )

    return {
        "handles": handles,
        "resultRef": result_ref,
        "payload": payload,
        "plan": plan,
    }


def _export_fixture(result: dict[str, Any]) -> Path:
    """把**本次真实执行**的产物落成端上可读的工件（不加工、不改写）。

    落点 `config/cross_process_chain.json`：与既有跨语言共享用例
    （`python_prepare_cases.json` / `mapping_slice_cases.json`）同一处，
    两端直接 import。
    工件里只有两类东西，都是服务端给的：
    - `result`：`POST /api/sdk/tool-bridge/result` 的真实响应体
      （`resultRef` / `headers` / `rows` / 异常明细在 `rows` 里 / `sourceSheets` …）；
    - `plan`：`compile_result_plan` 真实编译出的白名单 `AnalysisPlan`
      （含 `localResultRef` / `sourceDependencySheets` / 写入坐标）。
    """
    FIXTURE_PATH.parent.mkdir(parents=True, exist_ok=True)
    FIXTURE_PATH.write_text(
        json.dumps(
            {
                "workbookName": WORKBOOK_NAME,
                "workbookIdentity": WORKBOOK_IDENTITY,
                "sourceFingerprint": SOURCE_FINGERPRINT,
                "hostId": "excel-office",
                "memberSheet": MEMBERS,
                "orderSheet": ORDERS,
                "targetSheet": TARGET,
                "resultRef": result["resultRef"],
                "result": result["payload"],
                "plan": result["plan"].model_dump(mode="json"),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return FIXTURE_PATH


def _summary_of(payload: dict[str, Any]) -> dict[str, Any]:
    """等级汇总：矩阵 `rows[1:4]`（`rows[0]` 是分析代码产出的表头行）。

    注意 `BridgeResultResponse.headers` 是服务端给的通用列名（`列1..列N`），
    **不是**分析结果的自有表头 —— 自有表头就在 `rows[0]` 里。
    """
    return {row[0]: row[1] for row in payload["rows"][1:4]}


def _detail_of(payload: dict[str, Any]) -> dict[str, Any]:
    return {row[0]: row[1:] for row in payload["rows"][4:]}


def _matrix_of(payload: dict[str, Any]) -> list[list[Any]]:
    """端上矩阵 = 服务端返回的全部行（含分析代码产出的表头行）。

    空白先折叠：计划值经过**编译层**的 ``_blank_cell_of_write``，把分析结果里的
    ``""`` 折成 ``None``（见 ``test_result_plan_blank_folding.py``）。只读回执里
    仍是 ``""`` —— 同一空白的两种写法，必须归一后再比，否则会把正确的折叠误判成
    不一致。
    """
    return [
        [None if cell == "" else cell for cell in row] for row in payload["rows"]
    ]


def test_cross_process_chain_exports_artifact_for_panel_consumers():
    """把**本次真实执行**的产物落成端上可读的工件（供面板用例消费）。

    这是"接成同一次任务数据流"的**生产侧**：工件里的结果矩阵、异常明细、
    结果引用、白名单计划，全部来自上面那条真实链路（真实容器 + 真实编译），
    **不是**测试预制的常量。消费侧见 `apps/excel-addin/src/crossProcessChain.test.tsx`
    （它读本工件而不是在返回值里预制矩阵/计划）。
    """
    host = HostDouble({MEMBERS: HOST_MEMBER_ROWS, ORDERS: HOST_ORDER_ROWS, TARGET: []})
    result = asyncio.run(_run_cross_process_chain(host))
    fixture_path = _export_fixture(result)

    on_disk = json.loads(fixture_path.read_text(encoding="utf-8"))
    # 工件内容 == 这次执行的真实产物（逐项对拍，防止写歪）。
    assert on_disk["resultRef"] == result["resultRef"]
    assert on_disk["result"] == result["payload"]
    assert on_disk["plan"] == result["plan"].model_dump(mode="json")
    # 工件里的汇总与异常明细就是这次执行的结果。
    assert _summary_of(on_disk["result"]) == EXPECTED_LEVEL_TOTALS_BASELINE
    assert _detail_of(on_disk["result"])["未匹配键"][1] == "404"


def test_cross_process_chain_is_one_task_not_three():
    """一条真实连续调用：桥回执 → 服务端真实处理 → 真实容器 → 取回 → 编译计划。

    覆盖用户点名的 1/2/3：**回执交给真实服务端**（走 ASGI 应用，不是 fetch mock
    的固定成功）；`prepare_python_data` → `run_python_analysis` → 结果计划提交
    全部走产品的 handler/bridge/coordinator/隔离执行器/计划编译路径；
    面板收到的矩阵与异常明细**只能来自这次服务端执行**。
    """
    host = HostDouble({MEMBERS: HOST_MEMBER_ROWS, ORDERS: HOST_ORDER_ROWS, TARGET: []})
    raw_members = host.matrix(MEMBERS)
    raw_orders = host.matrix(ORDERS)

    result = asyncio.run(_run_cross_process_chain(host))
    payload = result["payload"]

    # ---- 分析结果确实来自服务端真实执行（不是 mock 的固定成功）----------
    assert payload["resultRef"] == result["resultRef"]
    assert payload["complete"] is True
    assert payload["calculation"]          # 计算说明来自本次执行
    assert payload["sourceSheets"] == [MEMBERS, ORDERS]
    # 端上拿到的是完整矩阵：1 行表头 + 3 汇总 + 3 明细 = 7 行 × 3 列。
    assert len(payload["headers"]) == 3
    assert len(payload["rows"]) == 7
    # ---- 汇总与异常明细逐项与手算一致（期望常量只用于断言）-------------
    summary = _summary_of(payload)
    assert summary == EXPECTED_LEVEL_TOTALS_BASELINE, payload["rows"]
    assert sum(summary.values()) == EXPECTED_TOTAL_BASELINE

    detail = _detail_of(payload)
    assert detail["空键数"][0] == 1
    assert detail["重复键数"][0] == 1
    assert detail["未匹配键"][0] == 1
    assert detail["未匹配键"][1] == "404"   # 可定位到具体键，不是"有异常"

    # 来源两表在整条链路里逐格未变（读表不产生副作用）。
    assert host.matrix(MEMBERS) == raw_members
    assert host.matrix(ORDERS) == raw_orders

    # ---- 白名单计划由**服务端**编译，矩阵来自这次执行 -------------------
    plan = result["plan"]
    assert len(plan.actions) == 1
    action = plan.actions[0]
    assert action.type == "writeValues"
    assert action.sheet == TARGET
    # 编译矩阵 = 1 行服务端表头 + 7 行分析产出 = 8 行 × 3 列。
    assert action.range == "A1:C8"
    # 计划矩阵的首行是服务端通用表头，其余 == 端上取回的全部行。
    assert action.values[0] == list(payload["headers"])
    assert action.values[1:] == _matrix_of(payload)
    assert plan.localResultRef == result["resultRef"]
    # 实际来源依赖随计划一起走（不含写入目标）。
    assert list(plan.sourceDependencySheets or []) == [MEMBERS, ORDERS]


def test_cross_process_chain_result_is_authorization_bound():
    """同一条链路的授权绑定：换授权后按引用取回必须被拒（不发布陈旧结果）。"""
    from server.app import local_results

    host = HostDouble({MEMBERS: HOST_MEMBER_ROWS, ORDERS: HOST_ORDER_ROWS, TARGET: []})
    result = asyncio.run(_run_cross_process_chain(host))
    result_ref = result["resultRef"]

    async def fetch_after_revoke():
        async with _local_client() as client:
            # 换一次授权（授权收窄/换簿）→ 旧结果与旧授权不再相配。
            local_results.local_results.bind(
                task_id=TASK_ID,
                host_id="excel-office",
                workbook_name=WORKBOOK_NAME,
                workbook_identity=WORKBOOK_IDENTITY,
                authorized_key="different-authorized-scope",
                source_fingerprint=SOURCE_FINGERPRINT,
            )
            return await client.post(
                "/api/sdk/tool-bridge/result", json=_bridge_body(resultRef=result_ref)
            )

    response = asyncio.run(fetch_after_revoke())
    assert response.status_code == 404, response.text
    assert response.json()["detail"]["code"] == "BRIDGE_RESULT_UNAVAILABLE"


def test_member_level_edit_changes_result_without_touching_mocks():
    """因果验证：改宿主里的会员表 → 重跑同一条链路 → 汇总随之变化。

    用户点名：把**第一次出现的键 7** 的等级从金卡改成银卡，重新执行后应为
    金卡 300、银卡 300、未匹配 400。**不修改任何模拟响应** —— 改的只有宿主
    替身里的单元格数据，其它一切（指针、端点、执行器、编译）原样复用。
    """
    host = HostDouble({MEMBERS: HOST_MEMBER_ROWS, ORDERS: HOST_ORDER_ROWS, TARGET: []})

    baseline = asyncio.run(_run_cross_process_chain(host))
    assert _summary_of(baseline["payload"]) == EXPECTED_LEVEL_TOTALS_BASELINE

    # **唯一**的改动：宿主数据。键 7 第一次出现在第 2 行（带空格会员）。
    host.set_level(MEMBERS, "7", "银卡")

    edited = asyncio.run(_run_cross_process_chain(host))
    summary = _summary_of(edited["payload"])

    # 结果真的跟着数据变了（金卡 400→300，银卡 200→300）。
    assert summary == EXPECTED_LEVEL_TOTALS_AFTER_EDIT, edited["payload"]["rows"]
    assert summary != EXPECTED_LEVEL_TOTALS_BASELINE
    assert sum(summary.values()) == EXPECTED_TOTAL_AFTER_EDIT
    # 异常明细不受等级改动影响（键集合没变）。
    assert _detail_of(edited["payload"])["未匹配键"][1] == "404"

    # 计划里的矩阵同样跟着变 —— 服务端重新编译，不是复用上一份。
    plan_values = edited["plan"].actions[0].values
    # 比"语义"而不是比"字面量"：计划值经过**编译层**的空白折叠
    # （``_blank_cell_of_write`` 把分析结果里的 ``""`` 折成 ``None``），所以明细行的
    # `未匹配成员` 一列在计划里是 ``None``、在只读回执里是 ``""`` —— 同一个空白的
    # 两种写法，归一后再比才不会被这次正确的折叠误判成不一致。
    assert plan_values[1] == _matrix_of(edited["payload"])[0]
    plan_totals = {row[0]: row[1] for row in plan_values[2:5]}
    assert plan_totals == EXPECTED_LEVEL_TOTALS_AFTER_EDIT
    assert edited["plan"].actions[0].values != baseline["plan"].actions[0].values

"""同一次任务的**连续链路**实测（真实 Docker 隔离执行器；引擎/镜像不可用时整体 skip）。

覆盖用户点名的连续链路（**一条**用例串完，不是分段拼装）：

    宿主读取 → prepare_python_data（数据句柄登记）→ run_python_analysis
    → 已验证的隔离执行器（真实容器）→ 结果登记与引用取回
    → （白名单计划编译与面板预览/确认/写回在两端 TypeScript 侧另测）

与 ``test_python_docker_real.py`` 的分工：
- 该文件验证**隔离机制本身**（网络/文件/凭据/限额/清理）；
- 本文件验证**产品逻辑的连续串接**：脏数据在真实容器里被真正清洗/关联/汇总，
  结果经登记、按引用取回后与**手算期望逐项一致**。

纪律（与仓库既有口径一致）：
- **不预制结果矩阵、不预制异常明细** —— 期望值在文件里手算写出，容器里跑的
  清洗/关联/汇总代码是本次测试提供的**分析代码**（等价于模型生成的分析逻辑），
  结果**只能**来自真实执行。
- 不得用 ``monkeypatch`` 顶替 ``run_isolated_analysis_internal`` / 
  ``run_registered_analysis_internal``；这两个必须走真实路径。
- 引擎不可用 ⇒ 整体 skip，如实记为环境阻断，**不用 mock 通过替代**。
"""
from __future__ import annotations

import json

import pytest

from server.app import python_analysis
from server.app.python_datasets import DatasetRejected, PythonDatasetRegistry
from server.app.models import PythonDataScope, PythonResultReceipt

from server.tests.test_python_docker_real import _docker_available


pytestmark = pytest.mark.skipif(
    not _docker_available(),
    reason="本机无可用的 Docker Linux 容器引擎或隔离镜像（环境阻断，非测试通过）",
)

ORIGIN = "https://localhost:3000"

# ---------------------------------------------------------------------------
# 脏数据（与前端合成案例 `syntheticResultEndToEnd.test.tsx` / 
# `WpsAgentPanel.test.tsx` **同一批**，两端同一套期望值）
# ---------------------------------------------------------------------------

MEMBERS = [
    ["会员编号", "会员", "等级"],
    ["  7  ", "带空格会员", "金卡"],
    ["7", "重复键会员", "银卡"],
    ["007", "前导零会员", "银卡"],
    ["", "空键会员", "铜卡"],
    ["9", "正常会员", "金卡"],
]
ORDERS = [
    ["会员编号", "金额"],
    ["7", 100],
    ["007", 200],
    ["9", 300],
    ["404", 400],
]

# 手算期望（独立于容器内代码）：
#   键 trim 后比对；保留前导零（`007` ≠ `7`）；空键排除；重复键保留第一条；
#   订单左连接会员表，按**会员等级**汇总金额：
#     金卡  = 100（7，关联到带空格行的金卡）+ 300（9）          = 400
#     银卡  = 200（007）                                       = 200
#     未匹配 = 400（404 在会员表里查不到）                      = 400
#     总计  = 1000
EXPECTED_LEVEL_TOTALS = {"金卡": 400, "银卡": 200, "未匹配": 400}
EXPECTED_TOTAL = 1000

# 容器内执行的分析代码（模型生成分析逻辑的等价物）。**不使用任何预制结果**：
# 它自己读输入、自己清洗、自己关联、自己汇总，并把明细一并输出。
ANALYSIS_CODE = r"""
import json

members = input["datasets"][0]["values"]
orders = input["datasets"][1]["values"]

# --- 清洗会员表：trim 比对、保留前导零、空键排除、重复键保留第一条 ---
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

# --- 订单左连接会员表，按等级汇总 ---
o_header, *o_body = orders
o_key, o_amount = o_header.index("会员编号"), o_header.index("金额")
totals = {}
unmatched = []
for offset, row in enumerate(o_body):
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
    entry = totals.setdefault(level, 0)
    totals[level] = entry + amount

rows = sorted([[level, total] for level, total in totals.items()])
# 结果**必须**恰好是 {"kind", "values"} 两键（`encode_matrix_result` 会拒绝额外的
# 字段）。工作程序只回一个 `result` 变量，所以把汇总与**异常明细**拼进同一张
# 定宽矩阵：前几行是等级汇总，末尾追加三行可定位明细。全部来自本次真实执行，
# **没有任何预制结果**参与。
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



def _scope():
    return PythonDataScope.model_validate(
        dict(
            version=1,
            taskId="continuous-chain-task",
            hostId="excel-office",
            workbookIdentityHash="a" * 64,
            authorizationVersion="b" * 64,
            sources=[
                dict(sheetId="sheet-1", sheetName="会员表", range="A1:C6"),
                dict(sheetId="sheet-2", sheetName="订单表", range="A1:B5"),
            ],
        )
    )


def test_continuous_chain_dirty_data_to_published_result():
    """一条连续链路：登记 → 真实容器执行 → 结果登记 → 按引用取回 → 与手算一致。"""
    registry = PythonDatasetRegistry()
    target = _scope()

    # 1) prepare_python_data：可信侧登记两张来源表（回执 = 句柄 + 指纹）。
    token = registry.create(target, ORIGIN)
    members_receipt = registry.register(token, target, ORIGIN, 0, MEMBERS)
    orders_receipt = registry.register(token, target, ORIGIN, 1, ORDERS)
    receipts = [members_receipt, orders_receipt]

    # 2) + 3) run_python_analysis → 已验证的隔离执行器（**真实容器**，非 mock）。
    response = python_analysis.run_registered_result_internal(
        data_registry=registry,
        scope_token=token,
        scope=target,
        origin=ORIGIN,
        receipts=receipts,
        code=ANALYSIS_CODE,
    )
    assert response["status"] == "completed", response

    # 4) 结果登记：拿到 result:// 引用。
    result_receipt = response["receipt"]
    assert result_receipt["resultHandle"].startswith("result://")

    # 5) 按引用取回结果 —— 值只能来自真实容器执行。
    payload = registry.read_result(
        token, target, ORIGIN, receipts,
        PythonResultReceipt.model_validate(result_receipt))

    assert payload["kind"] == "matrix"
    values = payload["values"]
    # 汇总部分与手算逐项一致。行序由容器内 `sorted()` 按码点决定（不是本用例
    # 能预设的），所以按等级取成映射再比，不比行序。
    summary = {row[0]: row[1] for row in values[1:4]}
    assert summary == EXPECTED_LEVEL_TOTALS
    assert sum(summary.values()) == EXPECTED_TOTAL

    # 异常明细**可定位**：空键数、重复键数、未匹配键都随结果一起返回（不静默丢弃）。
    detail = {row[0]: row[1:] for row in values[4:]}
    assert detail["空键数"][0] == 1
    assert detail["重复键数"][0] == 1
    assert detail["未匹配键"][0] == 1
    assert detail["未匹配键"][1] == "404"


def test_continuous_chain_result_is_authorization_bound():
    """结果与授权绑定：撤销授权后按引用取回必须被拒（不发布陈旧结果）。"""
    registry = PythonDatasetRegistry()
    target = _scope()
    token = registry.create(target, ORIGIN)
    receipts = [
        registry.register(token, target, ORIGIN, 0, MEMBERS),
        registry.register(token, target, ORIGIN, 1, ORDERS),
    ]
    response = python_analysis.run_registered_result_internal(
        data_registry=registry,
        scope_token=token,
        scope=target,
        origin=ORIGIN,
        receipts=receipts,
        code=ANALYSIS_CODE,
    )
    assert response["status"] == "completed", response
    receipt = PythonResultReceipt.model_validate(response["receipt"])

    # 取回一次：正常。
    assert registry.read_result(token, target, ORIGIN, receipts, receipt)["kind"] == "matrix"

    # 撤销授权后再取回：必须被拒（结果的生命周期绑定到那次授权）。
    registry.revoke(token, target, ORIGIN)
    with pytest.raises(DatasetRejected):
        registry.read_result(token, target, ORIGIN, receipts, receipt)

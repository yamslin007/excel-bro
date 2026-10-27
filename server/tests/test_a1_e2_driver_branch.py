"""A1-E2 驱动里「B 路到底有没有发生」这一判定的单测。

这不是端到端证据（真实模型那一轮在 `scripts/drive_real_sdk_result_expiry.py` 里，
工件在 `outputs/`），它守的是**判定**本身：

- 没再读 ⇒ 只能算 A 路（不许"看起来像读过"）；
- 再读且被拒 ⇒ 才是 B 路，且必须带上那次调用的操作身份与固定码；
- 再读但没被拒 ⇒ `B-unrefused`（那是缺陷，不是"通过"）；
- 被拒的读取**不算交值**，计划依据仍然是旧那一版；
- "变更发生在第几次桥调用之后"记的是**调用序号**，不是"第几次成功交值"——
  两者只要前面有一次 `query_table` 就不相等，比错了会把"变更后的交值"当成变更前，
  于是 `planSourceBasis` 给出反的结论。
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]


def _load_driver():
    path = REPO / "scripts" / "drive_real_sdk_result_expiry.py"
    spec = importlib.util.spec_from_file_location("a1e2_driver_under_test", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


driver = _load_driver()
MEMBERS = driver.MEMBERS
ORDERS = driver.ORDERS
FULL = [{"sheet": MEMBERS, "range": "A1:C6"}, {"sheet": ORDERS, "range": "A1:B5"}]


def _call(index: int, tool: str, *, refused: bool = False, code=None, sources=None, op=None):
    operation_id = op or f"op_{index:032x}"
    return {
        "index": index,
        "operationId": operation_id,
        "stepId": f"bridge.{tool}:{operation_id[len('op_'):][:12]}",
        "tool": tool,
        "sources": sources,
        "refused": refused,
        "code": code,
        "gate": {"ok": not refused, **({"errorCode": code} if code else {})},
    }


CHANGE_AFTER_FIRST = {
    "afterServedCall": 1,
    "sheet": ORDERS,
    "cell": "B2",
    "before": 100,
    "after": 250,
}


def test_no_source_change_is_not_a_branch():
    served = [_call(1, "prepare_python_data", sources=FULL)]
    assert driver.second_read_evidence(served, None)["branch"] == "none"
    assert driver.second_read_evidence(None, CHANGE_AFTER_FIRST)["branch"] == "none"


def test_model_that_never_reads_again_is_branch_a():
    served = [
        _call(1, "prepare_python_data", sources=FULL),
        _call(2, "query_table"),
    ]
    evidence = driver.second_read_evidence(served, CHANGE_AFTER_FIRST)
    assert evidence == {"branch": "A", "attempted": False, "attempts": []}


def test_refused_second_read_is_branch_b_with_the_production_code():
    served = [
        _call(1, "prepare_python_data", sources=FULL),
        _call(2, "prepare_python_data", refused=True, code="source_changed", sources=FULL),
    ]
    evidence = driver.second_read_evidence(served, CHANGE_AFTER_FIRST)
    assert evidence["branch"] == "B"
    assert evidence["attempted"] is True
    assert evidence["refused"] is True
    assert evidence["code"] == "source_changed"
    assert evidence["operationId"] == served[1]["operationId"]
    assert evidence["stepId"] == served[1]["stepId"]
    # 同一来源范围：两次读的是同一批表，两次之间只有"数据变了"。
    assert evidence["sameScope"] is True


def test_second_read_that_was_served_is_reported_as_a_defect_not_a_pass():
    served = [
        _call(1, "prepare_python_data", sources=FULL),
        _call(2, "prepare_python_data", sources=FULL),
    ]
    evidence = driver.second_read_evidence(served, CHANGE_AFTER_FIRST)
    assert evidence["branch"] == "B-unrefused"
    assert evidence["refused"] is False


def test_different_source_scope_is_flagged():
    served = [
        _call(1, "prepare_python_data", sources=FULL),
        _call(
            2,
            "prepare_python_data",
            refused=True,
            code="source_changed",
            sources=[{"sheet": ORDERS, "range": "A1:B5"}],
        ),
    ]
    assert driver.second_read_evidence(served, CHANGE_AFTER_FIRST)["sameScope"] is False


def test_refused_read_does_not_count_as_a_served_version():
    """被拒的读取只回了固定码、一个值都没给 ⇒ 计划依据仍然是**旧**那一版。"""
    served = [
        _call(1, "prepare_python_data", sources=FULL),
        _call(2, "prepare_python_data", refused=True, code="source_changed", sources=FULL),
    ]
    assert driver.plan_source_basis(served, CHANGE_AFTER_FIRST) == "old"


def test_change_position_is_a_call_index_not_a_serve_counter():
    """先 `query_table` 再 `prepare`：变更记在第 2 次调用之后，交值发生在第 2 次。

    把序号误写成"第几次成功交值"（=1）时，第 2 次那次**发生在变更之前**的交值会被
    当成"变更之后"，`plan_source_basis` 就会反着说成 `new`。这里钉住正确口径。
    """
    served = [
        _call(1, "query_table"),
        _call(2, "prepare_python_data", sources=FULL),
        _call(3, "prepare_python_data", refused=True, code="source_changed", sources=FULL),
    ]
    change = dict(CHANGE_AFTER_FIRST, afterServedCall=2)
    assert driver.plan_source_basis(served, change) == "old"
    assert driver.second_read_evidence(served, change)["branch"] == "B"
    # 反向读法（当成"交值计数"）会得出相反结论 —— 用它说明这不是等价的写法。
    wrong = dict(CHANGE_AFTER_FIRST, afterServedCall=1)
    assert driver.plan_source_basis(served, wrong) == "new"

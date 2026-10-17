"""结果计划编译层把分析结果的空字符串折成 ``None``（2026-10-10 实机缺陷回归）。

真实故障：WPS 端确认写回时，宿主把写入的 ``""`` 当"清空单元格"，回读得到
``null``；计划值仍是 ``""``，端上的裸严格比较判"回读与计划不一致"，拒绝写回
（Excel 端因为自己的值比对本就归一而能过 ⇒ 双端分叉）。

修法在**编译层**：分析结果里的 ``""`` 是"这一格没有内容"的占位，写入意义上的
空白只有 ``None`` 一种（计划核心的 ``blankCellMatrix`` 也一律用 ``null``）。
折完之后计划值与回读值同形，端上的严格比对**保持不动** —— 于是
``wpsEmptyStringBaseline`` 钉的"写入没生效必须被拒"仍然成立。

判据（缺一不算过）：
1. 编译出的动作值里没有 ``""``，对应位置是 ``None``；
2. 验收条件 ``rangeEquals.expected`` 与动作值**同形**；
3. 非空字符串**不被误折**（``"0"``、``" "``、``"金卡"`` 原样保留）；
4. 反向变异：拆掉折叠，用例必须变红。
"""

from __future__ import annotations

import pytest

from server.app.local_results import local_results
from server.app.models import DataToolResult, PlanRequest
from server.app.result_plan import compile_result_plan, result_matrix


@pytest.fixture(autouse=True)
def _clean():
    local_results.reset()


def _item(rows, headers=("等级", "金额", "明细")):
    result = DataToolResult(
        requestId="req-blank",
        title="Python 分析结果",
        headers=list(headers),
        rows=[list(row) for row in rows],
        sourceSheets=["订单表", "会员表"],
        scannedRows=len(rows),
        complete=True,
        calculation="按等级汇总订单金额，未匹配订单归入'未匹配'。",
    )
    return local_results.register("task-blank", result, None)


def _request():
    return PlanRequest.model_validate({
        "prompt": "写入结果",
        "hostId": "wps-et",
        "workbook": {
            "name": "w.xlsx",
            "capturedAt": "2026-10-10T00:00:00Z",
            "activeWorksheet": "汇总表",
            "worksheets": [{
                "name": "汇总表", "usedRange": "A1:C3",
                "rowCount": 3, "columnCount": 3, "headers": ["a", "b", "c"],
            }],
        },
    })


def test_result_matrix_folds_empty_string_to_none():
    item = _item([["未匹配", 400, ""], ["金卡", 400, ""]])
    matrix = result_matrix(item, include_header=True)
    assert matrix[0] == ["等级", "金额", "明细"]
    assert matrix[1] == ["未匹配", 400, None]
    assert matrix[2] == ["金卡", 400, None]
    # 正面锚点：折出来的必须是 None（`is None`，不是 == ""），否则等于没修
    assert matrix[1][2] is None
    assert all(cell != "" for row in matrix for cell in row)


def test_result_matrix_keeps_meaningful_strings():
    """非空字符串不得被误折 —— `"0"`/`" "`/中文都不是"空白占位"。"""
    item = _item([["0", " ", "金卡"]])
    matrix = result_matrix(item, include_header=False)
    assert matrix == [["0", " ", "金卡"]]


def test_compiled_plan_has_no_empty_string_cells():
    item = _item([["未匹配", 400, ""]])
    plan = compile_result_plan(
        item, _request(), target_sheet="汇总表", target_range="A1:C2"
    )
    values = plan.actions[0].values
    assert values[1] == ["未匹配", 400, None]
    assert all(cell != "" for row in values for cell in row)


def test_acceptance_criterion_matches_action_values():
    """验收条件与动作值必须同形 —— 否则公式路径的期望值又会与回读分叉。"""
    item = _item([["未匹配", 400, ""]])
    plan = compile_result_plan(
        item, _request(), target_sheet="汇总表", target_range="A1:C2"
    )
    criteria = [
        criterion for criterion in plan.acceptanceCriteria
        if criterion.type == "rangeEquals"
    ]
    assert len(criteria) == 1
    assert criteria[0].expected == plan.actions[0].values
    assert all(cell != "" for row in criteria[0].expected for cell in row)


def test_blank_folding_is_not_a_tautology():
    """把折叠函数换成恒等，折叠断言必须变红 —— 证明断言有判别力。"""
    from server.app import result_plan

    item = _item([["未匹配", 400, ""]])
    original = result_plan._blank_cell_of_write
    try:
        result_plan._blank_cell_of_write = lambda value: value  # 拆掉折叠
        folded = result_plan.result_matrix(item, include_header=False)
    finally:
        result_plan._blank_cell_of_write = original
    assert folded == [["未匹配", 400, ""]]  # 拆掉后确实折不出来
    # 还原后必须恢复折叠行为（逐字节回到同一实现）
    assert result_plan.result_matrix(item, include_header=False) == [
        ["未匹配", 400, None]
    ]

"""`map_column_values` 工具：模型给规则，服务端算值。

纯计算层的判定在 test_column_mapping.py。本文件只覆盖接线部分：
参数解析、快照越界拒绝、以及"模型不经手目标矩阵"这件事本身。
"""

from __future__ import annotations

import pytest

import server.app.excel_agent as excel_agent
from server.app.models import LocalPlanRuleResponse, PlanRequest, PlanResponse

MAPPING = [
    {"from": "办公用品", "to": "行政"},
    {"from": "电子设备", "to": "资产"},
]


def _request(*, truncated: bool = False, data_rows: list | None = None) -> PlanRequest:
    return PlanRequest.model_validate(
        {
            "prompt": "把类别列翻译成归属",
            "workbook": {
                "name": "orders.xlsx",
                "capturedAt": "2026-09-27T00:00:00Z",
                "activeWorksheet": "订单",
                "worksheets": [
                    {
                        "name": "订单",
                        "usedRange": "A1:C4",
                        "rowCount": 4,
                        "columnCount": 3,
                        "headers": ["单号", "类别", "归属"],
                        "dataRows": data_rows
                        if data_rows is not None
                        else [
                            ["X001", "办公用品", None],
                            ["X002", "电子设备", None],
                            ["X003", None, None],
                        ],
                        "truncated": truncated,
                    }
                ],
            },
        }
    )


def _local_request() -> PlanRequest:
    """`planningReadPolicy: "none"` 且 `dataRows` 为空 —— WPS 实际送上来的形态。

    值不在服务端手上，所以这条路径只能下发规则。构造时刻意保持 `dataRows: []`：
    若测试自己填了值，就测不出"服务端没有值也能出规则"这件事。
    """
    return PlanRequest.model_validate(
        {
            "prompt": "把类别列翻译成归属",
            "planningReadPolicy": "none",
            "workbook": {
                "name": "orders.xlsx",
                "capturedAt": "2026-09-27T00:00:00Z",
                "activeWorksheet": "订单",
                "worksheets": [
                    {
                        "name": "订单",
                        "usedRange": "A1:C4",
                        "rowCount": 4,
                        "columnCount": 3,
                        "headers": ["单号", "类别", "归属"],
                        "dataRows": [],
                        "truncated": True,
                    }
                ],
            },
        }
    )


def _call(request: PlanRequest, **overrides):
    arguments = {
        "sourceSheet": "订单",
        "sourceRange": "B2:B4",
        "targetRange": "C2:C4",
        "mapping": MAPPING,
        "blankLabel": "未分类",
    }
    arguments.update(overrides)
    return excel_agent._execute_tool(request, "map_column_values", arguments)


def test_tool_is_registered_in_full_and_terminal_sets() -> None:
    for tools in (excel_agent.AGENT_TOOLS, excel_agent.AGENT_TOOLS_WITH_DATA):
        names = [tool["function"]["name"] for tool in tools]
        assert "map_column_values" in names


def test_tool_is_absent_from_plan_only_repair_set() -> None:
    # plan_only 是 submit_plan 校验失败后的定向修复通道，
    # 此时只应暴露 submit_plan。
    names = [tool["function"]["name"] for tool in excel_agent.PLAN_ONLY_TOOLS]
    assert names == ["submit_plan"]


def test_tool_is_not_treated_as_readonly() -> None:
    # 它产出写入计划，不能进只读缓存，也不应计入只读批次。
    assert "map_column_values" not in excel_agent.READONLY_TOOLS


def test_server_computes_values_from_snapshot() -> None:
    result = _call(_request())
    assert isinstance(result, PlanResponse)
    assert result.plan.actions[0].type == "writeValues"
    assert result.plan.actions[0].sheet == "订单"
    assert result.plan.actions[0].range == "C2:C4"
    # 模型只给了两条规则和一个空白标签，这三个值是服务端算出来的。
    assert result.plan.actions[0].values == [["行政"], ["资产"], ["未分类"]]


def test_plan_warns_that_result_is_a_snapshot() -> None:
    # 查表结果仍会随源数据改动而过期，计划必须说明这一点。
    plan = _call(_request()).plan
    assert any("不会自动跟着变" in warning for warning in plan.warnings)
    assert any("会被覆盖" in warning for warning in plan.warnings)


def test_mapping_rule_count_is_recorded_as_assumption() -> None:
    plan = _call(_request()).plan
    assert any("未自行补充" in item for item in plan.assumptions)
    assert any("未分类" in item for item in plan.assumptions)


def test_unmapped_value_is_rejected_with_the_missing_values() -> None:
    request = _request(
        data_rows=[
            ["X001", "办公用品", None],
            ["X002", "差旅", None],
            ["X003", "礼品", None],
        ]
    )
    with pytest.raises(ValueError) as excinfo:
        _call(request)
    message = str(excinfo.value)
    assert "差旅" in message and "礼品" in message


def test_range_beyond_snapshot_is_rejected_not_padded() -> None:
    # 补 None 会被映射成空白标签，写进实际上没有数据的行。
    with pytest.raises(ValueError, match="超出已捕获范围"):
        _call(_request(), sourceRange="B2:B50", targetRange="C2:C50")


def test_truncated_snapshot_says_so_in_the_rejection() -> None:
    with pytest.raises(ValueError, match="快照已截断"):
        _call(_request(truncated=True), sourceRange="B2:B50", targetRange="C2:C50")


def test_range_starting_before_captured_origin_is_rejected() -> None:
    request = PlanRequest.model_validate(
        {
            "prompt": "映射",
            "workbook": {
                "name": "orders.xlsx",
                "capturedAt": "2026-09-27T00:00:00Z",
                "activeWorksheet": "订单",
                "worksheets": [
                    {
                        "name": "订单",
                        "usedRange": "B5:C7",
                        "rowCount": 3,
                        "columnCount": 2,
                        "headers": ["类别", "归属"],
                        "dataRows": [["办公用品", None], ["电子设备", None]],
                    }
                ],
            },
        }
    )
    with pytest.raises(ValueError, match="早于"):
        _call(request, sourceRange="B2:B4", targetRange="C2:C4")


def test_column_beyond_used_range_is_rejected() -> None:
    with pytest.raises(ValueError, match="已用区域"):
        _call(_request(), sourceRange="Z2:Z4", targetRange="C2:C4")


def test_duplicate_mapping_source_is_rejected_not_silently_overwritten() -> None:
    # 同一来源值给两个目标标签是需求本身有歧义，静默取一个等于替用户决定。
    with pytest.raises(ValueError, match="两个目标标签"):
        _call(
            _request(),
            mapping=[*MAPPING, {"from": "办公用品", "to": "杂项"}],
        )


def test_duplicate_mapping_with_identical_target_is_accepted() -> None:
    result = _call(_request(), mapping=[*MAPPING, {"from": "办公用品", "to": "行政"}])
    assert result.plan.actions[0].values == [["行政"], ["资产"], ["未分类"]]


def test_empty_blank_label_is_rejected() -> None:
    with pytest.raises(ValueError, match="blankLabel"):
        _call(_request(), blankLabel="   ")


def test_empty_mapping_is_rejected() -> None:
    with pytest.raises(ValueError, match="至少要有一条"):
        _call(_request(), mapping=[])


def test_malformed_mapping_entry_is_rejected() -> None:
    with pytest.raises(ValueError, match="非空字符串"):
        _call(_request(), mapping=[{"from": "办公用品", "to": ""}])


def test_target_sheet_defaults_to_source_sheet() -> None:
    plan = _call(_request(), targetSheet="").plan
    assert plan.actions[0].sheet == "订单"


def test_overlapping_source_and_target_is_rejected() -> None:
    with pytest.raises(ValueError, match="重叠"):
        _call(_request(), targetRange="B2:B4")


def test_unknown_source_sheet_is_rejected() -> None:
    with pytest.raises(ValueError, match="未找到工作表"):
        _call(_request(), sourceSheet="不存在的表")


def test_shape_mismatch_between_source_and_target_is_rejected() -> None:
    with pytest.raises(ValueError, match="形状不一致"):
        _call(_request(), targetRange="C2:C3")


def test_read_policy_none_returns_a_rule_instead_of_a_plan() -> None:
    """服务端手上没有值时下发规则，而不是拒绝、也不是补 None 硬算。

    同一批参数在 `"ifNeeded"` 下会被 §10.1 的越界检查拒绝（见
    `test_range_beyond_snapshot_is_rejected_not_padded`）；分叉后不再走到那里。
    """
    response = _call(_local_request())
    assert isinstance(response, LocalPlanRuleResponse)
    assert response.kind == "local_plan_rule"
    assert response.rule.sourceSheet == "订单"
    assert response.rule.targetSheet == "订单"
    assert response.rule.sourceRange == "B2:B4"
    assert response.rule.targetRange == "C2:C4"
    assert response.rule.blankLabel == "未分类"
    assert [(pair.source, pair.target) for pair in response.rule.mapping] == [
        ("办公用品", "行政"),
        ("电子设备", "资产"),
    ]


def test_local_rule_carries_no_cell_values() -> None:
    """规则里不能出现任何源值或映射结果 —— 单射映射可逆，带上就等于上传源列。"""
    dumped = _call(_local_request()).model_dump()
    assert "values" not in dumped["rule"]
    assert "actions" not in dumped


def test_local_rule_keeps_the_snapshot_warning() -> None:
    # 客户端算不改变"这是一张快照"的事实，告知不能因为换了计算位置就丢掉。
    response = _call(_local_request())
    assert any("不会自动跟着变" in item for item in response.warnings)
    assert any("未分类" in item for item in response.assumptions)


def test_local_rule_still_rejects_overlap_without_values() -> None:
    """重叠/形状/解析三类错误不需要值就能判定，必须在下发前拦下。

    留到客户端才发现等于把错误推到模型离开循环之后，那时它已无法修正参数。
    """
    with pytest.raises(ValueError, match="重叠"):
        _call(_local_request(), targetRange="B2:B4")


def test_local_rule_still_rejects_shape_mismatch_without_values() -> None:
    with pytest.raises(ValueError, match="形状不一致"):
        _call(_local_request(), targetRange="C2:C3")


def test_local_rule_still_rejects_unknown_source_sheet() -> None:
    with pytest.raises(ValueError, match="未找到工作表"):
        _call(_local_request(), sourceSheet="不存在的表")


def test_local_rule_still_rejects_empty_blank_label() -> None:
    with pytest.raises(ValueError, match="blankLabel"):
        _call(_local_request(), blankLabel="   ")


def test_local_rule_deduplicates_identical_mapping_pairs() -> None:
    # 目标标签相同的重复项无歧义，接受但只下发一条：客户端拿到的是去重结果。
    response = _call(
        _local_request(),
        mapping=[*MAPPING, {"from": "办公用品", "to": "行政"}],
    )
    assert len(response.rule.mapping) == 2


def test_local_rule_still_rejects_conflicting_mapping_pairs() -> None:
    with pytest.raises(ValueError, match="两个目标标签"):
        _call(
            _local_request(),
            mapping=[*MAPPING, {"from": "办公用品", "to": "杂项"}],
        )


def test_read_policy_if_needed_still_computes_on_the_server() -> None:
    """分叉没有改变既有通道的行为：快照带值时仍然服务端算出可执行计划。"""
    response = _call(_request())
    assert isinstance(response, PlanResponse)
    assert response.plan.actions[0].values == [["行政"], ["资产"], ["未分类"]]


def test_wire_keys_are_identical_on_both_serialization_paths() -> None:
    """两条序列化路径必须吐出同一批键名。

    仓库里既有 `response_model`（FastAPI 默认 `by_alias=True`）也有 SSE 里的裸
    `model_dump(mode="json")`。若 `ColumnMappingPair` 带别名，两者会分别吐出
    `from`/`to` 与 `source`/`target`，客户端的契约校验只在其中一条上通过 ——
    而这种失败只在跑到那条路径时才出现。这条用例把"无别名"钉住。
    """
    response = _call(_local_request())
    bare = response.model_dump(mode="json")
    aliased = response.model_dump(mode="json", by_alias=True)
    assert bare == aliased
    assert set(bare["rule"]["mapping"][0]) == {"source", "target"}


def test_progress_facts_expose_rule_count_and_blank_label() -> None:
    facts = excel_agent._tool_argument_facts(
        "map_column_values",
        {
            "sourceSheet": "订单",
            "sourceRange": "B2:B4",
            "targetRange": "C2:C4",
            "mapping": MAPPING,
            "blankLabel": "未分类",
        },
    )
    rendered = {fact["label"]: fact["value"] for fact in facts}
    assert rendered["映射规则"] == "2"
    assert rendered["空白写作"] == "未分类"
    assert rendered["目标"] == "C2:C4"

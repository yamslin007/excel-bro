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


def test_unmapped_value_becomes_a_clarification_not_a_plan() -> None:
    """未覆盖的来源值不再作为"工具报错"抛出，而是变成一份澄清响应。

    旧行为是把未覆盖值当 `ValueError` 抛给模型，模型收到的是一句错误文本，
    既不知道要补哪些类别、也没有结构化的补救动作。现在归因与缺口都进
    `clarification`，模型和界面都能据此直接问用户。
    """
    request = _request(
        data_rows=[
            ["X001", "办公用品", None],
            ["X002", "差旅", None],
            ["X003", "礼品", None],
        ]
    )
    response = _call(request)
    assert isinstance(response, LocalPlanRuleResponse)
    clarification = response.rule.clarification
    assert clarification is not None
    assert clarification.missingCategories == ["差旅", "礼品"]


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


def test_blank_cell_without_policy_returns_structured_clarification() -> None:
    """**读到**空白且策略未定 → 一张结构化澄清卡，不是报错也不是默认标签。

    关键在于它发生在读源值之后。旧实现在解析参数时就拒，等于要求调用方
    在读到数据之前先替用户回答。
    """
    response = _call(_request(), blankLabel="   ")
    assert isinstance(response, LocalPlanRuleResponse)
    assert response.rule.clarification is not None
    assert response.rule.clarification.reason == "blankPolicyNeeded"
    # 计数来自真实读取；没有读证据就不许声称"源区域有空白"。
    assert response.rule.clarification.blankCellCount == 1
    # 四个选项是空白那组，不含类别/分隔符那组 —— 混进来会误导用户。
    assert response.rule.clarification.resolvedBy == [
        "keepTarget",
        "clearTarget",
        "writeBlankValue",
        "writeEmptyString",
    ]


def test_no_blank_in_source_does_not_require_a_blank_policy() -> None:
    """源列没有空白时，没给空白策略也应当照常产出计划。"""
    response = _call(
        _request(
            data_rows=[
                ["X001", "办公用品", None],
                ["X002", "电子设备", None],
                ["X003", "办公用品", None],
            ]
        )
    )
    assert isinstance(response, PlanResponse)
    assert response.plan.actions[0].values == [["行政"], ["资产"], ["行政"]]


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


def test_local_rule_defers_formula_or_fixed_value_warning_to_client() -> None:
    # 服务端还没有源值，不能提前宣称客户端一定写固定值。
    response = _call(_local_request())
    assert not any("固定值" in item for item in response.warnings)
    assert any("会被覆盖" in item for item in response.warnings)
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


def test_local_rule_defers_blank_decision_to_the_client() -> None:
    """服务端手上没有源值，就**不能**声称源区域有空白、也不该拒。

    下发"策略未定"的规则，让客户端读到真实数据后再决定要不要问 ——
    这才是"按真实数据驱动澄清"。拒掉会让模型立刻去问用户一个
    谁都还不知道答案的问题。
    """
    response = _call(_local_request(), blankLabel="   ")
    assert isinstance(response, LocalPlanRuleResponse)
    assert response.rule.blankPolicy is None
    assert response.rule.clarification is None


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


def _tool_schema() -> dict:
    for tool in excel_agent.AGENT_TOOLS:
        if tool["function"]["name"] == "map_column_values":
            return tool["function"]
    raise AssertionError("map_column_values 不在工具表里")


def test_schema_guides_whole_cell_first_then_declared_delimiter() -> None:
    """映射的匹配顺序：先整格、后按**已声明**的分隔符拆分。

    顺序本身是规则的一部分 —— 反过来（先拆再查整格）会让调用方无法用
    "把整格原文写进映射表"来表达"这一格整体算一个类别"。
    """
    description = _tool_schema()["description"]
    assert "整格" in description
    assert "delimiters" in description
    assert "valueRank" in description
    # 关键动作：说不清就回去问，而不是自行假设。
    assert "submit_clarification" in description
    assert "分隔符" in description


def test_mapping_property_points_back_to_clarification() -> None:
    schema = _tool_schema()["parameters"]["properties"]["mapping"]
    description = schema["description"]
    assert "拆分" in description
    assert "整格" in description


def test_schema_requires_delimiter_and_rank_together() -> None:
    """只给分隔符、不给等级，等于没说完"取哪一个"，schema 必须讲明成对给出。"""
    properties = _tool_schema()["parameters"]["properties"]
    assert "一起给" in properties["delimiters"]["description"]
    assert "rank" in properties["valueRank"]["description"].lower()


def test_schema_pins_rank_polarity_against_the_number_in_the_grade_name() -> None:
    """`rank` 表达**优先级**，不是等级名里的字面数字。

    真实故障：业务等级常以"名字里的 0 最严重"命名，而系统的约定是
    **rank 越大越优先** —— 两者方向相反。模型照抄字面数字（最严重的给 rank 0）
    不会报任何错：服务端原样保留、core 按 rank 取最大、公式再忠实编码本地预期，
    于是整条链路**一致地**给出那个**较低**的等级，用户只会看到结果莫名其妙变了。
    所以 schema 必须把方向写死，并用**抽象**等级名举例。
    """
    properties = _tool_schema()["parameters"]["properties"]
    description = properties["valueRank"]["description"]
    # 正向锚点：方向被明确解释过，而不只是"rank 越大越优先"这种可被两种读法解释的话。
    assert "越大的越优先" in description
    # 必须点明 rank 表达的是**优先级**，并用一个抽象的、可比较的严重程度举反例。
    assert "优先级" in description
    assert "严重" in description
    # 必须留下一个可仿照的抽象例子（参数级说明同样要提方向）。
    assert '"rank":1' in description
    assert "越大越优先" in properties["valueRank"]["items"]["properties"]["rank"]["description"]
    # 负向：等级名属于规则数据，生产提示词里不得出现任何具体业务等级字面量。
    for literal in ("p0", "p1", "p2", "P0", "P1", "P2", "甲级", "乙级"):
        assert literal not in description


def test_prompt_warns_that_rank_is_priority_not_a_grade_number() -> None:
    """系统提示同样要讲方向：填反不报错，只会安静地取到较低等级。"""
    prompt = excel_agent.AGENT_SYSTEM_PROMPT
    assert "优先级" in prompt
    assert "更大的 rank" in prompt
    # 必须点明"不报错"这一危险特征，否则模型没有动力去核对方向。
    assert "不会报错" in prompt


def test_prompt_orders_whole_cell_match_before_splitting() -> None:
    """系统提示必须把三步顺序写死：整格精确 → 已声明分隔符拆分 → 归因。

    只断言"顺序存在"这件行为，不绑定具体分隔符或具体样例值 ——
    样例（某种分隔符、某个具体类别名）属于测试夹具，不进生产提示词。
    """
    prompt = excel_agent.AGENT_SYSTEM_PROMPT
    assert "整格精确匹配优先" in prompt
    assert "只按**已声明的**分隔符拆分" in prompt
    # 归因必须分成"确定"与"不确定"两类，不能混成一个诊断。
    assert "归因**确定**" in prompt
    assert "归因**不确定**" in prompt
    assert "submit_answer" in prompt
    assert "最高级" in prompt


def test_prompt_keeps_two_attributions_apart() -> None:
    """两类归因的措辞必须分开：类别缺口 ≠ 分隔形式未声明。"""
    prompt = excel_agent.AGENT_SYSTEM_PROMPT
    assert "补充类别映射" in prompt
    assert "不要**替用户判定" in prompt or "不要**自己拿一个没声明过的写法去拆" in prompt


def _clarification_request(
    *, data_rows: list, truncated: bool = False
) -> PlanRequest:
    return PlanRequest.model_validate(
        {
            "prompt": "按类别列翻译成归属，多值取最高级",
            "workbook": {
                "name": "orders.xlsx",
                "capturedAt": "2026-09-27T00:00:00Z",
                "activeWorksheet": "订单",
                "worksheets": [
                    {
                        "name": "订单",
                        "usedRange": f"A1:C{len(data_rows) + 1}",
                        "rowCount": len(data_rows) + 1,
                        "columnCount": 3,
                        "headers": ["单号", "类别", "归属"],
                        "dataRows": data_rows,
                        "truncated": truncated,
                    }
                ],
            },
        }
    )


def test_missing_category_is_attributed_to_the_category_gap() -> None:
    """已声明分隔符拆开后出现未覆盖项 → 归因**确定**是类别缺口。

    这条路径曾经被合成一个"分隔形式说不清"的统一诊断，用户会去改一个
    本来没问题的分隔符设置。现在归因与补救动作都收窄到「补类别映射」。
    """
    request = _clarification_request(
        data_rows=[["X001", "办公用品 差旅", None], ["X002", "电子设备", None]]
    )
    response = _call(
        request,
        sourceRange="B2:B3",
        targetRange="C2:C3",
        delimiters=[" "],
        valueRank=[{"value": "电子设备", "rank": 2}, {"value": "办公用品", "rank": 1}],
    )
    assert isinstance(response, LocalPlanRuleResponse)
    clarification = response.rule.clarification
    assert clarification is not None
    assert clarification.reason == "matchedDelimiterMissingCategories"
    assert clarification.missingCategories == ["差旅"]
    # 补救动作只有一条：把分隔符选项塞进来会误导用户。
    assert clarification.resolvedBy == ["addCategories"]
    # 不产出可执行计划。
    assert "actions" not in response.model_dump()


def test_undeclared_separator_stays_undecided() -> None:
    """整格未命中、又含未声明的分隔写法 → 归因**不确定**，给全三个选项。"""
    request = _clarification_request(
        data_rows=[["X001", "办公用品/差旅", None], ["X002", "电子设备", None]]
    )
    response = _call(
        request,
        sourceRange="B2:B3",
        targetRange="C2:C3",
        delimiters=[" "],
        undeclaredSeparators=["/"],
        valueRank=[{"value": "电子设备", "rank": 1}],
    )
    assert isinstance(response, LocalPlanRuleResponse)
    clarification = response.rule.clarification
    assert clarification is not None
    assert clarification.reason == "unresolvedSeparatorOrMapping"
    assert clarification.unresolvedCellCount == 1
    assert clarification.resolvedBy == [
        "specifyOtherDelimiter",
        "addCategories",
        "treatWholeCellAsCategory",
    ]


def test_clarification_carries_no_source_text() -> None:
    """澄清响应里不能落下源格原文 —— 它属于本端数据。"""
    request = _clarification_request(
        data_rows=[["X001", "办公用品/差旅", None], ["X002", "电子设备", None]]
    )
    response = _call(
        request,
        sourceRange="B2:B3",
        targetRange="C2:C3",
        delimiters=[" "],
        undeclaredSeparators=["/"],
        valueRank=[{"value": "电子设备", "rank": 1}],
    )
    dumped = repr(response.model_dump(mode="json"))
    assert "办公用品/差旅" not in dumped
    assert "X001" not in dumped


def test_category_gap_clarification_carries_no_whole_cell_text() -> None:
    """类别缺口的 `missingCategories` 装的是**拆出来的项**，不是整格原文。

    用户要补的是"丙"这一条映射，不是"丙 丁"这一整格 —— 把整格给他等于
    让他自己去拆。这两者的区别就是归因是否确定。
    """
    request = _clarification_request(data_rows=[["X001", "丙 丁", None]])
    response = _call(
        request,
        sourceRange="B2:B2",
        targetRange="C2:C2",
        delimiters=[" "],
        valueRank=[{"value": "甲", "rank": 1}],
    )
    clarification = response.rule.clarification
    assert clarification is not None
    assert clarification.missingCategories == ["丙", "丁"]
    assert "丙 丁" not in repr(response.model_dump(mode="json"))


def test_local_channel_also_refuses_incomplete_rule() -> None:
    """`"none"` 通道没有源值，但"规则不全"不需要源值就能判定。

    只在 `ifNeeded` 分支拦会让同一份残缺规则在下发通道被静默执行：
    客户端按 mapping 声明顺序硬排一个顺序，而那个顺序不是用户说的优先级。
    """
    response = _call(_local_request(), delimiters=[" "])
    assert isinstance(response, LocalPlanRuleResponse)
    assert response.rule.clarification is not None
    assert response.rule.clarification.reason == "unresolvedSeparatorOrMapping"


def test_clarification_never_returns_a_plan() -> None:
    """归因一旦确定"算不出正确答案"，就不能返回可执行计划。

    返回一份建立在猜测之上的矩阵，等于把猜测伪装成结果。
    """
    request = _clarification_request(
        data_rows=[["X001", "丙 丁", None]]
    )
    response = _call(
        request,
        sourceRange="B2:B2",
        targetRange="C2:C2",
        delimiters=[" "],
        valueRank=[{"value": "甲", "rank": 1}],
    )
    assert isinstance(response, LocalPlanRuleResponse)
    assert "actions" not in response.model_dump()
    assert any("未被改动" in warning for warning in response.warnings)

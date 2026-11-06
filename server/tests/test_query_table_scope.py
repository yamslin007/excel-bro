"""「本次查哪几张表」的协议侧：`sheets` 参数与完整性原因。

真机现场（2026-11-01）：四张表都勾了，用户只说「只读取来源!A1:B5」，结果卡却列出
四张来源表、三张各报「缺少列2」，提示语还是「扫描已达上限」。两处根因各有一组用例：

1. **目标表丢了**：授权集合被当成本次查询范围。协议必须能表达"只查这张表"
   （`sheets` / `range` 带表名），且**归一化不许把它丢掉**；
2. **不完整原因被压成一个布尔**：`complete=false` 一律被说成扫描触顶。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from server.app.agent_tools import local_query_tool_definitions, query_bindings_of
from server.app.models import DataToolResult, QueryTableArguments


def _arguments(**overrides: object) -> QueryTableArguments:
    return QueryTableArguments.model_validate({"mode": "rows", **overrides})


def _result(**overrides: object) -> DataToolResult:
    payload: dict[str, object] = {
        "requestId": "q1",
        "title": "查询结果",
        "headers": ["合计"],
        "rows": [[15]],
        "sourceSheets": ["来源"],
        "scannedRows": 5,
        "complete": True,
        "calculation": "求和",
        "warnings": [],
    }
    payload.update(overrides)
    return DataToolResult.model_validate(payload)


# ---------------------------------------------------------------------------
# 1. 目标表：可表达、可校验、不被归一化丢掉
# ---------------------------------------------------------------------------


def test_sheets_is_optional_and_kept_verbatim() -> None:
    assert _arguments().sheets == []
    assert _arguments(sheets=["来源"]).sheets == ["来源"]
    # 名字两端的空白由执行侧 trim，协议层不改写内容（改了就再也对不上表名）。
    assert _arguments(sheets=["来源", "结果1"]).sheets == ["来源", "结果1"]


def test_sheets_survives_normalization() -> None:
    """模型 → 工具链路上最容易静默丢字段的一跳。

    `claude_sdk_agent` 把工具参数过一遍 `QueryTableArguments.model_validate(...)
    .model_dump(mode="json")` 再送桥。pydantic 默认 `extra="ignore"`：少声明一个字段，
    模型传了也会在这里**无声消失**。所以"能传"和"能到执行器"必须分开断言。
    """
    normalized = QueryTableArguments.model_validate(
        {
            "mode": "aggregate",
            "scope": "selected",
            "sheets": ["来源"],
            "range": "来源!A1:B5",
            "hasHeaders": False,
            "metrics": [{"operation": "sum", "field": "列2", "outputName": "合计"}],
        }
    ).model_dump(mode="json")
    assert normalized["sheets"] == ["来源"]
    assert normalized["range"] == "来源!A1:B5"
    assert normalized["hasHeaders"] is False


def test_query_table_schema_exposes_sheets_and_states_the_boundary() -> None:
    tool = next(
        item
        for item in local_query_tool_definitions()
        if item["function"]["name"] == "query_table"
    )
    properties = tool["function"]["parameters"]["properties"]
    assert "sheets" in properties
    description = properties["sheets"]["description"]
    # 模型必须能从 schema 里读到"授权只是上限"这条口径，而不是靠提示词里的一句话。
    assert "上限" in description
    assert "本次查询范围" in description
    assert "range" in description


def test_tool_description_requires_binding_the_named_sheet() -> None:
    tool = next(
        item
        for item in local_query_tool_definitions()
        if item["function"]["name"] == "query_table"
    )
    description = tool["function"]["description"]
    assert "必须绑定本次要查的表" in description
    assert "不是本次查询范围" in description
    # 说清"不绑定会怎样"，否则模型只会以为是建议。
    assert "报错" in description


def test_query_bindings_keep_structure_and_drop_values() -> None:
    """留证据的那份参数只留结构：筛选值、单元格内容都不进去。"""
    bindings = query_bindings_of(
        {
            "mode": "rows",
            "scope": "selected",
            "sheets": ["来源"],
            "range": "来源!A1:B5",
            "hasHeaders": False,
            "filters": [{"field": "会员", "operator": "equals", "value": "张三"}],
            "fields": ["会员", "金额"],
        }
    )
    assert bindings == {
        "mode": "rows",
        "scope": "selected",
        "sheets": ["来源"],
        "range": "来源!A1:B5",
        "hasHeaders": False,
    }
    assert "张三" not in str(bindings)


# ---------------------------------------------------------------------------
# 2. 完整性：complete 与原因不许各说一套
# ---------------------------------------------------------------------------


def test_incomplete_reason_pins_the_kind_payload() -> None:
    assert _result(
        complete=False,
        incomplete=[{"kind": "missing_field", "sheet": "结果1", "fields": ["列2"]}],
    ).incomplete[0].fields == ["列2"]
    # 缺字段却没有字段名：没有信息量的原因等于没说，直接拒。
    with pytest.raises(ValidationError):
        _result(
            complete=False,
            incomplete=[{"kind": "missing_field", "sheet": "结果1"}],
        )
    # "没有数据行"不该带字段名。
    with pytest.raises(ValidationError):
        _result(
            complete=False,
            incomplete=[
                {"kind": "empty_sheet", "sheet": "汇总表", "fields": ["列2"]}
            ],
        )


def test_complete_and_reasons_must_agree() -> None:
    with pytest.raises(ValidationError):
        _result(complete=False)
    with pytest.raises(ValidationError):
        _result(
            complete=True,
            incomplete=[{"kind": "missing_field", "sheet": "结果1", "fields": ["列2"]}],
        )


def test_truncation_is_separate_from_completeness() -> None:
    truncated = _result(
        truncated=True, truncation={"shownRows": 20, "totalRows": 137, "limit": 20}
    )
    # 截断不是"不完整"：数据是全的，只是没全显示。
    assert truncated.complete is True
    assert truncated.incomplete == []
    with pytest.raises(ValidationError):
        _result(truncated=True)


def test_structure_carries_the_same_facts_as_the_card() -> None:
    """模型回执与结果卡读同一份事实。

    以前回执只有 `complete`，文案却在本地写死「扫描已达上限」；两边各说一套时，
    模型会顺着界面的说法去解释（"你只要了前 3 行"），那就更难查了。
    """
    from server.app.local_results import structure_of

    registered = _make_registered()
    structure = structure_of(registered)
    assert structure["complete"] is False
    assert structure["truncated"] is False
    assert structure["incomplete"] == [
        {"kind": "missing_field", "sheet": "结果1", "fields": ["列2"]}
    ]
    assert structure["truncation"] is None


def _make_registered():
    from server.app.local_results import LocalResultRegistry

    registry = LocalResultRegistry()
    return registry.register(
        "turn-scope-1",
        _result(
            complete=False,
            sourceSheets=["来源", "结果1"],
            warnings=["「结果1」未找到字段：列2"],
            incomplete=[
                {"kind": "missing_field", "sheet": "结果1", "fields": ["列2"]}
            ],
        ),
        None,
    )

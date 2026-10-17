"""规划前能力收窄：未验收的动作在**生成计划时**就被拒绝。

对齐指导 §4.3 的要求是"排除必须发生在规划前"。这条测试锁定三件事：

1. 请求声明了 WPS 宿主/版本时，`submit_plan` 提交未验收动作会被拒绝，
   而不是让用户看到一个生成好的计划再在执行预检被拒；
2. `map_column_values` 终局工具不能成为绕过这条检查的第二条路径；
3. **未声明宿主标识时行为不变** —— 兼容入口不能被变成新的拒绝源。

第 3 条与第 1 条同等重要：把一条兼容路径顺手改成"永远拒绝"，
同样是一种"外观相同却行为不同"的失误。
"""

from __future__ import annotations

import pytest

from server.app.excel_agent import _enforce_action_capabilities
from server.app.models import AnalysisPlan, PlanRequest


def _request(**overrides) -> PlanRequest:
    payload = {
        "prompt": "把分类列填好",
        "workbook": {
            "name": "夹具.xlsx",
            "capturedAt": "2026-09-28T00:00:00Z",
            "activeWorksheet": "订单",
            "worksheets": [
                {
                    "name": "订单",
                    "rowCount": 3,
                    "columnCount": 2,
                    "headers": ["分类", "标签"],
                    "dataRows": [["A", None], ["B", None]],
                }
            ],
        },
    }
    payload.update(overrides)
    return PlanRequest.model_validate(payload)


def _plan(actions: list[dict]) -> AnalysisPlan:
    return AnalysisPlan.model_validate(
        {
            "id": "cap-test",
            "title": "能力收窄用例",
            "summary": "用于验证规划前能力收窄。",
            "actions": actions,
        }
    )


WRITE_VALUES = {
    "type": "writeValues",
    "sheet": "订单",
    "range": "B2:B3",
    "values": [["已分类"], ["已分类"]],
}

CREATE_CHART = {
    "type": "createChart",
    "sheet": "订单",
    "sourceRange": "A1:B3",
    "chartType": "ColumnClustered",
}


def test_wps_request_rejects_unverified_action_before_planning() -> None:
    """WPS 上 createChart 未验收 → 生成计划阶段即拒绝，原因写清楚。"""
    request = _request(
        hostId="wps-et",
        hostVersion="12.0",
        selfReportedActions=["writeValues", "createChart"],
    )
    with pytest.raises(ValueError) as error:
        _enforce_action_capabilities(request, _plan([WRITE_VALUES, CREATE_CHART]))
    assert "createChart" in str(error.value)
    assert "验收" in str(error.value)


def test_wps_request_allows_verified_action() -> None:
    request = _request(
        hostId="wps-et",
        hostVersion="12.0",
        selfReportedActions=["writeValues"],
    )
    _enforce_action_capabilities(request, _plan([WRITE_VALUES]))  # 不抛


def test_clear_range_variant_is_checked_per_variant() -> None:
    """`clearRange:contents` 已开放，`formats` 未验收 —— 按变体判定。"""
    request = _request(
        hostId="wps-et",
        hostVersion="12.0",
        selfReportedActions=["clearRange:contents", "clearRange:formats"],
    )
    contents = {
        "type": "clearRange",
        "sheet": "订单",
        "range": "B2:B3",
        "applyTo": "contents",
    }
    formats = {**contents, "applyTo": "formats"}

    _enforce_action_capabilities(request, _plan([contents]))  # 不抛
    with pytest.raises(ValueError) as error:
        _enforce_action_capabilities(request, _plan([formats]))
    assert "clearRange:formats" in str(error.value)


def test_client_claim_cannot_promote_unverified_action() -> None:
    """红线：自报"我有这个动作"不能让它进入计划。"""
    request = _request(
        hostId="wps-et",
        hostVersion="12.0",
        # 自报把全部动作都写上也一样
        selfReportedActions=[
            "writeValues",
            "writeFormulas",
            "clearRange:contents",
            "createChart",
            "createPivotTable",
            "addImage",
        ],
    )
    with pytest.raises(ValueError):
        _enforce_action_capabilities(request, _plan([CREATE_CHART]))


def test_empty_claim_yields_rejection_not_openness() -> None:
    """红线：客户端漏报得到拒绝，不是"没说就放开"。"""
    request = _request(hostId="wps-et", hostVersion="12.0", selfReportedActions=[])
    with pytest.raises(ValueError):
        _enforce_action_capabilities(request, _plan([WRITE_VALUES]))


def test_unknown_host_is_rejected_not_open() -> None:
    """红线：未知宿主不能变成"全部放开"。

    这一条曾写错并被测试抓出来：初版实现按"能不能解析出宿主"分流，
    `hostId: "office-online"` 解析失败 → 直接 `return`，整道闸被跳过。
    正确分流是按"**字段有没有值**"：声明了但识别不了 → 上界为空 → 全部拒绝。
    """
    request = _request(
        hostId="office-online",
        hostVersion="12.0",
        selfReportedActions=["writeValues"],
    )
    with pytest.raises(ValueError):
        _enforce_action_capabilities(request, _plan([WRITE_VALUES]))


def test_blank_host_id_is_treated_as_absent() -> None:
    """空白字符串是"没声明"，不是"声明了一个识别不了的宿主"。"""
    request = _request(hostId="   ", hostVersion="12.0")
    _enforce_action_capabilities(request, _plan([CREATE_CHART]))  # 不抛


def test_unknown_version_is_rejected_not_open() -> None:
    request = _request(
        hostId="wps-et",
        hostVersion="99.0",
        selfReportedActions=["writeValues"],
    )
    with pytest.raises(ValueError):
        _enforce_action_capabilities(request, _plan([WRITE_VALUES]))


def test_request_without_host_id_keeps_existing_behaviour() -> None:
    """未声明 hostId 的调用方行为不变 —— 兼容入口不被变成新的拒绝源。

    这一类请求的动作边界由各自通道的后续校验负责（Office 执行预检、
    文件夹保存与撤销校验）。在这里顺手拒绝会误伤现有入口。
    """
    request = _request()  # 没有 hostId
    assert request.hostId is None
    _enforce_action_capabilities(request, _plan([CREATE_CHART]))  # 不抛


def test_office_host_allows_all_declared_actions() -> None:
    request = _request(
        hostId="excel-office",
        hostVersion="16.0.18025.20000",
        selfReportedActions=["writeValues", "createChart", "createPivotTable"],
    )
    _enforce_action_capabilities(request, _plan([WRITE_VALUES, CREATE_CHART]))


def test_office_host_still_narrows_by_client_claim() -> None:
    """Excel 侧虽登记了全部动作，客户端自报仍能收窄（交集，不是并集）。"""
    request = _request(
        hostId="excel-office",
        hostVersion="16.0.18025.20000",
        selfReportedActions=["writeValues"],
    )
    with pytest.raises(ValueError):
        _enforce_action_capabilities(request, _plan([CREATE_CHART]))


def test_rejection_message_names_the_offending_action() -> None:
    request = _request(
        hostId="wps-et",
        hostVersion="12.0",
        selfReportedActions=["writeValues", "addShape"],
    )
    add_shape = {
        "type": "addShape",
        "sheet": "订单",
        "shapeType": "rectangle",
        "targetRange": "A1",
    }
    with pytest.raises(ValueError) as error:
        _enforce_action_capabilities(request, _plan([WRITE_VALUES, add_shape]))
    assert "addShape" in str(error.value)


# ---------------------------------------------------------------------------
# 收窄后的 **schema 合法性**（与上面的"拒绝语义"是两件事）
#
# 上面锁的是"未验收动作进不了计划"；下面锁的是"收窄后的工具定义本身能发出去"。
# 二者必须分开：把"收窄到空集"实现成 `{"anyOf": []}`，语义上也是拒绝，
# 但**空 anyOf 是非法 JSON Schema**，多数模型网关（实测 DeepSeek 的
# anthropic 兼容端点）会直接拒绝**整轮请求**：
#   API Error: 400 Invalid schema for function 'submit_plan':
#   [] has less than 1 item
# 结果不是"submit_plan 不可用"，而是**整次规划请求 400**，
# 表现为"请求莫名失败"，把真正的协商失败原因（漏报 / 版本不匹配）完全掩盖。
# ---------------------------------------------------------------------------


def test_empty_bound_never_emits_empty_any_of() -> None:
    """红线：上界为空集时不得产出 `{"anyOf": []}`。

    空集的三条真实来路 —— 漏报 `selfReportedActions`、`hostVersion` 不匹配、
    未知宿主 —— 都必须走到"schema 合法 + 整轮能发出去 + 提交时被第二道防线拒绝"，
    而不是"schema 非法 + 整轮 400"。
    """
    from server.app.agent_tools import host_pruned_terminal_definitions

    for overrides in (
        {"hostId": "wps-et", "hostVersion": "12.0", "selfReportedActions": []},
        {"hostId": "wps-et", "hostVersion": "99.0", "selfReportedActions": ["writeValues"]},
        {"hostId": "office-online", "hostVersion": "12.0", "selfReportedActions": ["writeValues"]},
    ):
        request = _request(**overrides)
        definitions = host_pruned_terminal_definitions(request)
        names = {tool["function"]["name"] for tool in definitions}
        assert "submit_plan" not in names
        assert "submit_clarification" in names
        assert "submit_answer" in names


def test_pruned_schema_is_legal_json_schema_when_narrowed() -> None:
    """正向锚点：收窄**非空**时仍是原来的收窄语义（不是恒回落全量）。

    没有这一条，`test_empty_bound_never_emits_empty_any_of` 可以被
    "无条件回落到全部变体"蒙过去 —— 那样虽然不再有空 `anyOf`，
    但"排除必须发生在规划前"就整条失效了。
    """
    from server.app.agent_tools import host_pruned_terminal_definitions

    request = _request(
        hostId="wps-et", hostVersion="12.0", selfReportedActions=["writeValues"],
    )
    items = _submit_plan_action_items(host_pruned_terminal_definitions(request))
    types = {variant["properties"]["type"]["const"] for variant in items["anyOf"]}
    # WPS 支持集是 ['clearRange:contents','writeFormulas','writeValues']，
    # 自报只有 writeValues ⇒ 交集只应剩它一个。
    assert types == {"writeValues"}
    # 未验收动作不得出现为候选
    assert "createChart" not in types


def test_unpruned_definition_is_returned_unchanged_without_host_id() -> None:
    """未声明宿主 → 定义逐字不变，不引入任何新的 schema 形状。"""
    from server.app.agent_tools import host_pruned_terminal_definitions
    from server.app.excel_agent import AGENT_TOOLS_WITH_DATA

    request = _request()
    assert host_pruned_terminal_definitions(request) == AGENT_TOOLS_WITH_DATA


def _submit_plan_action_items(definitions: list[dict]) -> dict:
    for definition in definitions:
        if definition["function"]["name"] == "submit_plan":
            return definition["function"]["parameters"]["properties"]["plan"][
                "properties"
            ]["actions"]["items"]
    raise AssertionError("定义里没有 submit_plan")

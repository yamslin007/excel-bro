"""`query_table` 的表头三态与工具说明（2026-11-01 现场，独立于范围确认问题）。

现场：用户说「只读取来源表 A1:B5，告诉我 B 列五个数的合计」，模型却问
「B1 是否参与计算」，并声称算这五个数需要写入目标单元格。

模型侧要能**表达**「这块区域没有表头」，否则它只能猜，猜错就静默少算一行；
工具说明里也必须说清「无表头时字段名是列1/列2」「首行同样计入」
「只读不得转成写入方案」。这几条在这里钉住，避免说明被后续改动冲掉。
"""

from __future__ import annotations

from server.app.agent_tools import local_query_tool_definitions
from server.app.models import QueryTableArguments


def _tool(name: str) -> dict:
    for definition in local_query_tool_definitions():
        if definition["function"]["name"] == name:
            return definition
    raise AssertionError(f"缺少工具定义：{name}")


def test_header_mode_has_three_states_and_defaults_to_auto() -> None:
    # 省略 = 由本地按内容判断（既有行为）。
    assert QueryTableArguments.model_validate({"mode": "rows"}).hasHeaders is None
    # 用户说过「五个数 / 第一行也算」时，模型显式声明无表头。
    assert (
        QueryTableArguments.model_validate(
            {"mode": "rows", "hasHeaders": False}
        ).hasHeaders
        is False
    )
    assert (
        QueryTableArguments.model_validate(
            {"mode": "rows", "hasHeaders": True}
        ).hasHeaders
        is True
    )


def test_query_table_schema_exposes_header_mode() -> None:
    schema = _tool("query_table")["function"]["parameters"]
    assert "hasHeaders" in schema["properties"]


def test_query_table_description_states_headerless_and_read_only_rules() -> None:
    description = _tool("query_table")["function"]["description"]
    # 无表头区域的字段名怎么起。
    assert "列1" in description
    # 首行也是数据，不得排除。
    assert "首行" in description or "第一行" in description
    # 只读不得转成写入方案。
    assert "不得转成写入方案" in description
    # 区域查询是有的，且要说明「区域外不参与计算」，否则模型会拿整表结果冒充。
    assert "range" in description
    assert "区域外" in description


def test_describe_tool_does_not_call_row_one_a_header_blindly() -> None:
    description = _tool("describe_authorized_tables")["function"]["description"]
    assert "列1" in description
    assert "所有行" in description

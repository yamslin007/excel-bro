"""`query_table` 的显式只读区域（协议侧）。

原需求限定「来源!A1:B5」，整表查询不能代替它。执行侧的唯一判据在
`packages/core/src/queryRange.ts`；这里按同一套规则镜像校验，用例表与
`packages/core/src/queryRange.test.ts` 保持一致 —— 两边一起改，否则会出现
"协议放行、执行器拒"（用户看到两个说法）或"区域被悄悄忽略"（结果多算区域外数据）。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from server.app.agent_tools import local_query_tool_definitions
from server.app.models import QueryTableArguments


def _arguments(**overrides: object) -> QueryTableArguments:
    return QueryTableArguments.model_validate({"mode": "rows", **overrides})


def test_range_is_optional_and_kept_verbatim() -> None:
    assert _arguments().range is None
    assert _arguments(range="A1:B5").range == "A1:B5"
    assert _arguments(range="  来源!A1:B5  ").range == "来源!A1:B5"
    assert _arguments(range="$B$2:$D$4").range == "$B$2:$D$4"
    assert _arguments(range="'我的 表'!B2").range == "'我的 表'!B2"


@pytest.mark.parametrize(
    "value",
    [
        "",
        "   ",
        "A:A",
        "1:5",
        "A1:",
        "A1:B5:C7",
        "名称区域",
        "A0",
        "A1:B0",
        "A1:B5!",
        "a!b!A1",
        "A1048577",
        "A1:XFE1",
        "a/b!A1",
        "a?b!A1",
        "a[b]!A1",
        "a:b!A1",
        "x" * 40 + "!A1",
    ],
)
def test_range_rejects_what_would_break_the_promise(value: str) -> None:
    with pytest.raises(ValidationError):
        _arguments(range=value)


def test_query_table_schema_exposes_range() -> None:
    tool = next(
        definition
        for definition in local_query_tool_definitions()
        if definition["function"]["name"] == "query_table"
    )
    schema = tool["function"]["parameters"]
    assert "range" in schema["properties"]
    # 模型要能读懂"给了区域，区域外不参与计算"。
    assert "区域外" in schema["properties"]["range"]["description"]


def test_query_table_description_tells_the_model_to_use_range() -> None:
    tool = next(
        definition
        for definition in local_query_tool_definitions()
        if definition["function"]["name"] == "query_table"
    )
    description = tool["function"]["description"]
    assert "range" in description
    assert "区域外" in description

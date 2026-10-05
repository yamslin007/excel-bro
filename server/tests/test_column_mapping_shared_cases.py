"""共享夹具驱动的跨语言行为锁（Python 侧）。

用例来自 `config/column_mapping_cases.json`，TypeScript 侧
`packages/core/src/columnMapping.test.ts` 加载同一份文件、断言同一批期望值。
**单边修改任一实现都会让一侧变红。**

为什么不用 `test_safety_mirror.py` 那种抓源码文本的做法：那种做法能锁"清单"
（公式黑名单之类），锁不住"逻辑"。映射计算的一致性在于逐格输出，只有跑同一批
输入才能验。

本文件不重复 `test_column_mapping.py` 的单元用例，它只负责"两份实现是否一致"。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from server.app.column_mapping import (
    BlankPolicyNeededError,
    ColumnMappingError,
    NeedsClarificationError,
    OverlappingRangeError,
    ShapeMismatchError,
    UnmappedCategoryError,
    build_column_mapping_actions,
)

_CASES_PATH = (
    Path(__file__).resolve().parents[2] / "config" / "column_mapping_cases.json"
)

# 错误种类 → (类, 消息片段)。与 TS 侧 `EXPECTATIONS` 逐项对应。
#
# `categoriesUnknown` 与 `needsClarification` 是**两种不同的归因**，刻意分成
# 两个 key：前者补救动作单一（补类别映射），后者有两个可能（补分隔符 / 补映射），
# 不能互相冒充。把它们合成一个错误种类会让调用方把「要用户选」误当成
# 「补映射表就行」。
_EXPECTATIONS: dict[str, tuple[type[ColumnMappingError], str]] = {
    "unmapped": (UnmappedCategoryError, "不在映射表内"),
    "categoriesUnknown": (UnmappedCategoryError, "不在映射表内"),
    "needsClarification": (NeedsClarificationError, "无法归类"),
    "shape": (ShapeMismatchError, "形状"),
    "overlap": (OverlappingRangeError, "重叠"),
    "parse": (ColumnMappingError, "区域"),
    # 「读到真实空白但策略未定」是**待澄清**，不是报错：实现不替用户挑默认值。
    "blankPolicy": (BlankPolicyNeededError, "空白"),
    "sheetName": (ColumnMappingError, "工作表名"),
}


def _load_cases() -> list[dict[str, Any]]:
    payload = json.loads(_CASES_PATH.read_text(encoding="utf-8"))
    return payload["cases"]


_CASES = _load_cases()


def _invoke(case: dict[str, Any]) -> list[dict[str, Any]]:
    data = case["input"]
    return build_column_mapping_actions(
        source_sheet=data["sourceSheet"],
        target_sheet=data["targetSheet"],
        source_range=data["sourceRange"],
        target_range=data["targetRange"],
        source_values=data["sourceValues"],
        mapping=data["mapping"],
        blank_label=data.get("blankLabel"),
        blank_policy=data.get("blankPolicy"),
        delimiters=data.get("delimiters"),
        value_rank=data.get("valueRank"),
        undeclared_separators=data.get("undeclaredSeparators"),
    )


def test_fixture_is_non_empty_and_well_formed() -> None:
    assert _CASES
    for case in _CASES:
        declared = sum(
            1 for key in ("expected", "expectedActions", "error") if key in case
        )
        assert declared == 1, (
            f"用例「{case['name']}」必须且只能声明 "
            "expected / expectedActions / error 之一"
        )
        if "error" in case:
            assert case["error"] in _EXPECTATIONS


@pytest.mark.parametrize("case", _CASES, ids=[case["name"] for case in _CASES])
def test_shared_case(case: dict[str, Any]) -> None:
    if "expected" in case:
        actions = _invoke(case)
        assert len(actions) == 1
        assert actions[0]["type"] == "writeValues"
        assert actions[0]["sheet"] == case["input"]["targetSheet"]
        assert actions[0]["range"] == case["input"]["targetRange"]
        assert actions[0]["values"] == case["expected"]
        return

    if "expectedActions" in case:
        # 需要跳过/清空位置时，只断言"值的形状"是不够的 —— 必须钉住
        # 被跳过的位置根本没出现在任何 range 里。
        actions = _invoke(case)
        assert len(actions) == len(case["expectedActions"])
        for index, expected in enumerate(case["expectedActions"]):
            actual = actions[index]
            assert actual["sheet"] == case["input"]["targetSheet"]
            for key, value in expected.items():
                assert actual[key] == value, (
                    f"第 {index + 1} 个动作的 {key} 与期望不一致"
                )
        return

    error_class, fragment = _EXPECTATIONS[case["error"]]
    with pytest.raises(error_class) as excinfo:
        _invoke(case)
    assert fragment in str(excinfo.value)
    if "unmappedValues" in case:
        assert excinfo.value.values == case["unmappedValues"]
    if "unmatchedItems" in case:
        assert excinfo.value.values == case["unmatchedItems"]
    if "unresolvedCells" in case:
        assert excinfo.value.cells == case["unresolvedCells"]

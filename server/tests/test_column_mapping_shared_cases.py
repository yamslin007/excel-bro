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
    ColumnMappingError,
    OverlappingRangeError,
    ShapeMismatchError,
    UnmappedCategoryError,
    build_column_mapping_actions,
)

_CASES_PATH = (
    Path(__file__).resolve().parents[2] / "config" / "column_mapping_cases.json"
)

# 错误种类 → (类, 消息片段)。与 TS 侧 `EXPECTATIONS` 逐项对应。
_EXPECTATIONS: dict[str, tuple[type[ColumnMappingError], str]] = {
    "unmapped": (UnmappedCategoryError, "不在映射表内"),
    "shape": (ShapeMismatchError, "形状"),
    "overlap": (OverlappingRangeError, "重叠"),
    "parse": (ColumnMappingError, "区域"),
    "blankLabel": (ColumnMappingError, "空白标签"),
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
        blank_label=data["blankLabel"],
    )


def test_fixture_is_non_empty_and_well_formed() -> None:
    assert _CASES
    for case in _CASES:
        has_expected = "expected" in case
        has_error = "error" in case
        assert has_expected != has_error, (
            f"用例「{case['name']}」既没有期望值也没有期望错误，或两者都给了"
        )
        if has_error:
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

    error_class, fragment = _EXPECTATIONS[case["error"]]
    with pytest.raises(error_class) as excinfo:
        _invoke(case)
    assert fragment in str(excinfo.value)
    if "unmappedValues" in case:
        assert excinfo.value.values == case["unmappedValues"]

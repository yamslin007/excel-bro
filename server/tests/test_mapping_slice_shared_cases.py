"""切片逻辑的跨语言行为锁（Python 侧，阶段 4 §24）。

用例来自 `config/mapping_slice_cases.json`，TypeScript 侧
`apps/wps-addin/src/wpsMappingSlice.shared.test.ts` 加载同一份文件，对
`sliceMappingSource` 断言同一批期望值。**单边修改任一实现都会让一侧变红。**

为什么需要这一份：`column_mapping_cases.json` 锁的是纯函数
`build_column_mapping_actions`。两端各自的「从已读数据里切出源区域」不在它的
覆盖范围内，此前只有各自语言的单元测试，一致性靠人工对读源码 —— 而对读得出的
"两侧策略一致"结论是**错的**：服务端不拦左侧列早于原点、客户端不拦右侧列溢出，
各自静默折成空值。

错误只按**种类**共享，消息文案两侧各自声明：服务端说「快照/已捕获」，客户端说
「已读入」，同一件事在两侧的措辞本就不同，强行统一没有收益。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

import server.app.excel_agent as excel_agent
from server.app.models import PlanRequest

_CASES_PATH = (
    Path(__file__).resolve().parents[2] / "config" / "mapping_slice_cases.json"
)

# 错误种类 → 服务端消息片段。与 TS 侧 `EXPECTATIONS` 逐项对应。
_EXPECTATIONS: dict[str, str] = {
    "rowBeforeOrigin": "行，早于",
    "rowBeyondCaptured": "行，超出已捕获范围",
    "columnBeforeOrigin": "列，早于",
    "columnBeyondCaptured": "列，超出该范围",
    "parse": "range 必须是 A1 或 A1:B10 格式",
    "reversed": "结束位置不能早于开始位置",
}

_SHEET_NAME = "共享用例表"


def _load_cases() -> list[dict[str, Any]]:
    payload = json.loads(_CASES_PATH.read_text(encoding="utf-8"))
    return payload["cases"]


_CASES = _load_cases()


def _request(spec: dict[str, Any]) -> PlanRequest:
    """把夹具的「原点 + 值矩阵」装成快照形态。

    `rowCount` 刻意填成远大于实际捕获量的值：行边界按 `len(headers) +
    len(dataRows)` 判定而不是 `rowCount`，这条决定也一并锁住。

    `usedRange` 只写原点单格：`_used_origin` 只取起点，宽度由 `columnCount`
    单独给，写成区间反而可能与 `columnCount` 不一致而看不出是哪个在起作用。
    """
    values: list[list[Any]] = spec["values"]
    origin = f"{excel_agent._column_name(spec['startColumn'])}{spec['startRow']}"
    return PlanRequest.model_validate(
        {
            "prompt": "切片夹具",
            "workbook": {
                "name": "fixture.xlsx",
                "capturedAt": "2026-09-27T00:00:00Z",
                "activeWorksheet": _SHEET_NAME,
                "worksheets": [
                    {
                        "name": _SHEET_NAME,
                        "usedRange": origin,
                        "rowCount": spec["startRow"] + len(values) + 9999,
                        "columnCount": spec["columnCount"],
                        "headers": values[0] if values else [],
                        "dataRows": values[1:],
                        "truncated": spec["truncated"],
                    }
                ],
            },
        }
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
def test_shared_slice_case(case: dict[str, Any]) -> None:
    request = _request(case["sheet"])
    if "expected" in case:
        _, values = excel_agent._snapshot_region(request, _SHEET_NAME, case["range"])
        assert values == case["expected"]
        return

    with pytest.raises(ValueError) as excinfo:
        excel_agent._snapshot_region(request, _SHEET_NAME, case["range"])
    assert _EXPECTATIONS[case["error"]] in str(excinfo.value)

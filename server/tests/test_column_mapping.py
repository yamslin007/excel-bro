"""`server/app/column_mapping.py` 的单元用例。

跨宿主贯通由 test_column_mapping_baseline.py 覆盖（文件夹通道实际落盘）。
本文件只锁定纯计算层自身的判定，尤其是三处"拒绝而非兜底"。
"""

from __future__ import annotations

import pytest

from server.app.column_mapping import (
    BlankPolicyNeededError,
    OverlappingRangeError,
    ShapeMismatchError,
    UnmappedCategoryError,
    build_column_mapping_actions,
    map_column_values,
)

MAPPING = {"甲": "一类", "乙": "二类"}
BLANK = "未分类"


def _build(**overrides):
    kwargs = {
        "source_sheet": "源表",
        "target_sheet": "目标表",
        "source_range": "B2:B4",
        "target_range": "H2:H4",
        "source_values": [["甲"], ["乙"], [None]],
        "mapping": MAPPING,
        "blank_label": BLANK,
    }
    kwargs.update(overrides)
    return build_column_mapping_actions(**kwargs)


def test_produces_exactly_one_write_values_action() -> None:
    # 只产出一个动作：拆分会引入"部分动作成功"的中间态，
    # 而该中间态在三条通道上的恢复行为并不一致。
    actions = _build()
    assert len(actions) == 1
    assert actions[0] == {
        "type": "writeValues",
        "sheet": "目标表",
        "range": "H2:H4",
        "values": [["一类"], ["二类"], [BLANK]],
    }


def test_unmapped_values_are_all_reported_at_once() -> None:
    # 一次只报一个会让调用方反复试错；映射表是人工维护的，
    # 一轮拿到完整缺口才能一次补完。
    with pytest.raises(UnmappedCategoryError) as excinfo:
        _build(source_values=[["丙"], ["丁"], ["丙"]])
    assert excinfo.value.values == ["丙", "丁"]  # 去重且保序


def test_whitespace_only_string_counts_as_blank() -> None:
    assert map_column_values(
        source_values=[["  "], ["\t"], [""]], mapping=MAPPING, blank_label=BLANK
    ) == [[BLANK], [BLANK], [BLANK]]


def test_explicit_mapping_entry_beats_blank_policy() -> None:
    # 调用方把 "" 写进 mapping，说明它要区分"空串"与"真空格"，
    # 这个意图优先于空白策略。
    values = map_column_values(
        source_values=[[""], [None]],
        mapping={**MAPPING, "": "空串专用"},
        blank_label=BLANK,
    )
    assert values == [["空串专用"], [BLANK]]


def test_overlapping_ranges_on_same_sheet_are_rejected() -> None:
    # 写入逐格落地，重叠时"读源"与"写目标"的先后顺序决定结果，
    # 而这个顺序在三条通道上不一致。
    with pytest.raises(OverlappingRangeError):
        _build(source_sheet="同表", target_sheet="同表", target_range="B2:B4")


def test_adjacent_ranges_on_same_sheet_are_allowed() -> None:
    actions = _build(source_sheet="同表", target_sheet="同表", target_range="C2:C4")
    assert actions[0]["range"] == "C2:C4"


def test_cross_sheet_same_address_is_not_overlap() -> None:
    actions = _build(source_range="B2:B4", target_range="B2:B4")
    assert actions[0]["sheet"] == "目标表"


@pytest.mark.parametrize(
    ("overrides", "reason"),
    [
        ({"source_values": [["甲"], ["乙"]]}, "行数少于源区域"),
        ({"source_values": [["甲", "乙"], ["乙", "甲"], [None, None]]}, "列数多于源区域"),
        ({"target_range": "H2:H9"}, "目标区域与源区域行数不一致"),
        ({"target_range": "H2:I4"}, "目标区域与源区域列数不一致"),
    ],
)
def test_shape_mismatch_is_rejected(overrides: dict, reason: str) -> None:
    with pytest.raises(ShapeMismatchError):
        _build(**overrides)


def test_unparsable_range_is_rejected_before_mapping() -> None:
    with pytest.raises(ValueError, match="无法解析"):
        _build(target_range="这不是区域")


def test_unspecified_blank_policy_does_not_block_when_source_has_no_blank() -> None:
    """源区域**一格空白都没有**时，没说空白怎么办也不该拦住预览。

    这是"还没读源数据就先问空白怎么办"的正面反例：旧实现在入口处直接以
    「空白标签为空」拒绝，于是调用方只能在读到源值之前先去问用户 ——
    而多数源列根本没有空白，用户被问的是一个不存在的问题。
    """
    actions = _build(blank_label="", source_values=[["甲"], ["乙"], ["甲"]])
    assert actions == [
        {
            "type": "writeValues",
            "sheet": "目标表",
            "range": "H2:H4",
            "values": [["一类"], ["二类"], ["一类"]],
        }
    ]


def test_blank_cell_without_policy_asks_instead_of_picking_a_default() -> None:
    """**读到**空白且策略未定 → 待澄清，不替用户挑默认标签。

    静默挑一个（历史上就是空白一律写「未分类」）会让用户在单元格层面
    看不出那是实现替他选的。计数必须来自真实读取。
    """
    with pytest.raises(BlankPolicyNeededError) as excinfo:
        _build(blank_label="")
    assert excinfo.value.blank_count == 1


def test_keep_policy_excludes_blank_position_from_the_write_range() -> None:
    """保留目标原值 = 该位置**不出现在任何 range 里**，不是写个空值。"""
    actions = _build(blank_policy={"kind": "keep"})
    assert [action["range"] for action in actions] == ["H2:H3"]
    assert actions[0]["values"] == [["一类"], ["二类"]]


def test_clear_policy_emits_clear_range_not_an_empty_string() -> None:
    """清空走 `clearRange:contents`：写空串会留下常量空字符串，不是真空单元格。"""
    actions = _build(blank_policy={"kind": "clear"})
    assert [action["type"] for action in actions] == [
        "writeValues",
        "clearRange",
    ]
    assert actions[1] == {
        "type": "clearRange",
        "sheet": "目标表",
        "range": "H4:H4",
        "applyTo": "contents",
    }


def test_explicit_empty_string_is_written_as_a_value_not_skipped() -> None:
    """显式写空字符串：整块区域照写，`""` 是内容而不是"跳过"的记号。"""
    actions = _build(blank_policy={"kind": "value", "value": ""})
    assert actions[0]["values"] == [["一类"], ["二类"], [""]]


def test_single_cell_range_is_accepted() -> None:
    actions = _build(source_range="B2", target_range="H2", source_values=[["甲"]])
    assert actions[0]["values"] == [["一类"]]


def test_non_string_source_value_matches_mapping_by_text() -> None:
    # 宿主可能把数字类别读成 int；按文本匹配，但未命中时报回原值。
    assert map_column_values(
        source_values=[[2026]], mapping={"2026": "年度"}, blank_label=BLANK
    ) == [["年度"]]
    with pytest.raises(UnmappedCategoryError) as excinfo:
        map_column_values(source_values=[[7]], mapping=MAPPING, blank_label=BLANK)
    assert excinfo.value.values == [7]

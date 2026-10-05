"""列映射的**语义边界**（Python 侧镜像）。

与 TS 侧 `packages/core/src/columnMappingSemantics.test.ts` 一一对应，
钉同一批判断：空白与未知类别分开、显式项不因"像空白"而降级、
显式空串只覆盖空串。

多值部分钉的是**两类归因不得互相冒充**：

- 已按声明的分隔符拆开、拆出的项有未覆盖 → `UnmappedCategoryError`，
  补救动作只有「补类别映射」一条；
- 整格没命中、又像是用了别的分隔符写法 → `NeedsClarificationError`，
  补救动作有两个可能（补分隔符 / 补映射），必须让用户选。

把两者合成一个错误，会让"该让用户选"被误当成"补映射表就行"。
"""

from __future__ import annotations

from typing import Any

import pytest

from server.app.column_mapping import (
    NeedsClarificationError,
    UnmappedCategoryError,
    build_column_mapping_actions,
)


def _run(
    source_values: list[list[Any]],
    mapping: dict[str, str],
    blank_label: str,
    *,
    delimiters: list[str] | None = None,
    value_rank: list[dict[str, Any]] | None = None,
    undeclared_separators: list[str] | None = None,
) -> list[dict[str, Any]]:
    return build_column_mapping_actions(
        source_sheet="表一",
        target_sheet="表一",
        source_range=f"A1:A{len(source_values)}",
        target_range=f"C1:C{len(source_values)}",
        source_values=source_values,
        mapping=mapping,
        blank_label=blank_label,
        delimiters=delimiters,
        value_rank=value_rank,
        undeclared_separators=undeclared_separators,
    )


def test_blank_uses_blank_label_without_error() -> None:
    actions = _run([[None]], {"甲": "一"}, "空")
    assert actions[0]["values"] == [["空"]]


def test_unmapped_value_is_rejected_not_folded_into_blank_label() -> None:
    # 若实现把"未命中"兜成 blankLabel，这条会静默通过。
    with pytest.raises(UnmappedCategoryError):
        _run([["乙"]], {"甲": "一"}, "空")


def test_blank_and_unknown_mixed_reports_only_unknown() -> None:
    with pytest.raises(UnmappedCategoryError) as excinfo:
        _run([[None], [" "], ["乙"]], {"甲": "一"}, "空")
    assert excinfo.value.values == ["乙"]


def test_explicit_non_empty_entry_wins() -> None:
    # "无问题" 是非空来源值，不因语义像"没有内容"而降级到空白策略。
    actions = _run([["无问题"]], {"无问题": "/"}, "未填")
    assert actions[0]["values"] == [["/"]]


def test_blank_still_uses_blank_label_when_explicit_entry_exists() -> None:
    actions = _run([[None], ["无问题"]], {"无问题": "/"}, "未填")
    assert actions[0]["values"] == [["未填"], ["/"]]


def test_explicit_empty_string_entry_covers_only_empty_string() -> None:
    actions = _run([[""], [None], ["  "]], {"": "空串"}, "真空格")
    assert actions[0]["values"] == [["空串"], ["真空格"], ["真空格"]]


def test_all_missing_values_are_collected_once() -> None:
    with pytest.raises(UnmappedCategoryError) as excinfo:
        _run([["丙"], ["丁"], ["丙"]], {"甲": "一"}, "空")
    assert excinfo.value.values == ["丙", "丁"]


def test_symbol_labels_are_written_literally() -> None:
    actions = _run([["甲"], ["乙"]], {"甲": "/", "乙": "—"}, "空")
    assert actions[0]["values"] == [["/"], ["—"]]


# --- 多值：两类归因不得互相冒充 -------------------------------------------


def test_split_missing_category_is_attributed_to_category_gap() -> None:
    """已声明的分隔符对上了，缺的就是类别。

    报 `UnmappedCategoryError` 而不是要用户澄清 —— 补救动作是确定的：
    把缺的类别补进映射表。若误报成"分隔符不明"，用户会去改一个没问题
    的设置，而真正的缺口一直留着。
    """
    with pytest.raises(UnmappedCategoryError) as excinfo:
        _run(
            [["甲 丙"]],
            {"甲": "壹"},
            "空",
            delimiters=[" "],
            value_rank=[{"value": "甲", "rank": 1}],
        )
    # 缺失项是**拆开后的那一项**，不是整格原文：用户要补的是「丙」。
    assert excinfo.value.values == ["丙"]


def test_split_all_missing_still_attributed_to_category_gap() -> None:
    """拆开的项一个都没覆盖，归因仍然是类别缺口，且缺失项按首次出现顺序列全。"""
    with pytest.raises(UnmappedCategoryError) as excinfo:
        _run(
            [["丙 丁"]],
            {"甲": "壹"},
            "空",
            delimiters=[" "],
            value_rank=[{"value": "甲", "rank": 1}],
        )
    assert excinfo.value.values == ["丙", "丁"]


def test_unmatched_with_suspected_other_separator_needs_user_choice() -> None:
    """整格没命中、又含未声明的分隔符写法时，归因**不确定**。

    可能是分隔符没声明全，也可能是这格本身缺映射 —— 两种补救动作不同，
    报 `NeedsClarificationError` 让用户选，并由调用方据此给出三个具体选项。
    """
    with pytest.raises(NeedsClarificationError) as excinfo:
        _run(
            [["甲/乙"], ["甲"]],
            {"甲": "壹", "乙": "贰"},
            "空",
            delimiters=[" "],
            value_rank=[{"value": "甲", "rank": 1}, {"value": "乙", "rank": 0}],
            undeclared_separators=["/"],
        )
    assert excinfo.value.cells == ["甲/乙"]


def test_unmatched_without_any_separator_is_category_gap() -> None:
    """看不出任何分隔形式时归因是确定的：就是缺这个类别。

    这条与上一条成对：同一入口只差"格子里有没有像分隔符的写法"，
    归因却不同。两者的补救动作不同，输出也必须不同。
    """
    with pytest.raises(UnmappedCategoryError) as excinfo:
        _run(
            [["丙"]],
            {"甲": "壹"},
            "空",
            delimiters=[" "],
            value_rank=[{"value": "甲", "rank": 1}],
        )
    assert excinfo.value.values == ["丙"]


def test_whole_cell_match_wins_over_splitting() -> None:
    """整格能精确命中时就不拆。

    优先级是**整格精确 → 空白 → 拆分**：调用方把整格原文写进 mapping，
    就是在说"这一格整体算一个类别"。若拆分逻辑抢在前面，
    这个明确表达的意图会被推翻。
    """
    actions = _run(
        [["甲 乙"]],
        {"甲 乙": "合", "甲": "壹", "乙": "贰"},
        "空",
        delimiters=[" "],
        value_rank=[{"value": "甲", "rank": 1}, {"value": "乙", "rank": 0}],
    )
    assert actions[0]["values"] == [["合"]]


def test_highest_rank_wins_after_split() -> None:
    """拆分后按规则给出的等级取最高级。"""
    actions = _run(
        [["乙 甲"], ["甲"]],
        {"甲": "壹", "乙": "贰"},
        "空",
        delimiters=[" "],
        value_rank=[{"value": "甲", "rank": 1}, {"value": "乙", "rank": 0}],
    )
    assert actions[0]["values"] == [["壹"], ["壹"]]


def test_declared_delimiter_position_beats_declaration_order() -> None:
    """多候选并存时按**出现位置**切分，不按声明序。

    声明序表达的是优先级（谁更重要），不是切分方式（按什么切）。
    """
    actions = _run(
        [["甲-乙|丙"]],
        {"甲": "壹", "乙|丙": "合"},
        "空",
        delimiters=["|", "-"],
        value_rank=[{"value": "甲", "rank": 1}],
    )
    # '-' 在 index 1 先出现 -> 切成 ['甲','乙|丙']，两项都命中 -> 取 rank 高的甲。
    assert actions[0]["values"] == [["壹"]]


def test_clarification_message_carries_no_whole_cell_text() -> None:
    """归因不确定时，错误文案里不含整格原文。

    这个 message 会经 `ColumnMappingClarification.message` 上到界面/协议；
    带上原文就等于把本端数据顺着"澄清"这条通道送出去。要展示原文的界面
    本来就能读自己的源值，不需要（也不该）由实现代转。
    """
    with pytest.raises(NeedsClarificationError) as excinfo:
        _run(
            [["甲/乙"]],
            {"甲": "一"},
            "空",
            delimiters=[" "],
            undeclared_separators=["/"],
            value_rank=[{"value": "甲", "rank": 1}],
        )
    assert "甲/乙" not in str(excinfo.value)
    # 但"有几格处理不了"这个计数必须留在结构化字段里。
    assert excinfo.value.cells == ["甲/乙"]


def test_category_gap_reports_split_items_not_whole_cell() -> None:
    """类别缺口报的是拆开的项，不是整格原文。

    用户要补的是「丙」「丁」两条映射，给他「丙 丁」等于让他自己去拆；
    而这个"已经拆过"的事实正是归因确定与否的分界线。
    """
    with pytest.raises(UnmappedCategoryError) as excinfo:
        _run(
            [["丙 丁"]],
            {"甲": "一"},
            "空",
            delimiters=[" "],
            value_rank=[{"value": "甲", "rank": 1}],
        )
    assert excinfo.value.values == ["丙", "丁"]

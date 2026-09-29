"""参数化列映射的确定性计算层（方案 §3，阶段 4）。

三条通道共用同一实现：给定**源区域的值矩阵** + 映射表 + 空白策略，
返回目标矩阵与单个白名单动作。本模块**不读写工作簿、不碰宿主 API**，
纯函数，因此 Office / WPS / 文件夹三端拿到的映射结果必然一致 ——
这正是"两个平台呈现相同行为"的验收前提。

两条刻意的"拒绝"而非"兜底"：

1. 来源值不在映射表内 → 抛 `UnmappedCategoryError`，**不猜**。
   猜测在单元格层面无法被用户察觉，而错误的分类会静默流入下游统计。
2. 源/目标区域在同一工作表上重叠 → 抛 `OverlappingRangeError`。
   写入是逐格落地的，重叠时"读源"与"写目标"的先后顺序决定结果，
   而这个顺序在三条通道上并不一致，没有可承诺的语义。

空白与未匹配是**两件不同的事**：空白是"源本来没有值"，走 `blank_label`；
未匹配是"源有值但映射表没覆盖"，是调用方的映射表不完整。
若调用方确实想把某个空白样式的值（如 `""`）映射成别的标签，
在 `mapping` 里显式写这一项即可 —— 显式项优先于空白策略。
"""

from __future__ import annotations

import re
from typing import Any

__all__ = [
    "ColumnMappingError",
    "OverlappingRangeError",
    "ShapeMismatchError",
    "UnmappedCategoryError",
    "build_column_mapping_actions",
    "map_column_values",
    "validate_column_mapping_shape",
]


class ColumnMappingError(ValueError):
    """列映射计算层的基类错误。"""


class UnmappedCategoryError(ColumnMappingError):
    """来源值未命中映射表。携带全部未命中值，便于一次补全映射表。"""

    def __init__(self, values: list[Any]) -> None:
        self.values = values
        rendered = "、".join(repr(value) for value in values)
        super().__init__(
            f"来源值 {rendered} 不在映射表内，已拒绝猜测。请补全映射表后重试。"
        )


class ShapeMismatchError(ColumnMappingError):
    """源值矩阵与目标区域的形状不一致。"""


class OverlappingRangeError(ColumnMappingError):
    """同一工作表上源区域与目标区域重叠。"""


_CELL = r"\$?([A-Za-z]+)\$?(\d+)"


def _column_number(name: str) -> int:
    result = 0
    for character in name:
        result = result * 26 + ord(character.upper()) - ord("A") + 1
    return result


def _range_box(address: str) -> tuple[int, int, int, int]:
    """返回 (起始行, 起始列, 结束行, 结束列)，1 基闭区间。"""
    match = re.fullmatch(rf"{_CELL}(?::{_CELL})?", address.strip())
    if not match:
        raise ColumnMappingError(f"无法解析区域「{address}」。")
    start_column = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_column = _column_number(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if end_row < start_row or end_column < start_column:
        raise ColumnMappingError(f"区域「{address}」的起止顺序颠倒。")
    return start_row, start_column, end_row, end_column


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    return isinstance(value, str) and not value.strip()


def map_column_values(
    *,
    source_values: list[list[Any]],
    mapping: dict[str, str],
    blank_label: str,
) -> list[list[str]]:
    """逐格映射。未命中值**全部**收集后一次抛出，而不是遇到第一个就停。

    一次只报一个会让调用方反复试错；映射表通常是人工维护的，
    一次拿到完整缺口才能一轮补完。
    """
    mapped: list[list[str]] = []
    unmapped: list[Any] = []
    seen: set[str] = set()

    for row in source_values:
        mapped_row: list[str] = []
        for cell in row:
            # 显式映射项优先于空白策略：调用方若把 "" 写进 mapping，
            # 说明它想区分"空串"与"真空格"，这个意图应当被尊重。
            if isinstance(cell, str) and cell in mapping:
                mapped_row.append(mapping[cell])
                continue
            if _is_blank(cell):
                mapped_row.append(blank_label)
                continue
            key = cell if isinstance(cell, str) else str(cell)
            if key in mapping:
                mapped_row.append(mapping[key])
                continue
            if key not in seen:
                seen.add(key)
                unmapped.append(cell)
            mapped_row.append(blank_label)  # 占位，本次调用不会被返回
        mapped.append(mapped_row)

    if unmapped:
        raise UnmappedCategoryError(unmapped)
    return mapped


def validate_column_mapping_shape(
    *,
    source_sheet: str,
    target_sheet: str,
    source_range: str,
    target_range: str,
    blank_label: str,
) -> tuple[int, int]:
    """**不需要源值**就能做的那部分校验，返回 (行数, 列数)。

    单独抽出来是因为 `planningReadPolicy == "none"` 时服务端手上没有源值，
    但仍应在下发规则前拦掉明显错误 —— 否则模型要等客户端算完才知道写错了，
    而那时它已经不在循环里了。

    TS 镜像把这些检查内联在 `buildColumnMappingActions` 里，没有单独的函数：
    行为一致（同一批夹具用例覆盖），只是切分方式不同。
    """
    if not source_sheet.strip() or not target_sheet.strip():
        raise ColumnMappingError("源或目标工作表名为空。")
    if not blank_label:
        raise ColumnMappingError("空白标签为空，无法确定空白单元格应写入什么。")

    source_box = _range_box(source_range)
    target_box = _range_box(target_range)

    source_rows = source_box[2] - source_box[0] + 1
    source_columns = source_box[3] - source_box[1] + 1
    target_rows = target_box[2] - target_box[0] + 1
    target_columns = target_box[3] - target_box[1] + 1

    if (target_rows, target_columns) != (source_rows, source_columns):
        raise ShapeMismatchError(
            f"目标区域「{target_range}」（{target_rows}×{target_columns}）"
            f"与源区域「{source_range}」（{source_rows}×{source_columns}）形状不一致。"
        )

    if source_sheet == target_sheet:
        rows_overlap = source_box[0] <= target_box[2] and target_box[0] <= source_box[2]
        columns_overlap = (
            source_box[1] <= target_box[3] and target_box[1] <= source_box[3]
        )
        if rows_overlap and columns_overlap:
            raise OverlappingRangeError(
                f"源区域「{source_range}」与目标区域「{target_range}」在同一工作表上重叠，"
                "读写先后顺序在各通道不一致，已拒绝。"
            )

    return source_rows, source_columns


def build_column_mapping_actions(
    *,
    source_sheet: str,
    target_sheet: str,
    source_range: str,
    target_range: str,
    source_values: list[list[Any]],
    mapping: dict[str, str],
    blank_label: str,
) -> list[dict[str, Any]]:
    """生成单个 `writeValues` 动作。

    只产出一个动作，不拆分：目标是一块连续区域，拆开只会让验收判定
    多出"部分动作成功"的中间态，而这个中间态在三条通道上恢复行为不同。
    """
    source_rows, source_columns = validate_column_mapping_shape(
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        source_range=source_range,
        target_range=target_range,
        blank_label=blank_label,
    )

    if len(source_values) != source_rows or any(
        len(row) != source_columns for row in source_values
    ):
        raise ShapeMismatchError(
            f"源值矩阵形状与区域「{source_range}」（{source_rows}×{source_columns}）不一致。"
        )

    values = map_column_values(
        source_values=source_values, mapping=mapping, blank_label=blank_label
    )
    return [
        {
            "type": "writeValues",
            "sheet": target_sheet,
            "range": target_range,
            "values": values,
        }
    ]

"""原地撤销工作表删除：把归档里的表放回**当前打开的原工作簿**。

这个模块只认宿主对象（`Worksheet` / `Workbook` 的 COM 形状），不认 COM、不认路径、
不认 HTTP。理由：恢复的正确性判据必须能被测试替身逐条驱动，而不是只能靠真机试。

为什么用 `Worksheet.Copy` 而不是重写单元格：
- 用户要求保留**对象**——表格对象、图表、批注、合并、条件格式、数据验证、行高列宽。
  只搬 `Value` 会把它们全部丢掉，那就不是撤销，是"重新造了一张长得像的表"。
- `Copy` 是宿主自己的整表复制语义；我们随后逐项核对名称、位置、可见性与内容指纹，
  核对不过就如实报"未完成核对"，不假装成功。

为什么必须修跨表公式：删除工作表时，幸存表里 `=Sheet7!A1` 会立刻变成 `=#REF!A1`。
把 Sheet7 放回来**不会**让它自动复原。所以要用执行前归档里的公式原文，
只修"当前仍含 `#REF!` 且归档里确实引用了本次被删表"的格子 —— 其余一律不动。
"""

from __future__ import annotations

import hashlib
import json
import re
from typing import Callable, Iterable, Sequence

# 与 Excel / ET 一致的可见性常量。
VISIBLE_SHEET = -1


class RestoreConflict(Exception):
    """恢复前就发现无法安全执行（同名表已存在等）。这不是"失败"，是"要用户决定"。"""


def _items(collection) -> list:
    return [collection.Item(index) for index in range(1, int(collection.Count) + 1)]


def sheet_names(book) -> list[str]:
    return [str(sheet.Name) for sheet in _items(book.Worksheets)]


def _sheet_by_name(book, name: str):
    for sheet in _items(book.Worksheets):
        if str(sheet.Name) == name:
            return sheet
    return None


def _normalise_matrix(value) -> tuple[tuple[str, ...], ...]:
    """把 `UsedRange.Formula` 归一成字符串矩阵。

    COM 在单格区域返回标量、多格返回二维元组，WPS 有时返回 list；这里统一形状，
    否则"内容是否一致"的比对会被容器差异干扰。
    """
    if value is None:
        return (("",),)
    if isinstance(value, str):
        return ((value,),)
    rows: list[tuple[str, ...]] = []
    for row in value:
        if isinstance(row, str):
            rows.append((row,))
            continue
        rows.append(tuple("" if cell is None else str(cell) for cell in row))
    return tuple(rows)


def _matrix_digest(matrix: tuple[tuple[str, ...], ...]) -> str:
    return hashlib.sha256(json.dumps(matrix, ensure_ascii=False).encode("utf-8")).hexdigest()


def _used(book_sheet) -> tuple[int, int, int, int]:
    """(行数, 列数, 左上行, 左上列)；空表按 1x1 处理。"""
    used = book_sheet.UsedRange
    return (max(int(used.Rows.Count), 1), max(int(used.Columns.Count), 1),
            int(used.Row), int(used.Column))


def _mentions(formula: str, names: Iterable[str]) -> bool:
    """公式是否引用了这些工作表（`Sheet7!` 或 `'Sheet 7'!`，不区分大小写）。"""
    for name in names:
        if re.search(r"(?:'" + re.escape(name) + r"'|" + re.escape(name) + r")!", formula, re.IGNORECASE):
            return True
    return False


def _anchor(current: Sequence[str], archive_order: Sequence[str], name: str) -> tuple[str, str | None]:
    """把表放回**原来的相对位置**：优先贴住归档里它前面最近、且现在还活着的那张表。

    只有在前面一张都不剩时才改贴后面最近的一张，再不行就追加到末尾。
    这样"删了中间三张、其余没动"的常见情形会精确回到原位。
    """
    index = list(archive_order).index(name)
    for previous in reversed(archive_order[:index]):
        if previous in current:
            return ("after", previous)
    for following in archive_order[index + 1:]:
        if following in current:
            return ("before", following)
    return ("end", None)


def _copy_sheet(archive_sheet, anchor_mode: str, anchor_name: str | None, book):
    target = _sheet_by_name(book, anchor_name) if anchor_name else None
    if anchor_mode == "after" and target is not None:
        archive_sheet.Copy(None, target)
    elif anchor_mode == "before" and target is not None:
        archive_sheet.Copy(target, None)
    else:
        archive_sheet.Copy(None, _items(book.Worksheets)[-1])


def _restore_one(book, archive, name: str, archive_order: list[str], current: list[str],
                 max_cells: int) -> tuple[bool, str | None]:
    """放回一张表并核对。返回 (内容是否已核对, 未核对原因)。"""
    archive_sheet = _sheet_by_name(archive, name)
    if archive_sheet is None:
        raise ValueError(f"归档中找不到工作表「{name}」，未执行撤销。")
    mode, anchor_name = _anchor(current, archive_order, name)
    expected = list(current)
    position = len(expected) if mode == "end" else (
        expected.index(anchor_name) + 1 if mode == "after" else expected.index(anchor_name))
    expected.insert(position, name)
    _copy_sheet(archive_sheet, mode, anchor_name, book)
    # 位置与名字是**硬判据**：对不上就说明宿主没有按我们要求的位置放，
    # 后续的指纹核对再准也没有意义。
    if sheet_names(book) != expected:
        raise ValueError(f"工作表「{name}」未能放回原位置，已停止撤销。")
    restored = _sheet_by_name(book, name)
    if restored is None:
        raise ValueError(f"工作表「{name}」复制结果无法读取，未继续。")
    restored.Visible = int(archive_sheet.Visible)
    rows, columns, _, _ = _used(archive_sheet)
    if rows * columns > max_cells:
        return (False, f"「{name}」内容核对超出本次限额（{rows * columns} 格），未逐格比对")
    if _matrix_digest(_normalise_matrix(archive_sheet.UsedRange.Formula)) != \
            _matrix_digest(_normalise_matrix(restored.UsedRange.Formula)):
        raise ValueError(f"工作表「{name}」内容与执行前归档不一致，未继续撤销。")
    if int(restored.Visible) != int(archive_sheet.Visible):
        raise ValueError(f"工作表「{name}」可见性与执行前不一致，未继续撤销。")
    return (True, None)


def _repair_formula_cell(live_sheet, row: int, column: int, expected: str) -> bool:
    cell = live_sheet.Cells(row, column)
    current = cell.Formula
    if isinstance(current, str) and "#REF!" in current and current != expected:
        cell.Formula = expected
        return True
    return False


def _repair_formulas(book, archive, survivors: Sequence[str], deleted: Sequence[str],
                     max_cells: int) -> tuple[int, list[str]]:
    """把被本次删除打断的跨表公式按归档原文修回来。

    只碰同时满足两条的格子：归档里这条公式**引用了本次删除的表**，且现场那条公式
    **仍然含 `#REF!`**。用户事后自己改过（比如换成了别的公式、或手写了新引用）的格子
    不含 `#REF!`，一律不动。
    """
    repaired = 0
    limitations: list[str] = []
    for name in survivors:
        live_sheet = _sheet_by_name(book, name)
        archive_sheet = _sheet_by_name(archive, name)
        if live_sheet is None or archive_sheet is None:
            continue
        rows, columns, top, left = _used(archive_sheet)
        if rows * columns > max_cells:
            limitations.append(f"「{name}」的跨表公式核对超出限额（{rows * columns} 格），未自动修复")
            continue
        matrix = _normalise_matrix(archive_sheet.UsedRange.Formula)
        for row_index, row in enumerate(matrix):
            for column_index, formula in enumerate(row):
                if not isinstance(formula, str) or not formula.startswith("="):
                    continue
                if not _mentions(formula, deleted):
                    continue
                if _repair_formula_cell(live_sheet, top + row_index, left + column_index, formula):
                    repaired += 1
    return repaired, limitations


def _repair_names(book, archive, deleted: Sequence[str]) -> tuple[int, list[str]]:
    """工作簿级名称：`=SUM(Sheet7!A1:A9)` 删表后会变成 `=#REF!`，同样按原文修回。"""
    repaired = 0
    limitations: list[str] = []
    try:
        names = _items(archive.Names)
    except Exception:  # 宿主不提供名称集合时如实记录，不假装修过了
        return 0, ["宿主未提供工作簿级名称集合，本次未核对名称定义"]
    for item in names:
        try:
            expected = str(item.RefersTo)
        except Exception:
            continue
        if not expected or not _mentions(expected, deleted):
            continue
        try:
            live = book.Names.Item(str(item.Name))
            current = str(live.RefersTo)
        except Exception:
            limitations.append(f"名称「{item.Name}」在当前工作簿中不存在，未改动")
            continue
        if "#REF!" in current and current != expected:
            try:
                live.RefersTo = expected
                repaired += 1
            except Exception:
                limitations.append(f"名称「{item.Name}」无法写回，未改动")
    return repaired, limitations


def restore_in_place(*, book, archive, deleted: Sequence[str], archive_order: Sequence[str],
                     already_restored: Sequence[str] = (), max_cells: int = 200_000,
                     before_sheet: Callable[[str], None] | None = None,
                     on_sheet: Callable[[str], None] | None = None) -> dict:
    """把 `deleted` 里的表按归档顺序放回 `book`，并修回被删除打断的引用。

    `already_restored` 是**上一次中止后已经放回去的表**：跳过它们，也不因此报冲突 ——
    否则"重试"永远无法手工收尾。除此之外，同名表现存即视为冲突（用户后来新建的），
    绝不覆盖。
    """
    deleted = [name for name in archive_order if name in set(deleted)]
    if not deleted:
        raise ValueError("本次没有可撤销的删除记录。")
    done = [name for name in deleted if name in set(already_restored)]
    pending = [name for name in deleted if name not in set(already_restored)]

    live = sheet_names(book)
    conflicts = [name for name in pending if name in live]
    if conflicts:
        raise RestoreConflict(
            "以下工作表在当前工作簿中已存在，未覆盖：" + "、".join(conflicts) +
            "。请先移走或重命名同名表，再重试撤销。")

    missing = [name for name in pending if name not in list(archive_order)]
    if missing:
        raise ValueError("归档记录与待恢复工作表不一致，未执行撤销。")

    current = list(live)
    restored: list[str] = list(done)
    unverified: list[str] = []
    limitations: list[str] = []
    for name in pending:
        if before_sheet is not None:
            before_sheet(name)
        verified, reason = _restore_one(book, archive, name, list(archive_order), current, max_cells)
        current = sheet_names(book)
        restored.append(name)
        if not verified and reason:
            unverified.append(name)
            limitations.append(reason)
        if on_sheet is not None:
            on_sheet(name)

    survivors = [name for name in archive_order if name not in set(deleted)]
    repaired_formulas, formula_limits = _repair_formulas(book, archive, survivors, deleted, max_cells)
    repaired_names, name_limits = _repair_names(book, archive, deleted)
    limitations.extend(formula_limits)
    limitations.extend(name_limits)

    return {
        "status": "completed" if not unverified else "partial",
        "restoredSheets": restored,
        "failedSheets": [],
        "unverifiedSheets": unverified,
        "repairedFormulas": repaired_formulas,
        "repairedNames": repaired_names,
        "limitations": limitations,
    }

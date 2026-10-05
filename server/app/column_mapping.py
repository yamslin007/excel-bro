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
    "NeedsClarificationError",
    "OverlappingRangeError",
    "ShapeMismatchError",
    "UnmappedCategoryError",
    "BlankPolicyNeededError",
    "BlankSplitLimitError",
    "DEFAULT_MAX_MAPPING_ACTIONS",
    "blank_policy_from_legacy",
    "build_column_mapping_actions",
    "map_column_values",
    "map_column_values_with_outcomes",
    "split_target_actions",
    "validate_column_mapping_shape",
]


class ColumnMappingError(ValueError):
    """列映射计算层的基类错误。"""


class UnmappedCategoryError(ColumnMappingError):
    """来源值未命中映射表，且**能确定**归因是「类别映射缺口」。

    典型情形：已声明分隔符把它拆开了（或它本来就是一个整值），拆出的项里
    有映射表没覆盖的类别。补救动作单一 —— **补类别映射**，可补全后重试。

    与 `NeedsClarificationError` 刻意分开：那个是「该补映射还是该补分隔符
    说不准」，补救动作有两个可能，必须让用户选。合成一个错误会让用户
    在两件事之间瞎猜，而猜错的那一半会把分隔符问题当成缺失类别补掉。
    """

    def __init__(self, values: list[Any]) -> None:
        self.values = values
        rendered = "、".join(repr(value) for value in values)
        super().__init__(
            f"来源值 {rendered} 不在映射表内，已拒绝猜测。请补全映射表后重试。"
        )


class NeedsClarificationError(ColumnMappingError):
    """来源格无法在不猜测的前提下归类，而**归因本身也不确定**。

    情形：整格未命中、且像是用了某种没被声明为可切分的写法。此时既可能是
    「用错了分隔符」，也可能是「这个整格本身就缺映射」—— 实现**不判断是
    哪一个**，只把「这格处理不了」报上去。

    与 `UnmappedCategoryError` 的区别是**补救动作的数量**：这个有两种可能
    （补分隔符 / 补映射），必须让用户选。
    """

    def __init__(self, cells: list[Any]) -> None:
        self.cells = cells
        # 文案里**不嵌入原文**：这个 message 会经 `ColumnMappingClarification.message`
        # 下发到客户端，而原文属于本端数据。要展示原文的界面有自己的源值，
        # 不需要（也不该）从服务端把它带回来。
        super().__init__(
            f"有 {len(cells)} 个来源格无法归类：可能是用了未声明的分隔符，"
            "也可能是映射表缺少这个类别。已拒绝猜测。"
        )


class ShapeMismatchError(ColumnMappingError):
    """源值矩阵与目标区域的形状不一致。"""


class OverlappingRangeError(ColumnMappingError):
    """同一工作表上源区域与目标区域重叠。"""


class BlankPolicyNeededError(ColumnMappingError):
    """源区域**真的有空白格**，而空白怎么处理还没定。

    这是**待澄清**，不是错误：实现不替用户挑一个默认标签（历史上那正是
    空白被一律写成「/」或「未分类」的来源）。调用方应把它翻成一张带四个
    选项的澄清卡，而不是把 message 原样当报错文案丢给用户。

    反过来：**源区域一格空白都没有时不该抛**。没读到空白就先问「空白怎么
    办」，用户只能凭空回答一个不存在的问题。
    """

    def __init__(self, blank_count: int) -> None:
        self.blank_count = blank_count
        super().__init__(
            f"源区域有 {blank_count} 个空白格，但还没有确定空白应当怎么处理，"
            "已拒绝替你选一个默认标签。"
        )


class BlankSplitLimitError(ColumnMappingError):
    """需要跳过的空白位置太多，一次预览无法用有限个连续写入段表达。

    宁可说清楚**为什么表达不了**，也不把「保留原值」偷偷降级成「写个空值」
    —— 那会让用户以为自己选的是保留，目标格的公式与格式却被改掉了。
    """

    def __init__(self, required_actions: int, max_actions: int) -> None:
        self.required_actions = required_actions
        self.max_actions = max_actions
        super().__init__(
            f"按当前空白分布需要拆成 {required_actions} 段写入，"
            f"超过单次预览的上限 {max_actions} 段。"
            "请改用「写入指定值」或缩小来源区域。"
        )


DEFAULT_MAX_MAPPING_ACTIONS = 64

# 「保留目标原值」靠把位置排除出写入范围实现，位置零散时目标区域必须拆成
# 多段连续区间。段数必须有上限：预览、验收与撤销都不可控。
_BLANK_PLACEHOLDER = ""


def blank_policy_from_legacy(blank_label: str | None) -> dict[str, Any]:
    """旧协议 `blank_label` → 结构化空白策略。

    非空串 = `{"kind": "value", ...}`，语义与旧版**逐格一致**，不改变任何
    一条既有计划的写法。旧协议里空串本来就被拒绝，所以折成 `undecided`
    不会影响任何曾经合法的旧计划。
    """
    if isinstance(blank_label, str) and blank_label:
        return {"kind": "value", "value": blank_label}
    return {"kind": "undecided"}


def _blank_cell_outcome(policy: dict[str, Any]) -> dict[str, Any]:
    """单个空白格在给定策略下的落地方式：`skip` / `clear` / `write`。"""
    kind = policy.get("kind")
    if kind == "keep":
        return {"kind": "skip"}
    if kind == "clear":
        return {"kind": "clear"}
    if kind == "value":
        return {"kind": "write", "value": policy.get("value", "")}
    if kind == "undecided":
        raise ColumnMappingError("空白策略尚未确定，不能落地任何写入动作。")
    raise ColumnMappingError(f"无法识别的空白策略 kind：{kind}")


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


def map_column_values_with_outcomes(
    *,
    source_values: list[list[Any]],
    mapping: dict[str, str],
    blank_policy: dict[str, Any] | None = None,
    blank_label: str | None = None,
    delimiters: list[str] | None = None,
    value_rank: list[dict[str, Any]] | None = None,
    undeclared_separators: list[str] | None = None,
) -> tuple[list[list[str]], list[list[dict[str, Any]]]]:
    """逐格映射，并给出每一格的**落地方式**。

    之所以要把 `outcomes` 一并算出来：仅凭值矩阵无法区分「写一个空字符串」
    与「这个位置不写入」—— 两者在矩阵里长得一模一样，而它们的语义完全
    相反。让调用方拿着值矩阵去猜，历史上正是「保留原值被实现成写空值」
    的根源。
    """

    policy = blank_policy or blank_policy_from_legacy(blank_label)
    return _map_column_cells(
        source_values=source_values,
        mapping=mapping,
        policy=policy,
        delimiters=delimiters,
        value_rank=value_rank,
        undeclared_separators=undeclared_separators,
    )


def map_column_values(
    *,
    source_values: list[list[Any]],
    mapping: dict[str, str],
    blank_policy: dict[str, Any] | None = None,
    blank_label: str | None = None,
    delimiters: list[str] | None = None,
    value_rank: list[dict[str, Any]] | None = None,
    undeclared_separators: list[str] | None = None,
) -> list[list[str]]:
    """只取目标矩阵。新调用方一律用 `map_column_values_with_outcomes`。"""
    values, _ = map_column_values_with_outcomes(
        source_values=source_values,
        mapping=mapping,
        blank_policy=blank_policy,
        blank_label=blank_label,
        delimiters=delimiters,
        value_rank=value_rank,
        undeclared_separators=undeclared_separators,
    )
    return values


def _map_column_cells(
    *,
    source_values: list[list[Any]],
    mapping: dict[str, str],
    policy: dict[str, Any],
    delimiters: list[str] | None = None,
    value_rank: list[dict[str, Any]] | None = None,
    undeclared_separators: list[str] | None = None,
) -> tuple[list[list[str]], list[list[dict[str, Any]]]]:
    """逐格映射，多值语义由**规则数据**驱动（`delimiters` / `value_rank`）。

    判断顺序刻意是**整格精确 → 空白 → 按已声明分隔符拆分 → 未覆盖**：

    1. 先试整格精确匹配。这是最高优先级：调用方把整格原文写进 mapping，
       就是在说「这一格整体算一个类别」，拆分不该推翻它。
    2. 空白走 `blank_label`，不参与拆分。
    3. 有 `delimiters` 才拆，取**出现位置最靠前**的那个切分点。
       位置优先而非声明序优先：声明序表达的是**优先级**（谁更重要），
       不是**切分方式**（按什么切）。用它决定切分会让
       `["；", ","]` 这类表达反直觉。
    4. 拆开后逐项查映射表：

       - 有项未覆盖 → `UnmappedCategoryError`，**按项**收集缺失类别
         （不按整格收 —— 用户要的是「把这项补进映射表」，给整格原文
         等于让他自己去拆）。归因**确定**是类别缺口。
       - 全部命中 → 按 `value_rank` 取最高级。

    5. 整格里找不到任何已声明分隔符、又没命中：

       - 含 `undeclared_separators` 里的字符 → `NeedsClarificationError`。
         归因**不确定**（可能分隔符没声明全，也可能这格本身缺映射），
         两种补救动作不同，不替用户选。
       - 看不出任何分隔形式 → 归因确定，就是映射表缺这个类别。

    一次只报一个会让调用方反复试错；映射表通常人工维护，一轮拿到完整缺口
    才能一次补完，所以缺失项**全部**收集后一次抛出。
    """
    delimiters = delimiters or []
    value_rank = value_rank or []
    undeclared_separators = undeclared_separators or []

    # 无显式 rank 时用 mapping 的声明顺序兜底。Python 3.7+ dict 保序，
    # 所以这里的下标就是「规则声明的先后」。
    explicit_ranks = {entry["value"]: entry["rank"] for entry in value_rank}
    mapping_order = {key: index for index, key in enumerate(mapping)}

    mapped: list[list[str]] = []
    outcome_rows: list[list[dict[str, Any]]] = []
    unmatched: list[Any] = []
    unresolved: list[Any] = []
    seen_unmatched: set[str] = set()
    seen_unresolved: set[str] = set()

    for row in source_values:
        mapped_row: list[str] = []
        outcome_row: list[dict[str, Any]] = []

        def push_mapped(value: str) -> None:
            mapped_row.append(value)
            outcome_row.append({"kind": "write", "value": value})

        def push_placeholder() -> None:
            """未命中位置的占位：这些格子必然让本函数抛错，矩阵不会返回。"""
            mapped_row.append(_BLANK_PLACEHOLDER)
            outcome_row.append({"kind": "write", "value": _BLANK_PLACEHOLDER})

        for cell in row:
            # 1. 整格精确匹配优先于一切。
            if isinstance(cell, str) and cell in mapping:
                push_mapped(mapping[cell])
                continue
            # 2. 空白：源本来没有值，按**空白策略**落地，不参与拆分。
            #
            #    策略未定时先记成 `skip` 占位：本函数不抛错，由
            #    `build_column_mapping_actions` 在确认"确实有空白"之后
            #    再决定是否升级为一次澄清 —— 没读到空白就不该拦。
            if _is_blank(cell):
                if policy.get("kind") == "undecided":
                    outcome = {"kind": "skip"}
                else:
                    outcome = _blank_cell_outcome(policy)
                mapped_row.append(
                    outcome["value"] if outcome["kind"] == "write" else _BLANK_PLACEHOLDER
                )
                outcome_row.append(outcome)
                continue

            key = cell if isinstance(cell, str) else str(cell)

            # 3. 没有分隔符声明 = 一个格子只当一个值，与扩展前逐格一致。
            if not delimiters:
                if key in mapping:
                    push_mapped(mapping[key])
                    continue
                if key not in seen_unmatched:
                    seen_unmatched.add(key)
                    unmatched.append(cell)
                push_placeholder()
                continue

            # 4. 用已声明的分隔符切分：取出现位置最靠前的那个。
            delimiter: str | None = None
            delimiter_index = len(key) + 1
            for candidate in delimiters:
                if not candidate:
                    continue
                index = key.find(candidate)
                if 0 <= index < delimiter_index:
                    delimiter = candidate
                    delimiter_index = index

            if delimiter is None:
                # 4a. 整格里没有任何已声明的分隔符。
                if key in mapping:
                    push_mapped(mapping[key])
                    continue
                # 含规则明示为分隔符写法、却未被声明为可切分的字符：
                # 归因不确定，两种补救动作不同，报上去让用户选。
                if any(
                    candidate and candidate in key
                    for candidate in undeclared_separators
                ):
                    if key not in seen_unresolved:
                        seen_unresolved.add(key)
                        unresolved.append(cell)
                    push_placeholder()
                    continue
                # 看不出任何分隔形式：归因确定 —— 映射表缺这个类别。
                if key not in seen_unmatched:
                    seen_unmatched.add(key)
                    unmatched.append(cell)
                push_placeholder()
                continue

            # 4b. 用已声明的分隔符拆开。
            items = [item.strip() for item in key.split(delimiter)]
            items = [item for item in items if item]

            # 整格以分隔符开头/结尾，拆完一项不剩：当整格未覆盖处理。
            if not items:
                if key not in seen_unmatched:
                    seen_unmatched.add(key)
                    unmatched.append(cell)
                push_placeholder()
                continue

            # 5. 拆分已按用户声明的分隔符完成，归因**确定**是类别缺口。
            missing = [item for item in items if item not in mapping]
            if missing:
                for item in missing:
                    if item not in seen_unmatched:
                        seen_unmatched.add(item)
                        unmatched.append(item)
                push_placeholder()
                continue

            # 6. 取规则指定的最高级：显式 rank 先看，没有就看声明顺序。
            winner = _pick_highest_rank(items, explicit_ranks, mapping_order)
            push_mapped(mapping[winner])
        mapped.append(mapped_row)
        outcome_rows.append(outcome_row)

    # 先说「归因不确定」的那类：它需要用户做选择，信息量更大。
    if unresolved:
        raise NeedsClarificationError(unresolved)
    if unmatched:
        raise UnmappedCategoryError(unmatched)
    return mapped, outcome_rows


def _pick_highest_rank(
    items: list[str],
    explicit_ranks: dict[str, int],
    mapping_order: dict[str, int],
) -> str:
    """多值取最高级。比较的是「优先级」，不是数值大小。

    `value_rank` 里 rank 越大越优先；没有 rank 的项按 `mapping` 声明顺序决定。
    只给了一边的等级时，另一边的等级按「最低」处理：显式等级是调用方的
    明确声明，未声明的项不该反过来压过它。
    """
    winner = items[0]
    for item in items:
        if _compare_rank(item, winner, explicit_ranks, mapping_order) > 0:
            winner = item
    return winner


def _compare_rank(
    left: str,
    right: str,
    explicit_ranks: dict[str, int],
    mapping_order: dict[str, int],
) -> int:
    if left in explicit_ranks or right in explicit_ranks:
        left_value = explicit_ranks.get(left)
        right_value = explicit_ranks.get(right)
        left_value = float("-inf") if left_value is None else left_value
        right_value = float("-inf") if right_value is None else right_value
        if left_value != right_value:
            return 1 if left_value > right_value else -1
    left_order = mapping_order.get(left, len(mapping_order))
    right_order = mapping_order.get(right, len(mapping_order))
    if left_order == right_order:
        return 0
    # 声明越靠前越优先，所以顺序号越小越优先 —— 返回反号。
    return 1 if left_order < right_order else -1


def validate_column_mapping_shape(
    *,
    source_sheet: str,
    target_sheet: str,
    source_range: str,
    target_range: str,
    blank_label: str | None = None,
) -> tuple[int, int]:
    """**不需要源值**就能做的那部分校验，返回 (行数, 列数)。

    **不再**要求先给出空白标签：那是"还没读源数据就先问空白怎么办"的根源。
    空白策略未定不是错误，只有真的读到空白时才需要决定。


    单独抽出来是因为 `planningReadPolicy == "none"` 时服务端手上没有源值，
    但仍应在下发规则前拦掉明显错误 —— 否则模型要等客户端算完才知道写错了，
    而那时它已经不在循环里了。

    TS 镜像把这些检查内联在 `buildColumnMappingActions` 里，没有单独的函数：
    行为一致（同一批夹具用例覆盖），只是切分方式不同。
    """
    if not source_sheet.strip() or not target_sheet.strip():
        raise ColumnMappingError("源或目标工作表名为空。")

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


def _column_letter(index: int) -> str:
    """1 基列号 → 列字母（1→A）。与 `packages/core/src/range.ts` 同算法。"""
    letters = ""
    n = index - 1
    while True:
        letters = chr(65 + n % 26) + letters
        n = n // 26 - 1
        if n < 0:
            return letters


def split_target_actions(
    *,
    target_sheet: str,
    target_box: tuple[int, int, int, int],
    target_range: str,
    values: list[list[str]],
    outcomes: list[list[dict[str, Any]]],
    max_actions: int | None = None,
) -> list[dict[str, Any]]:
    """目标矩阵 → 写入动作序列。

    **没有任何需要跳过或清空的位置时，只产出一个覆盖整块区域的动作** ——
    与扩展前逐格一致，也是绝大多数映射的真实形态。

    一旦有位置要跳过（保留原值）或清空，整块连续区域就不再成立：把它整块
    写进去等于把"跳过"写成"写了个空值"，目标格的公式与格式当场被替换掉。
    这时按**列内连续段**拆开，被跳过的位置**不出现在任何动作的 range 里**。
    """
    limit = max_actions or DEFAULT_MAX_MAPPING_ACTIONS
    start_row, start_column, end_row, end_column = target_box
    row_count = end_row - start_row + 1
    column_count = end_column - start_column + 1

    needs_split = any(
        outcome["kind"] != "write" for row in outcomes for outcome in row
    )
    if not needs_split:
        return [
            {
                "type": "writeValues",
                "sheet": target_sheet,
                "range": target_range,
                "values": values,
            }
        ]

    actions: list[dict[str, Any]] = []
    for column_offset in range(column_count):
        letter = _column_letter(start_column + column_offset)
        run_start: int | None = None
        run_kind: str | None = None

        def flush(end_offset: int) -> None:
            nonlocal run_start, run_kind
            if run_start is None or run_kind is None:
                return
            range_text = (
                f"{letter}{start_row + run_start}:{letter}{start_row + end_offset}"
            )
            if run_kind == "clear":
                actions.append(
                    {
                        "type": "clearRange",
                        "sheet": target_sheet,
                        "range": range_text,
                        "applyTo": "contents",
                    }
                )
            else:
                actions.append(
                    {
                        "type": "writeValues",
                        "sheet": target_sheet,
                        "range": range_text,
                        "values": [
                            [values[offset][column_offset]]
                            for offset in range(run_start, end_offset + 1)
                        ],
                    }
                )
            run_start = None
            run_kind = None

        for row_offset in range(row_count):
            kind = outcomes[row_offset][column_offset]["kind"]
            if kind == "skip":
                flush(row_offset - 1)
                continue
            if run_kind != kind:
                flush(row_offset - 1)
                run_start = row_offset
                run_kind = kind
        flush(row_count - 1)

    if len(actions) > limit:
        raise BlankSplitLimitError(len(actions), limit)
    return actions


def build_column_mapping_actions(
    *,
    source_sheet: str,
    target_sheet: str,
    source_range: str,
    target_range: str,
    source_values: list[list[Any]],
    mapping: dict[str, str],
    blank_policy: dict[str, Any] | None = None,
    blank_label: str | None = None,
    delimiters: list[str] | None = None,
    value_rank: list[dict[str, Any]] | None = None,
    undeclared_separators: list[str] | None = None,
    max_split_actions: int | None = None,
) -> list[dict[str, Any]]:
    """生成列映射的写入动作序列。

    不需要跳过/清空位置时只产出一个 `writeValues`（与扩展前一致）；
    需要时按列内连续段拆分，被跳过的位置不进入任何 range。
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

    policy = blank_policy or blank_policy_from_legacy(blank_label)
    values, outcomes = map_column_values_with_outcomes(
        source_values=source_values,
        mapping=mapping,
        blank_policy=policy,
        delimiters=delimiters,
        value_rank=value_rank,
        undeclared_separators=undeclared_separators,
    )

    # 策略未定 + **真的读到了空白** → 待澄清，不替用户挑默认标签。
    # 一格空白都没有时不触发 —— 那时"空白怎么办"是个不存在的问题。
    if policy.get("kind") == "undecided":
        blank_count = sum(
            1 for row in outcomes for outcome in row if outcome["kind"] == "skip"
        )
        if blank_count:
            raise BlankPolicyNeededError(blank_count)

    return split_target_actions(
        target_sheet=target_sheet,
        target_box=_range_box(target_range),
        target_range=target_range,
        values=values,
        outcomes=outcomes,
        max_actions=max_split_actions,
    )

"""
将 FormulaCall JSON 编译成 Excel 公式文本。

编译是纯确定性代码：模型只负责选函数和参数，语法正确性由这里保证。
"""

from __future__ import annotations

import math
import re

from .schema import (
    BooleanLiteral,
    CellRef,
    FunctionCall,
    NumberLiteral,
    RangeRef,
    TextLiteral,
)


def compile_formula(call: FunctionCall) -> str:
    """将函数调用树编译成 Excel 公式文本（带前导 =）。"""
    formula_text = _compile_node(call)
    return f"={formula_text}"


def _compile_node(node) -> str:
    """递归编译节点。"""
    if isinstance(node, FunctionCall):
        return _compile_function(node)
    if isinstance(node, CellRef):
        return node.address
    if isinstance(node, RangeRef):
        return _compile_range(node)
    if isinstance(node, TextLiteral):
        # Excel 文本字面量用双引号包裹，内部双引号转义成两个。
        escaped = node.value.replace('"', '""')
        return f'"{escaped}"'
    if isinstance(node, NumberLiteral):
        return _format_number(node.value)
    if isinstance(node, BooleanLiteral):
        return "TRUE" if node.value else "FALSE"
    raise ValueError(f"未知节点类型: {type(node)}")


def _format_number(value: float) -> str:
    """数字格式化：整数值不带 .0，避免公式里出现 A2>0.0 这类冗余写法。"""
    if not math.isfinite(value):
        raise ValueError(f"公式不支持非有限数字: {value}")
    if value == int(value):
        return str(int(value))
    return repr(value)


_BINARY_FUNCTIONS: dict[str, tuple[str, str]] = {
    "EQ": ("=", "等于"),
    "NE": ("<>", "不等于"),
    "GT": (">", "大于"),
    "LT": ("<", "小于"),
    "GTE": (">=", "大于等于"),
    "LTE": ("<=", "小于等于"),
    "ADD": ("+", "加法"),
    "SUBTRACT": ("-", "减法"),
    "MULTIPLY": ("*", "乘法"),
    "DIVIDE": ("/", "除法"),
}


def _compile_function(call: FunctionCall) -> str:
    """编译函数调用：二元运算符特殊处理，其余按 Excel 函数名编译。"""
    func_name = call.function.upper()

    binary = _BINARY_FUNCTIONS.get(func_name)
    if binary is not None:
        if len(call.args) != 2:
            raise ValueError(f"{func_name} 需要 2 个参数，实际 {len(call.args)} 个")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}{binary[0]}{right})"

    compiled_args = [_compile_node(arg) for arg in call.args]
    args_str = ",".join(compiled_args)
    return f"{func_name}({args_str})"


# 不需要加引号的工作表名：字母数字下划线 + 中文，且不像单元格地址。
_SAFE_SHEET_NAME = re.compile(r"^[A-Za-z0-9_\u4e00-\u9fff]+$")
_CELL_LIKE_NAME = re.compile(r"^[A-Za-z]{1,3}[0-9]+$")
# 外部工作簿文件名：允许扩展名中的点与常见连接符。
_SAFE_WORKBOOK_NAME = re.compile(r"^[A-Za-z0-9_.\-\u4e00-\u9fff]+$")


def _needs_quoting(name: str) -> bool:
    """名称含空格/特殊字符，或长得像单元格地址时，Excel 要求加单引号。"""
    if _CELL_LIKE_NAME.match(name):
        return True
    return not bool(_SAFE_SHEET_NAME.match(name))


def _compile_range(range_ref: RangeRef) -> str:
    """编译区域引用。"""
    if range_ref.workbook:
        workbook_ok = bool(_SAFE_WORKBOOK_NAME.match(range_ref.workbook))
        sheet_ok = not _needs_quoting(range_ref.sheet)
        sheet_ref = f"[{range_ref.workbook}]{range_ref.sheet}"
        # 文件名或表名含空格/特殊字符时，整个 [文件]表 部分要加单引号，
        # 且内部单引号翻倍转义；Excel 不允许只在表名外单独加引号。
        if not (workbook_ok and sheet_ok):
            sheet_ref = f"'{sheet_ref.replace(chr(39), chr(39) * 2)}'"
    else:
        sheet_ref = range_ref.sheet
        if _needs_quoting(sheet_ref):
            sheet_ref = f"'{sheet_ref.replace(chr(39), chr(39) * 2)}'"

    col = range_ref.col.upper()
    if range_ref.row2 is None:
        cell_range = f"${col}:${col}"  # 整列引用
    else:
        cell_range = f"${col}${range_ref.row1}:${col}${range_ref.row2}"

    return f"{sheet_ref}!{cell_range}"

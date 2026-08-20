"""
Excel 函数定义注册表。

用途：
1. 模型 Prompt（告诉模型有哪些函数可用，以及参数格式）
2. 编译器验证的参考（运算符参数个数等约束）
3. 未来扩展（增加新函数只需在这里登记）

注意：示例 JSON 必须是合法、完整的调用，不能出现 `[...]` 占位符，
否则会把错误格式教给模型。
"""

from __future__ import annotations

from typing import TypedDict


class FunctionDef(TypedDict):
    """函数定义。"""

    name: str  # 函数名
    description: str  # 功能说明
    args: list[str]  # 参数说明列表
    example: str  # 使用示例（合法 JSON 字符串）


# Excel 函数注册表（按类别组织）
EXCEL_FUNCTIONS: dict[str, list[FunctionDef]] = {
    "查找与引用": [
        {
            "name": "XLOOKUP",
            "description": "在区域中查找值并返回对应结果（Excel 365+）",
            "args": [
                "lookup_value: 要查找的值",
                "lookup_array: 查找范围",
                "return_array: 返回范围",
                "if_not_found?: 未找到时的默认值",
            ],
            "example": '{"function": "XLOOKUP", "args": [{"type": "cell", "address": "A2"}, {"type": "range", "sheet": "字典", "col": "A", "row1": 2, "row2": 100}, {"type": "range", "sheet": "字典", "col": "B", "row1": 2, "row2": 100}, {"type": "text", "value": ""}]}',
        },
        {
            "name": "VLOOKUP",
            "description": "垂直查找（兼容旧版 Excel），第四参数必须 FALSE",
            "args": [
                "lookup_value: 要查找的值",
                "table_array: 查找表区域",
                "col_index_num: 返回第几列",
                "range_lookup: FALSE=精确匹配, TRUE=近似匹配",
            ],
            "example": '{"function": "VLOOKUP", "args": [{"type": "cell", "address": "A2"}, {"type": "range", "sheet": "字典", "col": "A", "row1": 2, "row2": 100}, {"type": "number", "value": 2}, {"type": "boolean", "value": false}]}',
        },
        {
            "name": "INDEX",
            "description": "返回指定位置的值",
            "args": [
                "array: 区域",
                "row_num: 行号",
                "col_num?: 列号",
            ],
            "example": '{"function": "INDEX", "args": [{"type": "range", "sheet": "数据", "col": "B", "row1": 2, "row2": 100}, {"type": "number", "value": 5}]}',
        },
        {
            "name": "MATCH",
            "description": "返回值在区域中的位置，第三参数必须 0（精确匹配）",
            "args": [
                "lookup_value: 要查找的值",
                "lookup_array: 查找范围",
                "match_type: 0=精确, 1=小于等于, -1=大于等于",
            ],
            "example": '{"function": "MATCH", "args": [{"type": "text", "value": "产品A"}, {"type": "range", "sheet": "列表", "col": "A", "row1": 1, "row2": 100}, {"type": "number", "value": 0}]}',
        },
    ],
    "逻辑函数": [
        {
            "name": "IF",
            "description": "条件判断",
            "args": [
                "condition: 条件表达式",
                "value_if_true: 条件为真时的值",
                "value_if_false: 条件为假时的值",
            ],
            "example": '{"function": "IF", "args": [{"function": "GT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 100}]}, {"type": "text", "value": "合格"}, {"type": "text", "value": "不合格"}]}',
        },
        {
            "name": "AND",
            "description": "所有条件都为真",
            "args": [
                "condition1, condition2, ...: 条件表达式",
            ],
            "example": '{"function": "AND", "args": [{"function": "GT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 0}]}, {"function": "LT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 100}]}]}',
        },
        {
            "name": "OR",
            "description": "任一条件为真",
            "args": [
                "condition1, condition2, ...: 条件表达式",
            ],
            "example": '{"function": "OR", "args": [{"function": "EQ", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": "甲"}]}, {"function": "EQ", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": "乙"}]}]}',
        },
        {
            "name": "NOT",
            "description": "条件取反",
            "args": [
                "condition: 条件表达式",
            ],
            "example": '{"function": "NOT", "args": [{"function": "EQ", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": ""}]}]}',
        },
    ],
    "文本函数": [
        {
            "name": "SEARCH",
            "description": "在文本中查找子串（不区分大小写），找不到返回 #VALUE!",
            "args": [
                "find_text: 要查找的文本",
                "within_text: 被查找的文本",
                "start_num?: 起始位置",
            ],
            "example": '{"function": "SEARCH", "args": [{"type": "text", "value": "错误"}, {"type": "cell", "address": "D2"}]}',
        },
        {
            "name": "ISNUMBER",
            "description": "判断是否为数字（常与 SEARCH 配合判断是否包含关键词）",
            "args": [
                "value: 要判断的值",
            ],
            "example": '{"function": "ISNUMBER", "args": [{"function": "SEARCH", "args": [{"type": "text", "value": "错误"}, {"type": "cell", "address": "D2"}]}]}',
        },
        {
            "name": "CONCATENATE",
            "description": "连接文本",
            "args": [
                "text1, text2, ...: 要连接的文本",
            ],
            "example": '{"function": "CONCATENATE", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": "-"}, {"type": "cell", "address": "B2"}]}',
        },
        {
            "name": "LEFT",
            "description": "提取左侧字符",
            "args": [
                "text: 文本",
                "num_chars: 字符数",
            ],
            "example": '{"function": "LEFT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 3}]}',
        },
        {
            "name": "RIGHT",
            "description": "提取右侧字符",
            "args": [
                "text: 文本",
                "num_chars: 字符数",
            ],
            "example": '{"function": "RIGHT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 2}]}',
        },
        {
            "name": "MID",
            "description": "提取中间字符",
            "args": [
                "text: 文本",
                "start_num: 起始位置",
                "num_chars: 字符数",
            ],
            "example": '{"function": "MID", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 3}, {"type": "number", "value": 5}]}',
        },
    ],
    "数学与统计": [
        {
            "name": "SUM",
            "description": "求和",
            "args": [
                "number1, number2, ...: 数字或区域",
            ],
            "example": '{"function": "SUM", "args": [{"type": "range", "sheet": "数据", "col": "C", "row1": 2, "row2": 100}]}',
        },
        {
            "name": "SUMIF",
            "description": "条件求和",
            "args": [
                "range: 条件范围",
                "criteria: 条件",
                "sum_range?: 求和范围",
            ],
            "example": '{"function": "SUMIF", "args": [{"type": "range", "sheet": "数据", "col": "A", "row1": 2, "row2": 100}, {"type": "text", "value": "产品A"}, {"type": "range", "sheet": "数据", "col": "B", "row1": 2, "row2": 100}]}',
        },
        {
            "name": "SUMPRODUCT",
            "description": "数组乘积求和（可对整块区域做条件计数）",
            "args": [
                "array1, array2, ...: 数组",
            ],
            "example": '{"function": "SUMPRODUCT", "args": [{"function": "MULTIPLY", "args": [{"function": "ISNUMBER", "args": [{"function": "SEARCH", "args": [{"type": "range", "sheet": "关键词", "col": "A", "row1": 2, "row2": 20}, {"type": "cell", "address": "D2"}]}]}, {"function": "EQ", "args": [{"type": "range", "sheet": "关键词", "col": "B", "row1": 2, "row2": 20}, {"type": "text", "value": "p0"}]}]}]}',
        },
        {
            "name": "AVERAGE",
            "description": "求平均值",
            "args": [
                "number1, number2, ...: 数字或区域",
            ],
            "example": '{"function": "AVERAGE", "args": [{"type": "range", "sheet": "数据", "col": "C", "row1": 2, "row2": 100}]}',
        },
        {
            "name": "COUNT",
            "description": "计数（仅数字）",
            "args": [
                "value1, value2, ...: 值或区域",
            ],
            "example": '{"function": "COUNT", "args": [{"type": "range", "sheet": "数据", "col": "A", "row1": 2, "row2": 100}]}',
        },
        {
            "name": "COUNTA",
            "description": "计数（非空单元格）",
            "args": [
                "value1, value2, ...: 值或区域",
            ],
            "example": '{"function": "COUNTA", "args": [{"type": "range", "sheet": "数据", "col": "A", "row1": 2, "row2": 100}]}',
        },
    ],
    "比较运算符": [
        {
            "name": "EQ",
            "description": "等于（=）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "EQ", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": ""}]}',
        },
        {
            "name": "NE",
            "description": "不等于（<>）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "NE", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": ""}]}',
        },
        {
            "name": "GT",
            "description": "大于（>）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "GT", "args": [{"type": "number", "value": 5}, {"type": "number", "value": 0}]}',
        },
        {
            "name": "LT",
            "description": "小于（<）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "LT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 100}]}',
        },
        {
            "name": "GTE",
            "description": "大于等于（>=）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "GTE", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 0}]}',
        },
        {
            "name": "LTE",
            "description": "小于等于（<=）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "LTE", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 100}]}',
        },
    ],
    "算术运算符": [
        {
            "name": "ADD",
            "description": "加法（+）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "ADD", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 10}]}',
        },
        {
            "name": "SUBTRACT",
            "description": "减法（-）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "SUBTRACT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 5}]}',
        },
        {
            "name": "MULTIPLY",
            "description": "乘法（*）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "MULTIPLY", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 1.1}]}',
        },
        {
            "name": "DIVIDE",
            "description": "除法（/）",
            "args": [
                "left: 左侧值",
                "right: 右侧值",
            ],
            "example": '{"function": "DIVIDE", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 2}]}',
        },
    ],
}


def get_functions_prompt() -> str:
    """生成函数列表的 Prompt 文本（含参数说明与合法 JSON 示例）。"""
    lines = ["# 可用的 Excel 函数\n"]

    for category, functions in EXCEL_FUNCTIONS.items():
        lines.append(f"\n## {category}\n")
        for func in functions:
            lines.append(f"### {func['name']}")
            lines.append(f"{func['description']}\n")
            lines.append("参数：")
            for arg in func["args"]:
                lines.append(f"- {arg}")
            lines.append(f"\n示例：\n```json\n{func['example']}\n```\n")

    return "\n".join(lines)

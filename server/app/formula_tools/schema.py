"""
公式工具调用的数据模型定义。

模型只返回结构化的函数调用树（JSON），不直接生成公式文本；
公式文本由 compiler 从这里的模型递归编译，保证语法正确。
"""

from __future__ import annotations

from typing import Literal, Union

from pydantic import BaseModel, Field


class CellRef(BaseModel):
    """单元格引用，如 "D2"、"A1"。"""

    type: Literal["cell"] = "cell"
    address: str


class RangeRef(BaseModel):
    """区域引用：工作表 + 单列范围（支持外部工作簿）。"""

    type: Literal["range"] = "range"
    workbook: str | None = None  # None 表示当前工作簿
    sheet: str  # 工作表名
    col: str  # 列字母（如 "A"）
    row1: int  # 起始行
    row2: int | None = None  # 结束行，None 表示到底


class TextLiteral(BaseModel):
    """文本字面量。"""

    type: Literal["text"] = "text"
    value: str


class NumberLiteral(BaseModel):
    """数字字面量。"""

    type: Literal["number"] = "number"
    value: float


class BooleanLiteral(BaseModel):
    """布尔字面量。"""

    type: Literal["boolean"] = "boolean"
    value: bool


class FunctionCall(BaseModel):
    """函数调用：函数名 + 参数列表（参数可递归嵌套）。"""

    function: str  # 函数名（如 "XLOOKUP"、"IF"、"SUMPRODUCT"）
    args: list[
        Union[
            "FunctionCall",
            CellRef,
            RangeRef,
            TextLiteral,
            NumberLiteral,
            BooleanLiteral,
        ]
    ] = Field(default_factory=list)


# 支持递归定义（Pydantic v2）
FunctionCall.model_rebuild()


class FormulaToolCallResponse(BaseModel):
    """模型返回的完整响应。"""

    formula: FunctionCall  # 主公式（顶层函数调用）
    explanation: str  # 公式说明（用户可读）

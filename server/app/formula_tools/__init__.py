"""公式工具调用架构：模型返回结构化函数调用，代码编译成 Excel 公式。"""

from .schema import (
    BooleanLiteral,
    CellRef,
    FormulaToolCallResponse,
    FunctionCall,
    NumberLiteral,
    RangeRef,
    TextLiteral,
)
from .compiler import compile_formula
from .function_registry import EXCEL_FUNCTIONS, get_functions_prompt
from .model_prompt import SYSTEM_PROMPT, build_user_message

__all__ = [
    "BooleanLiteral",
    "CellRef",
    "EXCEL_FUNCTIONS",
    "FormulaToolCallResponse",
    "FunctionCall",
    "NumberLiteral",
    "RangeRef",
    "SYSTEM_PROMPT",
    "TextLiteral",
    "build_user_message",
    "compile_formula",
    "get_functions_prompt",
]

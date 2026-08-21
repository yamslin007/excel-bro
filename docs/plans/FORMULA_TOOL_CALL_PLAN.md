# /function 工具调用架构方案

## 目标

将公式生成从"模型端到端生成公式文本"改为"模型返回结构化函数调用 → 代码编译成公式"，实现：
- ✅ 语法 100% 正确（代码生成，不依赖模型）
- ✅ 支持任意 Excel 函数组合（不枚举场景）
- ✅ 可扩展（新函数只需定义 schema）
- ✅ 可验证（JSON 可检查，失败可重试）

---

## 核心架构

```
用户需求："根据 D 列问题类型返回最高错误级别"
    ↓
┌─────────────────────────────────────────────────────────────┐
│ 第一步：模型分析 → 返回结构化函数调用（JSON）                │
└─────────────────────────────────────────────────────────────┘
    ↓
{
  "function": "IF",
  "args": [
    { "function": "EQ", "args": [
      { "type": "cell", "address": "D2" },
      { "type": "text", "value": "" }
    ]},
    { "type": "text", "value": "" },
    { "function": "IF", "args": [
      { "function": "GT", "args": [
        { "function": "SUMPRODUCT", "args": [
          { "function": "MULTIPLY", "args": [
            { "function": "ISNUMBER", "args": [
              { "function": "SEARCH", "args": [
                { "type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "A", "row1": 2, "row2": 999 },
                { "type": "cell", "address": "D2" }
              ]}
            ]},
            { "function": "EQ", "args": [
              { "type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "B", "row1": 2, "row2": 999 },
              { "type": "text", "value": "p0" }
            ]}
          ]}
        ]},
        { "type": "number", "value": 0 }
      ]},
      { "type": "text", "value": "p0" },
      { "function": "IF", "args": [...嵌套 p1 和 / 的判断...] }
    ]}
  ]
}
    ↓
┌─────────────────────────────────────────────────────────────┐
│ 第二步：代码编译 JSON → Excel 公式文本                       │
└─────────────────────────────────────────────────────────────┘
    ↓
=IF(D2="","",IF(SUMPRODUCT(ISNUMBER(SEARCH([tt.xlsx]tt!$A$2:$A$999,D2))*([tt.xlsx]tt!$B$2:$B$999="p0"))>0,"p0",IF(...)))
```

---

## 文件结构

### 新增文件

```
server/app/
├── formula_tools/
│   ├── __init__.py                    # 导出主要接口
│   ├── schema.py                      # FormulaCall 数据模型定义
│   ├── compiler.py                    # JSON → 公式文本编译器
│   ├── function_registry.py           # 支持的函数定义（可扩展）
│   └── model_prompt.py                # 模型 Prompt（System + Few-shot）
└── formula_tools_integration.py       # 集成到现有 rule_generator.py
```

---

## 实施步骤

### 步骤 1：定义数据模型（schema.py）

```python
"""
公式工具调用的数据模型定义
"""
from typing import Literal, Union
from pydantic import BaseModel, Field


class CellRef(BaseModel):
    """单元格引用"""
    type: Literal["cell"] = "cell"
    address: str  # 例如 "D2", "A1"
    
    
class RangeRef(BaseModel):
    """区域引用"""
    type: Literal["range"] = "range"
    workbook: str | None = None  # None 表示当前工作簿
    sheet: str  # 工作表名
    col: str  # 列字母（如 "A"）
    row1: int  # 起始行
    row2: int | None = None  # 结束行，None 表示到底


class TextLiteral(BaseModel):
    """文本字面量"""
    type: Literal["text"] = "text"
    value: str


class NumberLiteral(BaseModel):
    """数字字面量"""
    type: Literal["number"] = "number"
    value: float


class BooleanLiteral(BaseModel):
    """布尔字面量"""
    type: Literal["boolean"] = "boolean"
    value: bool


class FunctionCall(BaseModel):
    """函数调用"""
    function: str  # 函数名（如 "XLOOKUP", "IF", "SUMPRODUCT"）
    args: list[Union[
        "FunctionCall",
        CellRef,
        RangeRef,
        TextLiteral,
        NumberLiteral,
        BooleanLiteral
    ]] = Field(default_factory=list)


# 支持递归定义
FunctionCall.model_rebuild()


class FormulaToolCallResponse(BaseModel):
    """模型返回的完整响应"""
    formula: FunctionCall  # 主公式（顶层函数调用）
    explanation: str  # 公式说明（用户可读）
```

---

### 步骤 2：实现编译器（compiler.py）

```python
"""
将 FormulaCall JSON 编译成 Excel 公式文本
"""
from .schema import (
    FunctionCall,
    CellRef,
    RangeRef,
    TextLiteral,
    NumberLiteral,
    BooleanLiteral,
)


def compile_formula(call: FunctionCall) -> str:
    """
    将函数调用 JSON 编译成 Excel 公式
    
    Args:
        call: 函数调用对象
        
    Returns:
        Excel 公式文本（带前导 =）
        
    Example:
        >>> call = FunctionCall(function="SUM", args=[...])
        >>> compile_formula(call)
        '=SUM(A1:A10)'
    """
    formula_text = _compile_node(call)
    return f"={formula_text}"


def _compile_node(node) -> str:
    """递归编译节点"""
    if isinstance(node, FunctionCall):
        return _compile_function(node)
    elif isinstance(node, CellRef):
        return node.address
    elif isinstance(node, RangeRef):
        return _compile_range(node)
    elif isinstance(node, TextLiteral):
        # Excel 文本需要双引号，内部的双引号要转义
        escaped = node.value.replace('"', '""')
        return f'"{escaped}"'
    elif isinstance(node, NumberLiteral):
        return str(node.value)
    elif isinstance(node, BooleanLiteral):
        return "TRUE" if node.value else "FALSE"
    else:
        raise ValueError(f"未知节点类型: {type(node)}")


def _compile_function(call: FunctionCall) -> str:
    """编译函数调用"""
    func_name = call.function.upper()
    
    # 特殊处理：比较运算符
    if func_name == "EQ":
        if len(call.args) != 2:
            raise ValueError("EQ 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}={right})"
    
    if func_name == "NE":
        if len(call.args) != 2:
            raise ValueError("NE 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}<>{right})"
    
    if func_name == "GT":
        if len(call.args) != 2:
            raise ValueError("GT 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}>{right})"
    
    if func_name == "LT":
        if len(call.args) != 2:
            raise ValueError("LT 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}<{right})"
    
    if func_name == "GTE":
        if len(call.args) != 2:
            raise ValueError("GTE 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}>={right})"
    
    if func_name == "LTE":
        if len(call.args) != 2:
            raise ValueError("LTE 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}<={right})"
    
    # 特殊处理：算术运算符
    if func_name == "ADD":
        if len(call.args) != 2:
            raise ValueError("ADD 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}+{right})"
    
    if func_name == "SUBTRACT":
        if len(call.args) != 2:
            raise ValueError("SUBTRACT 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}-{right})"
    
    if func_name == "MULTIPLY":
        if len(call.args) != 2:
            raise ValueError("MULTIPLY 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}*{right})"
    
    if func_name == "DIVIDE":
        if len(call.args) != 2:
            raise ValueError("DIVIDE 需要 2 个参数")
        left = _compile_node(call.args[0])
        right = _compile_node(call.args[1])
        return f"({left}/{right})"
    
    # 常规函数调用
    compiled_args = [_compile_node(arg) for arg in call.args]
    args_str = ",".join(compiled_args)
    return f"{func_name}({args_str})"


def _compile_range(range_ref: RangeRef) -> str:
    """编译区域引用"""
    # 构建工作表引用
    if range_ref.workbook:
        sheet_ref = f"[{range_ref.workbook}]{range_ref.sheet}"
    else:
        sheet_ref = range_ref.sheet
    
    # 构建列引用
    col = range_ref.col.upper()
    row1 = range_ref.row1
    row2 = range_ref.row2
    
    if row2 is None:
        # 整列引用
        cell_range = f"${col}:${col}"
    else:
        # 区域引用
        cell_range = f"${col}${row1}:${col}${row2}"
    
    return f"{sheet_ref}!{cell_range}"
```

---

### 步骤 3：定义函数注册表（function_registry.py）

```python
"""
Excel 函数定义注册表

用于：
1. 模型 Prompt（告诉模型有哪些函数可用）
2. 编译器验证（检查函数调用是否合法）
3. 未来扩展（增加新函数）
"""
from typing import TypedDict


class FunctionDef(TypedDict):
    """函数定义"""
    name: str  # 函数名
    description: str  # 功能说明
    args: list[str]  # 参数说明列表
    example: str  # 使用示例（JSON）


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
                "if_not_found?: 未找到时的默认值"
            ],
            "example": '{"function": "XLOOKUP", "args": [{"type": "cell", "address": "A2"}, {"type": "range", "sheet": "字典", "col": "A", "row1": 2, "row2": 100}, {"type": "range", "sheet": "字典", "col": "B", "row1": 2, "row2": 100}, {"type": "text", "value": ""}]}'
        },
        {
            "name": "VLOOKUP",
            "description": "垂直查找（兼容旧版 Excel）",
            "args": [
                "lookup_value: 要查找的值",
                "table_array: 查找表区域",
                "col_index_num: 返回第几列",
                "range_lookup: FALSE=精确匹配, TRUE=近似匹配"
            ],
            "example": '{"function": "VLOOKUP", "args": [{"type": "cell", "address": "A2"}, {"type": "range", "sheet": "字典", "col": "A", "row1": 2, "row2": 100}, {"type": "number", "value": 2}, {"type": "boolean", "value": false}]}'
        },
        {
            "name": "INDEX",
            "description": "返回指定位置的值",
            "args": [
                "array: 区域",
                "row_num: 行号",
                "col_num?: 列号"
            ],
            "example": '{"function": "INDEX", "args": [{"type": "range", "sheet": "数据", "col": "B", "row1": 2, "row2": 100}, {"type": "number", "value": 5}]}'
        },
        {
            "name": "MATCH",
            "description": "返回值在区域中的位置",
            "args": [
                "lookup_value: 要查找的值",
                "lookup_array: 查找范围",
                "match_type: 0=精确, 1=小于等于, -1=大于等于"
            ],
            "example": '{"function": "MATCH", "args": [{"type": "text", "value": "产品A"}, {"type": "range", "sheet": "列表", "col": "A", "row1": 1, "row2": 100}, {"type": "number", "value": 0}]}'
        }
    ],
    
    "逻辑函数": [
        {
            "name": "IF",
            "description": "条件判断",
            "args": [
                "condition: 条件表达式",
                "value_if_true: 条件为真时的值",
                "value_if_false: 条件为假时的值"
            ],
            "example": '{"function": "IF", "args": [{"function": "GT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 100}]}, {"type": "text", "value": "合格"}, {"type": "text", "value": "不合格"}]}'
        },
        {
            "name": "AND",
            "description": "所有条件都为真",
            "args": [
                "condition1, condition2, ...: 条件表达式"
            ],
            "example": '{"function": "AND", "args": [{"function": "GT", "args": [...]}, {"function": "LT", "args": [...]}]}'
        },
        {
            "name": "OR",
            "description": "任一条件为真",
            "args": [
                "condition1, condition2, ...: 条件表达式"
            ],
            "example": '{"function": "OR", "args": [{"function": "EQ", "args": [...]}, {"function": "EQ", "args": [...]}]}'
        },
        {
            "name": "NOT",
            "description": "条件取反",
            "args": [
                "condition: 条件表达式"
            ],
            "example": '{"function": "NOT", "args": [{"function": "EQ", "args": [...]}]}'
        }
    ],
    
    "文本函数": [
        {
            "name": "SEARCH",
            "description": "在文本中查找子串（不区分大小写）",
            "args": [
                "find_text: 要查找的文本",
                "within_text: 被查找的文本",
                "start_num?: 起始位置"
            ],
            "example": '{"function": "SEARCH", "args": [{"type": "text", "value": "错误"}, {"type": "cell", "address": "D2"}]}'
        },
        {
            "name": "ISNUMBER",
            "description": "判断是否为数字",
            "args": [
                "value: 要判断的值"
            ],
            "example": '{"function": "ISNUMBER", "args": [{"function": "SEARCH", "args": [...]}]}'
        },
        {
            "name": "CONCATENATE",
            "description": "连接文本",
            "args": [
                "text1, text2, ...: 要连接的文本"
            ],
            "example": '{"function": "CONCATENATE", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": "-"}, {"type": "cell", "address": "B2"}]}'
        },
        {
            "name": "LEFT",
            "description": "提取左侧字符",
            "args": [
                "text: 文本",
                "num_chars: 字符数"
            ],
            "example": '{"function": "LEFT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 3}]}'
        },
        {
            "name": "RIGHT",
            "description": "提取右侧字符",
            "args": [
                "text: 文本",
                "num_chars: 字符数"
            ],
            "example": '{"function": "RIGHT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 2}]}'
        },
        {
            "name": "MID",
            "description": "提取中间字符",
            "args": [
                "text: 文本",
                "start_num: 起始位置",
                "num_chars: 字符数"
            ],
            "example": '{"function": "MID", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 3}, {"type": "number", "value": 5}]}'
        }
    ],
    
    "数学与统计": [
        {
            "name": "SUM",
            "description": "求和",
            "args": [
                "number1, number2, ...: 数字或区域"
            ],
            "example": '{"function": "SUM", "args": [{"type": "range", "sheet": "数据", "col": "C", "row1": 2, "row2": 100}]}'
        },
        {
            "name": "SUMIF",
            "description": "条件求和",
            "args": [
                "range: 条件范围",
                "criteria: 条件",
                "sum_range?: 求和范围"
            ],
            "example": '{"function": "SUMIF", "args": [{"type": "range", "sheet": "数据", "col": "A", "row1": 2, "row2": 100}, {"type": "text", "value": "产品A"}, {"type": "range", "sheet": "数据", "col": "B", "row1": 2, "row2": 100}]}'
        },
        {
            "name": "SUMPRODUCT",
            "description": "数组乘积求和",
            "args": [
                "array1, array2, ...: 数组"
            ],
            "example": '{"function": "SUMPRODUCT", "args": [{"function": "MULTIPLY", "args": [{"function": "ISNUMBER", "args": [...]}, {"function": "EQ", "args": [...]}]}]}'
        },
        {
            "name": "AVERAGE",
            "description": "求平均值",
            "args": [
                "number1, number2, ...: 数字或区域"
            ],
            "example": '{"function": "AVERAGE", "args": [{"type": "range", "sheet": "数据", "col": "C", "row1": 2, "row2": 100}]}'
        },
        {
            "name": "COUNT",
            "description": "计数（仅数字）",
            "args": [
                "value1, value2, ...: 值或区域"
            ],
            "example": '{"function": "COUNT", "args": [{"type": "range", "sheet": "数据", "col": "A", "row1": 2, "row2": 100}]}'
        },
        {
            "name": "COUNTA",
            "description": "计数（非空单元格）",
            "args": [
                "value1, value2, ...: 值或区域"
            ],
            "example": '{"function": "COUNTA", "args": [{"type": "range", "sheet": "数据", "col": "A", "row1": 2, "row2": 100}]}'
        }
    ],
    
    "比较运算符": [
        {
            "name": "EQ",
            "description": "等于（=）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "EQ", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": ""}]}'
        },
        {
            "name": "NE",
            "description": "不等于（<>）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "NE", "args": [{"type": "cell", "address": "A2"}, {"type": "text", "value": ""}]}'
        },
        {
            "name": "GT",
            "description": "大于（>）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "GT", "args": [{"type": "number", "value": 5}, {"type": "number", "value": 0}]}'
        },
        {
            "name": "LT",
            "description": "小于（<）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "LT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 100}]}'
        },
        {
            "name": "GTE",
            "description": "大于等于（>=）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "GTE", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 0}]}'
        },
        {
            "name": "LTE",
            "description": "小于等于（<=）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "LTE", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 100}]}'
        }
    ],
    
    "算术运算符": [
        {
            "name": "ADD",
            "description": "加法（+）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "ADD", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 10}]}'
        },
        {
            "name": "SUBTRACT",
            "description": "减法（-）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "SUBTRACT", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 5}]}'
        },
        {
            "name": "MULTIPLY",
            "description": "乘法（*）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "MULTIPLY", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 1.1}]}'
        },
        {
            "name": "DIVIDE",
            "description": "除法（/）",
            "args": [
                "left: 左侧值",
                "right: 右侧值"
            ],
            "example": '{"function": "DIVIDE", "args": [{"type": "cell", "address": "A2"}, {"type": "number", "value": 2}]}'
        }
    ]
}


def get_functions_prompt() -> str:
    """生成函数列表的 Prompt 文本"""
    lines = ["# 可用的 Excel 函数\n"]
    
    for category, functions in EXCEL_FUNCTIONS.items():
        lines.append(f"\n## {category}\n")
        for func in functions:
            lines.append(f"### {func['name']}")
            lines.append(f"{func['description']}\n")
            lines.append("参数：")
            for arg in func['args']:
                lines.append(f"- {arg}")
            lines.append(f"\n示例：\n```json\n{func['example']}\n```\n")
    
    return "\n".join(lines)
```

---

### 步骤 4：编写模型 Prompt（model_prompt.py）

```python
"""
模型 Prompt：指导模型生成结构化的公式调用
"""
from .function_registry import get_functions_prompt


SYSTEM_PROMPT = f"""你是 Excel 公式工具调用生成器。你的任务是将用户的自然语言需求转换成结构化的 Excel 函数调用（JSON 格式）。

# 核心原则

1. **你不直接生成公式文本**，而是返回结构化的函数调用 JSON
2. **代码会将你的 JSON 编译成 Excel 公式**，保证语法 100% 正确
3. **你只需专注于理解需求，选择合适的函数和参数**

# 数据类型

## 值类型

### 单元格引用
```json
{{"type": "cell", "address": "D2"}}
```

### 区域引用
```json
{{
  "type": "range",
  "workbook": "tt.xlsx",  // 外部工作簿文件名，当前工作簿则为 null
  "sheet": "字典",         // 工作表名
  "col": "A",             // 列字母
  "row1": 2,              // 起始行
  "row2": 999             // 结束行，null 表示整列
}}
```

### 文本字面量
```json
{{"type": "text", "value": "p0"}}
{{"type": "text", "value": "/"}}  // 特殊字符原样保留
```

### 数字字面量
```json
{{"type": "number", "value": 0}}
```

### 布尔字面量
```json
{{"type": "boolean", "value": false}}
```

## 函数调用

```json
{{
  "function": "函数名",
  "args": [参数列表]
}}
```

函数调用可以嵌套：
```json
{{
  "function": "IF",
  "args": [
    {{"function": "GT", "args": [{{"type": "cell", "address": "A2"}}, {{"type": "number", "value": 0}}]}},
    {{"type": "text", "value": "正数"}},
    {{"type": "text", "value": "非正数"}}
  ]
}}
```

{get_functions_prompt()}

# 输出格式

你必须返回以下 JSON 格式：

```json
{{
  "formula": {{
    "function": "...",
    "args": [...]
  }},
  "explanation": "公式说明（中文，用户可读）"
}}
```

# 重要提示

1. **引用完整区域，不要硬编码值**
   - ❌ 错误：枚举所有关键词 `{{"语法错误", "类型错误"}}`
   - ✅ 正确：引用字典表区域 `{{"type": "range", "sheet": "tt", "col": "A", "row1": 2, "row2": 999}}`

2. **字面量原样保留**
   - "/" 这种字符要写成 `{{"type": "text", "value": "/"}}`，不要遗漏

3. **优先使用现代函数**
   - 优先 XLOOKUP（Excel 365+）
   - 兼容性需求才用 VLOOKUP

4. **比较运算用函数表示**
   - `A2=""` 写成 `{{"function": "EQ", "args": [{{"type": "cell", "address": "A2"}}, {{"type": "text", "value": ""}}]}}`
   - `A2>0` 写成 `{{"function": "GT", "args": [{{"type": "cell", "address": "A2"}}, {{"type": "number", "value": 0}}]}}`

5. **算术运算用函数表示**
   - `A2+10` 写成 `{{"function": "ADD", "args": [{{"type": "cell", "address": "A2"}}, {{"type": "number", "value": 10}}]}}`
   - `A2*1.1` 写成 `{{"function": "MULTIPLY", "args": [{{"type": "cell", "address": "A2"}}, {{"type": "number", "value": 1.1}}]}}`

# Few-shot 示例

## 示例 1：精确查表

**用户需求**：
```
目标单元格：E2
需求：根据 D 列的员工 ID 查找对应的部门
列头映射：D列=员工ID、E列=部门
当前工作簿内有字典表「员工信息」（A列=ID、B列=部门，共 50 行数据）
```

**你的输出**：
```json
{{
  "formula": {{
    "function": "IF",
    "args": [
      {{"function": "EQ", "args": [{{"type": "cell", "address": "D2"}}, {{"type": "text", "value": ""}}]}},
      {{"type": "text", "value": ""}},
      {{"function": "XLOOKUP", "args": [
        {{"type": "cell", "address": "D2"}},
        {{"type": "range", "workbook": null, "sheet": "员工信息", "col": "A", "row1": 2, "row2": 51}},
        {{"type": "range", "workbook": null, "sheet": "员工信息", "col": "B", "row1": 2, "row2": 51}},
        {{"type": "text", "value": ""}}
      ]}}
    ]
  }},
  "explanation": "如果 D 列为空则返回空，否则使用 XLOOKUP 在员工信息表中查找 D 列的 ID，返回对应的部门。未找到时返回空。"
}}
```

## 示例 2：优先级关键词匹配

**用户需求**：
```
目标单元格：E2
需求：D 列包含问题类型（可能有多个，逗号分隔），根据问题类型返回最高错误级别
列头映射：D列=问题类型、E列=错误级别
外部工作簿「tt.xlsx」工作表「tt」（A列=关键词、B列=级别，包含 p0/p1/等级别，共 20 行）
优先级：p0 > p1 > /
```

**你的输出**：
```json
{{
  "formula": {{
    "function": "IF",
    "args": [
      {{"function": "EQ", "args": [{{"type": "cell", "address": "D2"}}, {{"type": "text", "value": ""}}]}},
      {{"type": "text", "value": ""}},
      {{"function": "IF", "args": [
        {{"function": "GT", "args": [
          {{"function": "SUMPRODUCT", "args": [
            {{"function": "MULTIPLY", "args": [
              {{"function": "ISNUMBER", "args": [
                {{"function": "SEARCH", "args": [
                  {{"type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "A", "row1": 2, "row2": 21}},
                  {{"type": "cell", "address": "D2"}}
                ]}}
              ]}},
              {{"function": "EQ", "args": [
                {{"type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "B", "row1": 2, "row2": 21}},
                {{"type": "text", "value": "p0"}}
              ]}}
            ]}}
          ]}},
          {{"type": "number", "value": 0}}
        ]}},
        {{"type": "text", "value": "p0"}},
        {{"function": "IF", "args": [
          {{"function": "GT", "args": [
            {{"function": "SUMPRODUCT", "args": [
              {{"function": "MULTIPLY", "args": [
                {{"function": "ISNUMBER", "args": [
                  {{"function": "SEARCH", "args": [
                    {{"type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "A", "row1": 2, "row2": 21}},
                    {{"type": "cell", "address": "D2"}}
                  ]}}
                ]}},
                {{"function": "EQ", "args": [
                  {{"type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "B", "row1": 2, "row2": 21}},
                  {{"type": "text", "value": "p1"}}
                ]}}
              ]}}
            ]}},
            {{"type": "number", "value": 0}}
          ]}},
          {{"type": "text", "value": "p1"}},
          {{"function": "IF", "args": [
            {{"function": "GT", "args": [
              {{"function": "SUMPRODUCT", "args": [
                {{"function": "MULTIPLY", "args": [
                  {{"function": "ISNUMBER", "args": [
                    {{"function": "SEARCH", "args": [
                      {{"type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "A", "row1": 2, "row2": 21}},
                      {{"type": "cell", "address": "D2"}}
                    ]}}
                  ]}},
                  {{"function": "EQ", "args": [
                    {{"type": "range", "workbook": "tt.xlsx", "sheet": "tt", "col": "B", "row1": 2, "row2": 21}},
                    {{"type": "text", "value": "/"}}
                  ]}}
                ]}}
              ]}},
              {{"type": "number", "value": 0}}
            ]}},
            {{"type": "text", "value": "/"}},
            {{"type": "text", "value": ""}}
          ]}}
        ]}}
      ]}}
    ]
  }},
  "explanation": "如果 D 列为空则返回空，否则检测 D 列包含哪些关键词（支持逗号分隔的多个问题）。使用 SUMPRODUCT+ISNUMBER+SEARCH 在 tt.xlsx 的字典表中匹配关键词，按优先级 p0 > p1 > / 返回最高级别。未匹配到任何关键词时返回空。"
}}
```

## 示例 3：条件统计

**用户需求**：
```
目标单元格：B10
需求：统计 B 列中大于 100 的数字个数
列头映射：B列=销售额
数据范围：B2:B50
```

**你的输出**：
```json
{{
  "formula": {{
    "function": "SUMPRODUCT",
    "args": [
      {{"function": "GT", "args": [
        {{"type": "range", "workbook": null, "sheet": "Sheet1", "col": "B", "row1": 2, "row2": 50}},
        {{"type": "number", "value": 100}}
      ]}}
    ]
  }},
  "explanation": "使用 SUMPRODUCT 统计 B2:B50 中大于 100 的数字个数。比较运算返回 TRUE/FALSE 数组，SUMPRODUCT 自动将 TRUE 视为 1 进行求和。"
}}
```

# 开始工作

现在，根据用户提供的需求，生成结构化的公式调用 JSON。
"""


def build_user_message(
    active_cell: str,
    description: str,
    headers: list[str] | None,
    columns: list[str] | None,
    dictionary: dict | None,
    extra_sheets: list[dict] | None,
) -> str:
    """构建用户消息"""
    lines = [
        f"目标单元格：{active_cell}",
        f"需求：{description}",
    ]
    
    if headers and columns:
        mapping = "、".join(f"{col}列={h}" for col, h in zip(columns, headers))
        lines.append(f"列头映射：{mapping}")
    
    if dictionary:
        lines.append(
            f"\n当前工作簿内有字典表「{dictionary['name']}」"
            f"（A列={dictionary.get('keyHeader', '键')}、"
            f"B列={dictionary.get('valueHeader', '值')}，"
            f"共 {dictionary.get('rowCount', '?')} 行数据）"
        )
    
    if extra_sheets:
        for extra in extra_sheets:
            mapping = "、".join(
                f"{col}列={h}" for col, h in zip(extra['columns'], extra['headers'])
            )
            lines.append(
                f"\n外部工作簿「{extra['sourceFile']}」工作表「{extra['sheetName']}」"
                f"（{mapping}，共 {extra.get('rowCount', '?')} 行数据）"
            )
    
    lines.append("\n请生成结构化的公式调用 JSON。")
    
    return "\n".join(lines)
```

---

### 步骤 5：集成到现有流程（formula_tools_integration.py）

```python
"""
将工具调用架构集成到现有的公式生成流程
"""
import json
import logging
from typing import Any

from .formula_tools.schema import FormulaToolCallResponse, FunctionCall
from .formula_tools.compiler import compile_formula
from .formula_tools.model_prompt import SYSTEM_PROMPT, build_user_message
from .llm import formula_model_config, selected_model_config
from .llm.client import OpenAICompatibleClient
from .capabilities import capability_float


logger = logging.getLogger(__name__)


async def generate_formula_with_tool_call(
    active_cell: str,
    description: str,
    headers: list[str] | None,
    columns: list[str] | None,
    dictionary: dict | None,
    extra_sheets: list[dict] | None,
    model_id: str | None,
) -> dict[str, str]:
    """
    使用工具调用架构生成公式
    
    Returns:
        {
            "modernFormula": "=...",
            "modernExplanation": "...",
            "compatFormula": "=...",  # 暂时与 modern 相同
            "compatExplanation": "..."
        }
    """
    # 构建消息
    user_message = build_user_message(
        active_cell=active_cell,
        description=description,
        headers=headers,
        columns=columns,
        dictionary=dictionary,
        extra_sheets=extra_sheets,
    )
    
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_message},
    ]
    
    # 调用模型
    config = formula_model_config() or selected_model_config(model_id)
    if not config:
        raise ValueError("未配置模型")
    
    timeout = capability_float("llm", "timeoutSeconds")
    
    async with OpenAICompatibleClient(config, timeout=timeout) as client:
        response = await client.chat_completions(
            messages=messages,
            max_tokens=4000,
            temperature=0.1,  # 低温度保证稳定性
        )
    
    content = response["choices"][0]["message"].get("content", "")
    
    # 解析 JSON
    try:
        json_match = _extract_json(content)
        data = json.loads(json_match)
        parsed = FormulaToolCallResponse(**data)
    except Exception as e:
        logger.error(f"解析模型返回失败: {e}\n内容: {content}")
        raise ValueError(f"模型返回格式错误: {e}")
    
    # 编译公式
    try:
        formula = compile_formula(parsed.formula)
    except Exception as e:
        logger.error(f"编译公式失败: {e}\nJSON: {parsed.formula}")
        raise ValueError(f"公式编译失败: {e}")
    
    return {
        "modernFormula": formula,
        "modernExplanation": parsed.explanation,
        "compatFormula": formula,  # TODO: 未来可生成兼容版本（VLOOKUP 等）
        "compatExplanation": parsed.explanation,
    }


def _extract_json(text: str) -> str:
    """从文本中提取 JSON（支持 ```json 代码块）"""
    import re
    
    # 尝试提取 ```json 代码块
    code_block_match = re.search(r"```(?:json)?\s*\n(.*?)\n```", text, re.DOTALL)
    if code_block_match:
        return code_block_match.group(1).strip()
    
    # 尝试提取裸 JSON
    json_match = re.search(r"\{.*\}", text, re.DOTALL)
    if json_match:
        return json_match.group(0)
    
    raise ValueError("未找到 JSON")
```

---

### 步骤 6：修改主流程（rule_generator.py）

```python
# 在 server/app/rule_generator.py 的 generate_formula 函数中增加工具调用入口

async def generate_formula(request: GenerateFormulaRequest) -> GenerateFormulaResponse:
    """
    生成公式：优先使用工具调用架构，失败则回退到原有方式
    """
    # 第一步：尝试工具调用架构
    try:
        from .formula_tools_integration import generate_formula_with_tool_call
        
        result = await generate_formula_with_tool_call(
            active_cell=request.activeCell,
            description=request.description,
            headers=request.headers,
            columns=request.columns,
            dictionary=request.dictionary.model_dump() if request.dictionary else None,
            extra_sheets=[s.model_dump() for s in request.extraSheets] if request.extraSheets else None,
            model_id=request.modelId,
        )
        
        return GenerateFormulaResponse(**result)
        
    except Exception as e:
        import logging
        logging.warning(f"工具调用架构失败，回退到原有方式: {e}")
    
    # 第二步：回退到现有的完整模型生成（保留现有逻辑）
    config = formula_model_config() or selected_model_config(request.modelId)
    if not config:
        raise ValueError("未配置模型")
    
    messages = _build_formula_generation_prompt(request)
    # ... 现有逻辑 ...
```

---

### 步骤 7：编写单元测试

```python
# server/tests/test_formula_tools.py

import pytest
from server.app.formula_tools.schema import (
    FunctionCall,
    CellRef,
    RangeRef,
    TextLiteral,
    NumberLiteral,
)
from server.app.formula_tools.compiler import compile_formula


def test_compile_simple_sum():
    """测试简单 SUM 函数"""
    call = FunctionCall(
        function="SUM",
        args=[
            RangeRef(type="range", workbook=None, sheet="数据", col="A", row1=1, row2=10)
        ]
    )
    result = compile_formula(call)
    assert result == "=SUM(数据!$A$1:$A$10)"


def test_compile_xlookup():
    """测试 XLOOKUP 函数"""
    call = FunctionCall(
        function="XLOOKUP",
        args=[
            CellRef(type="cell", address="D2"),
            RangeRef(type="range", workbook=None, sheet="字典", col="A", row1=2, row2=100),
            RangeRef(type="range", workbook=None, sheet="字典", col="B", row1=2, row2=100),
            TextLiteral(type="text", value=""),
        ]
    )
    result = compile_formula(call)
    assert result == '=XLOOKUP(D2,字典!$A$2:$A$100,字典!$B$2:$B$100,"")'


def test_compile_external_workbook():
    """测试外部工作簿引用"""
    call = FunctionCall(
        function="XLOOKUP",
        args=[
            CellRef(type="cell", address="D2"),
            RangeRef(type="range", workbook="tt.xlsx", sheet="tt", col="A", row1=2, row2=999),
            RangeRef(type="range", workbook="tt.xlsx", sheet="tt", col="B", row1=2, row2=999),
            TextLiteral(type="text", value=""),
        ]
    )
    result = compile_formula(call)
    assert result == '=XLOOKUP(D2,[tt.xlsx]tt!$A$2:$A$999,[tt.xlsx]tt!$B$2:$B$999,"")'


def test_compile_comparison():
    """测试比较运算符"""
    call = FunctionCall(
        function="EQ",
        args=[
            CellRef(type="cell", address="A2"),
            TextLiteral(type="text", value=""),
        ]
    )
    result = compile_formula(call)
    assert result == '=(A2="")'


def test_compile_nested_if():
    """测试嵌套 IF"""
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="GT",
                args=[
                    CellRef(type="cell", address="A2"),
                    NumberLiteral(type="number", value=0),
                ]
            ),
            TextLiteral(type="text", value="正数"),
            TextLiteral(type="text", value="非正数"),
        ]
    )
    result = compile_formula(call)
    assert result == '=IF((A2>0),"正数","非正数")'


def test_compile_special_characters():
    """测试特殊字符（/ 等）原样保留"""
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="EQ",
                args=[
                    CellRef(type="cell", address="A2"),
                    TextLiteral(type="text", value="/"),
                ]
            ),
            TextLiteral(type="text", value="斜杠"),
            TextLiteral(type="text", value="其他"),
        ]
    )
    result = compile_formula(call)
    assert result == '=IF((A2="/"),"斜杠","其他")'


def test_compile_text_with_quotes():
    """测试文本中包含双引号"""
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="EQ",
                args=[
                    CellRef(type="cell", address="A2"),
                    TextLiteral(type="text", value='他说"你好"'),
                ]
            ),
            TextLiteral(type="text", value="匹配"),
            TextLiteral(type="text", value="不匹配"),
        ]
    )
    result = compile_formula(call)
    # Excel 中双引号要转义成两个双引号
    assert result == '=IF((A2="他说""你好"""),"匹配","不匹配")'


def test_compile_sumproduct_priority_match():
    """测试优先级关键词匹配公式（完整示例）"""
    call = FunctionCall(
        function="IF",
        args=[
            FunctionCall(
                function="EQ",
                args=[
                    CellRef(type="cell", address="D2"),
                    TextLiteral(type="text", value=""),
                ]
            ),
            TextLiteral(type="text", value=""),
            FunctionCall(
                function="IF",
                args=[
                    FunctionCall(
                        function="GT",
                        args=[
                            FunctionCall(
                                function="SUMPRODUCT",
                                args=[
                                    FunctionCall(
                                        function="MULTIPLY",
                                        args=[
                                            FunctionCall(
                                                function="ISNUMBER",
                                                args=[
                                                    FunctionCall(
                                                        function="SEARCH",
                                                        args=[
                                                            RangeRef(type="range", workbook="tt.xlsx", sheet="tt", col="A", row1=2, row2=999),
                                                            CellRef(type="cell", address="D2"),
                                                        ]
                                                    )
                                                ]
                                            ),
                                            FunctionCall(
                                                function="EQ",
                                                args=[
                                                    RangeRef(type="range", workbook="tt.xlsx", sheet="tt", col="B", row1=2, row2=999),
                                                    TextLiteral(type="text", value="p0"),
                                                ]
                                            )
                                        ]
                                    )
                                ]
                            ),
                            NumberLiteral(type="number", value=0),
                        ]
                    ),
                    TextLiteral(type="text", value="p0"),
                    TextLiteral(type="text", value=""),
                ]
            )
        ]
    )
    
    result = compile_formula(call)
    expected = '=IF((D2=""),"",IF((SUMPRODUCT((ISNUMBER(SEARCH([tt.xlsx]tt!$A$2:$A$999,D2))*([tt.xlsx]tt!$B$2:$B$999="p0")))>0),"p0",""))'
    assert result == expected
```

---

## 实施时间估算

| 步骤 | 预计时间 | 说明 |
|------|----------|------|
| 1. 定义数据模型 | 30 分钟 | schema.py，Pydantic 模型定义 |
| 2. 实现编译器 | 1 小时 | compiler.py，递归编译逻辑 |
| 3. 定义函数注册表 | 45 分钟 | function_registry.py，初期支持 20+ 函数 |
| 4. 编写模型 Prompt | 1 小时 | model_prompt.py，System Prompt + Few-shot |
| 5. 集成到现有流程 | 30 分钟 | formula_tools_integration.py |
| 6. 修改主流程 | 15 分钟 | rule_generator.py 增加入口 |
| 7. 编写单元测试 | 1 小时 | test_formula_tools.py |
| 8. 集成测试与调试 | 1 小时 | 真实场景测试、修复问题 |
| **总计** | **6 小时** | 包含测试和调试 |

---

## 验收标准

### 功能验收

1. **精确查表场景**（如员工 ID → 部门）
   - 生成的公式语法 100% 正确
   - 引用完整区域（不硬编码值）
   - 支持外部工作簿引用

2. **优先级关键词匹配场景**（如问题类型 → 错误级别）
   - 特殊字符（"/"）不丢失
   - SUMPRODUCT 公式结构正确
   - 优先级逻辑准确

3. **回退机制**
   - 工具调用失败时自动回退到原有方式
   - 不破坏现有功能

### 性能验收

- 单次调用延迟 < 5 秒（模型推理时间）
- 编译器性能：< 10ms（纯代码，不依赖模型）

### 测试覆盖

- 单元测试：≥ 80% 覆盖率
- 集成测试：覆盖 3 种以上典型场景

---

## 扩展计划（可选）

### 短期扩展（1-2 周）

1. **增加函数支持**
   - 日期函数（TODAY, DATE, YEAR, MONTH, DAY, DATEDIF）
   - 条件聚合（SUMIFS, COUNTIFS, AVERAGEIF）
   - 数组函数（FILTER, SORT, UNIQUE - Excel 365）

2. **生成兼容版本**
   - 自动将 XLOOKUP 转换成 INDEX+MATCH 或 VLOOKUP
   - 检测用户 Excel 版本，返回对应公式

3. **错误处理增强**
   - 捕获编译错误，提示用户具体问题
   - 提供修复建议

### 长期扩展（1-2 月）

1. **公式验算集成**
   - 编译后自动调用前端验算
   - 错误时反馈给模型重新生成

2. **公式优化**
   - 检测冗余的嵌套 IF
   - 建议更简洁的实现

3. **自定义函数支持**
   - 用户可定义自己的函数 schema
   - 代码生成自定义函数调用

---

## 给 Codex 的实施指令

### 前置条件

- Python 3.10+
- 依赖：pydantic（已有）

### 实施步骤

1. **创建文件夹结构**
   ```bash
   mkdir -p server/app/formula_tools
   touch server/app/formula_tools/__init__.py
   ```

2. **按顺序创建文件**（确保依赖顺序）：
   - `server/app/formula_tools/schema.py`
   - `server/app/formula_tools/compiler.py`
   - `server/app/formula_tools/function_registry.py`
   - `server/app/formula_tools/model_prompt.py`
   - `server/app/formula_tools_integration.py`
   
3. **修改现有文件**：
   - `server/app/rule_generator.py`：在 `generate_formula` 函数开头增加工具调用入口

4. **创建测试文件**：
   - `server/tests/test_formula_tools.py`

5. **运行测试**：
   ```bash
   python -m pytest server/tests/test_formula_tools.py -v --basetemp=.pytest-tmp
   ```

6. **集成测试**：
   用真实场景测试（评分表 + tt.xlsx 字典表）

### 关键注意事项

1. **字面量转义**：
   - 文本中的双引号要转义成 `""`
   - 特殊字符（"/"）原样保留

2. **区域引用格式**：
   - 当前工作簿：`Sheet1!$A$2:$A$100`
   - 外部工作簿：`[文件名.xlsx]Sheet1!$A$2:$A$100`
   - 列字母大写，行号加 `$`

3. **模型温度**：
   - 设置 `temperature=0.1`（低温度保证稳定性）

4. **错误处理**：
   - 工具调用失败时记录日志，回退到原有方式
   - 不要让工具调用失败导致整个接口报错

5. **向后兼容**：
   - 保留原有的公式生成逻辑作为兜底
   - 用户无感知切换

---

## 模型 Prompt 设计要点（重要）

### Prompt 核心原则

1. **明确角色**："你是工具调用生成器，不是公式生成器"
2. **强调约束**："不要生成公式文本，只返回 JSON"
3. **提供完整 Schema**：所有数据类型的定义和示例
4. **Few-shot 示例**：至少 3 个典型场景（简单 → 复杂）
5. **特别提示**：常见错误（硬编码、字符丢失）

### Prompt 优化技巧

1. **用负面案例教学**：
   ```
   ❌ 错误：{"args": [{"type": "text", "value": "语法错误"}]}
   ✅ 正确：{"args": [{"type": "range", "sheet": "tt", "col": "A", ...}]}
   ```

2. **强调"完整区域引用"**：
   ```
   重要提示：永远引用字典表的完整区域，不要枚举具体值。
   ```

3. **低温度**：
   ```python
   temperature=0.1  # 保证稳定性，减少随机性
   ```

4. **结构化输出**：
   ```python
   # 如果模型支持 JSON Schema 约束（如 OpenAI），可传入 schema
   response_format={"type": "json_object"}
   ```

---

## 风险与应对

### 风险 1：模型无法稳定输出合法 JSON

**应对**：
- 增加 Few-shot 示例（从 2 个增加到 5 个）
- 使用 `response_format={"type": "json_object"}`（如果模型支持）
- 重试机制：失败 3 次后回退到原有方式

### 风险 2：复杂嵌套公式 JSON 过长

**应对**：
- 将 `max_tokens` 提升到 4000（现在是 2000）
- 如果仍不够，提示用户"公式过于复杂，建议简化需求"

### 风险 3：编译器 Bug 导致公式错误

**应对**：
- 完善单元测试（覆盖所有函数和边界情况）
- 增加编译后验算（调用前端的 `formulaSafety.ts`）
- 记录编译失败的 JSON，持续修复 Bug

### 风险 4：用户反馈"不如之前"

**应对**：
- 初期保持两种方式并存（A/B 测试）
- 收集失败案例，持续优化 Prompt
- 记录成功率对比数据

---

## 成功标准

### 质量标准

- ✅ 语法正确率：100%（代码生成，不依赖模型）
- ✅ 字面量保留率：100%（"/", "p0" 等不丢失）
- ✅ 场景覆盖率：≥ 80%（复杂场景可回退）

### 对比原方案

| 指标 | 原方案（端到端） | 新方案（工具调用） |
|------|-----------------|-------------------|
| 语法正确率 | ~70% | 100% |
| 字面量丢失 | 常见（"/"） | 不会 |
| 引用错误 | 常见（硬编码） | 不会 |
| 扩展性 | 差（依赖 Prompt） | 好（增加函数定义） |
| 调试难度 | 高（黑盒） | 低（JSON 可检查） |

---

## 后续优化方向

1. **公式可视化**：将 JSON 渲染成树状图，帮助用户理解
2. **交互式构建**：前端提供拖拽式公式构建器
3. **公式库**：保存常用公式模板，一键复用
4. **多语言支持**：英文函数名 → 中文函数名（Excel 国际化）

---

## 参考资料

- [OpenAI Function Calling 文档](https://platform.openai.com/docs/guides/function-calling)
- [Anthropic Tool Use 文档](https://docs.anthropic.com/claude/docs/tool-use)
- [Excel 函数参考](https://support.microsoft.com/zh-cn/office/excel-函数-按类别列出-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)
- [Pydantic 文档](https://docs.pydantic.dev/)

---

**方案编写完成时间**：2026-08-20  
**预计实施完成时间**：2026-08-21（6 小时开发 + 测试）  
**方案版本**：v1.0


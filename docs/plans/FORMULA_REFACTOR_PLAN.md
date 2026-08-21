# /function 两级架构改造方案

## 目标

将公式生成从"模型端到端生成完整公式文本"改为"模型分析场景 → 本地代码生成公式"，实现零失误的结构化场景覆盖。

## 核心原则

- **模型负责**：语义理解、场景识别、字段映射
- **代码负责**：公式文本生成、语法正确性、完整性保证
- **渐进式**：先支持常见场景，复杂场景回退到现有模式

---

## 改造文件清单

### 1. 后端：场景分析器（新增）
**文件**：`server/app/formula_analyzer.py`

定义场景枚举和分析逻辑：
```python
from enum import Enum
from pydantic import BaseModel

class FormulaScenario(str, Enum):
    """公式场景类型"""
    LOOKUP = "lookup"  # 精确查表
    PRIORITY_KEYWORD_MATCH = "priority_keyword_match"  # 优先级+关键词匹配
    KEYWORD_MATCH = "keyword_match"  # 简单关键词匹配
    CALCULATION = "calculation"  # 数学计算
    COMPLEX = "complex"  # 复杂逻辑（回退到模型生成）

class DictionaryReference(BaseModel):
    """字典表引用信息"""
    sourceFile: str | None = None  # 外部文件名（如 tt.xlsx），None 表示当前工作簿
    sheetName: str  # 工作表名
    keyColumn: str  # 键列字母（如 "A"）
    valueColumn: str  # 值列字母（如 "B"）
    startRow: int = 2  # 数据起始行（默认第2行，跳过表头）
    endRow: int | None = None  # 数据结束行（None 表示自动检测）

class FormulaAnalysis(BaseModel):
    """场景分析结果"""
    scenario: FormulaScenario
    sourceColumn: str  # 源列字母（如 "D"）
    targetColumn: str  # 目标列字母（从 activeCell 解析，如 "E"）
    dictionary: DictionaryReference | None = None
    priorities: list[str] | None = None  # 优先级顺序（如 ["p0", "p1", "/"]）
    defaultValue: str = ""  # 默认值
    explanation: str = ""  # 分析说明

async def analyze_formula_scenario(request: GenerateFormulaRequest) -> FormulaAnalysis:
    """
    调用模型分析场景，返回结构化的场景描述。
    
    Prompt 设计：
    - 只要求模型返回场景类型和字段映射
    - 不要求模型生成公式文本
    - 输出格式：严格的 JSON schema
    """
    pass  # TODO: 实现
```

**关键点**：
- 模型只返回结构化 JSON，不生成公式文本
- 场景枚举明确，避免歧义
- 字典表引用信息完整（支持跨文件）

---

### 2. 后端：公式生成器（新增）
**文件**：`server/app/formula_generator.py`

根据场景分析结果生成公式：
```python
from .formula_analyzer import FormulaAnalysis, FormulaScenario, DictionaryReference

def _build_dict_range(dictionary: DictionaryReference, column: str) -> str:
    """构建字典表的列引用字符串"""
    sheet_ref = (
        f"[{dictionary.sourceFile}]{dictionary.sheetName}"
        if dictionary.sourceFile
        else dictionary.sheetName
    )
    end = dictionary.endRow or 999  # 如果没指定结束行，用 999 作为安全边界
    return f"{sheet_ref}!${column}${dictionary.startRow}:${column}${end}"

def generate_lookup_formula(analysis: FormulaAnalysis, modern: bool = True) -> tuple[str, str]:
    """生成精确查表公式
    
    返回：(formula, explanation)
    """
    source = f"{analysis.sourceColumn}2"
    dict_key = _build_dict_range(analysis.dictionary, analysis.dictionary.keyColumn)
    dict_value = _build_dict_range(analysis.dictionary, analysis.dictionary.valueColumn)
    
    if modern:
        # 现代版：XLOOKUP
        formula = f'=IF({source}="","",XLOOKUP({source},{dict_key},{dict_value},""))'
        explanation = f"使用 XLOOKUP 在字典表中精确查找 {analysis.sourceColumn} 列的值，返回对应的结果"
    else:
        # 兼容版：VLOOKUP
        dict_range = _build_dict_range(analysis.dictionary, analysis.dictionary.keyColumn).replace(
            f"${analysis.dictionary.keyColumn}$",
            f"${analysis.dictionary.keyColumn}$"
        )
        # 需要构建完整的两列区域
        start_col = analysis.dictionary.keyColumn
        end_col = analysis.dictionary.valueColumn
        sheet_ref = (
            f"[{analysis.dictionary.sourceFile}]{analysis.dictionary.sheetName}"
            if analysis.dictionary.sourceFile
            else analysis.dictionary.sheetName
        )
        end = analysis.dictionary.endRow or 999
        vlookup_range = f"{sheet_ref}!${start_col}${analysis.dictionary.startRow}:${end_col}${end}"
        col_index = ord(end_col) - ord(start_col) + 1
        
        formula = f'=IF({source}="","",IFERROR(VLOOKUP({source},{vlookup_range},{col_index},FALSE),""))'
        explanation = f"使用 VLOOKUP 在字典表中精确查找（第4参数=FALSE），返回对应值"
    
    return (formula, explanation)

def generate_priority_keyword_formula(analysis: FormulaAnalysis, modern: bool = True) -> tuple[str, str]:
    """生成优先级+关键词匹配公式
    
    场景：D列可能包含多个关键词（逗号分隔），需要检测所有关键词并按优先级返回最高级别
    
    返回：(formula, explanation)
    """
    source = f"{analysis.sourceColumn}2"
    dict_key = _build_dict_range(analysis.dictionary, analysis.dictionary.keyColumn)
    dict_value = _build_dict_range(analysis.dictionary, analysis.dictionary.valueColumn)
    
    # 构建按优先级嵌套的 IF 判断
    conditions = []
    for priority in analysis.priorities:
        # SUMPRODUCT((ISNUMBER(SEARCH(关键词区域, 源格)))*(值区域=优先级))>0
        condition = f'SUMPRODUCT((ISNUMBER(SEARCH({dict_key},{source})))*({dict_value}="{priority}"))>0'
        conditions.append((condition, priority))
    
    # 构建嵌套 IF
    formula_parts = [f'IF({source}="",""']
    for condition, value in conditions:
        formula_parts.append(f',IF({condition},"{value}"')
    formula_parts.append(f',"{analysis.defaultValue}"')
    formula_parts.append(')' * len(conditions))
    formula_parts.append(')')
    
    formula = '=' + ''.join(formula_parts)
    
    explanation = (
        f"检测 {analysis.sourceColumn} 列包含哪些关键词（支持逗号分隔的多个问题），"
        f"按优先级 {' > '.join(analysis.priorities)} 返回最高级别。"
        f"使用 SUMPRODUCT+ISNUMBER+SEARCH 进行关键词匹配，引用字典表的完整区域。"
    )
    
    return (formula, explanation)

def generate_keyword_match_formula(analysis: FormulaAnalysis, modern: bool = True) -> tuple[str, str]:
    """生成简单关键词匹配公式（单一匹配，不涉及优先级）"""
    # TODO: 实现
    pass

def generate_formula_from_analysis(analysis: FormulaAnalysis) -> dict:
    """根据场景分析生成公式
    
    返回：
    {
        "modernFormula": "=...",
        "modernExplanation": "...",
        "compatFormula": "=...",
        "compatExplanation": "..."
    }
    """
    if analysis.scenario == FormulaScenario.LOOKUP:
        modern_formula, modern_exp = generate_lookup_formula(analysis, modern=True)
        compat_formula, compat_exp = generate_lookup_formula(analysis, modern=False)
    elif analysis.scenario == FormulaScenario.PRIORITY_KEYWORD_MATCH:
        modern_formula, modern_exp = generate_priority_keyword_formula(analysis, modern=True)
        compat_formula, compat_exp = generate_priority_keyword_formula(analysis, modern=False)
    elif analysis.scenario == FormulaScenario.KEYWORD_MATCH:
        modern_formula, modern_exp = generate_keyword_match_formula(analysis, modern=True)
        compat_formula, compat_exp = generate_keyword_match_formula(analysis, modern=False)
    else:
        # 复杂场景：回退到现有的模型生成逻辑
        raise ValueError(f"场景 {analysis.scenario} 需要回退到完整模型生成")
    
    return {
        "modernFormula": modern_formula,
        "modernExplanation": modern_exp,
        "compatFormula": compat_formula,
        "compatExplanation": compat_exp,
    }
```

**关键点**：
- 纯代码生成，无 LLM 调用
- 公式语法 100% 正确
- 引用完整区域，不硬编码值

---

### 3. 后端：主流程改造
**文件**：`server/app/rule_generator.py`

在 `generate_formula` 函数中增加两级架构的入口：
```python
async def generate_formula(request: GenerateFormulaRequest) -> GenerateFormulaResponse:
    """短链生成原生 Excel 公式：优先用两级架构，复杂场景回退到模型生成"""
    
    # 第一步：尝试场景分析
    try:
        from .formula_analyzer import analyze_formula_scenario
        from .formula_generator import generate_formula_from_analysis
        
        analysis = await analyze_formula_scenario(request)
        
        # 如果是结构化场景，用代码生成
        if analysis.scenario != FormulaScenario.COMPLEX:
            result = generate_formula_from_analysis(analysis)
            return GenerateFormulaResponse(**result)
    except Exception as e:
        # 场景分析失败，回退到原有逻辑
        import logging
        logging.warning(f"场景分析失败，回退到模型生成: {e}")
    
    # 第二步：回退到现有的完整模型生成
    # [保留现有的 generate_formula 逻辑]
    config = formula_model_config() or selected_model_config(request.modelId)
    if not config:
        raise ValueError("未配置模型")
    
    messages = _build_formula_generation_prompt(request)
    # ... 现有逻辑 ...
```

**关键点**：
- 优先尝试两级架构
- 失败则回退，保证兼容性
- 逐步扩展支持的场景

---

### 4. 场景分析 Prompt（新增）
**文件**：`server/app/formula_analyzer.py` 中的 prompt

```python
SCENARIO_ANALYSIS_PROMPT = """你是 Excel 公式场景分析器。你的任务是理解用户需求，识别属于哪种公式场景，并提取关键信息。

支持的场景类型：
1. lookup（精确查表）：根据某个值在字典表中查找对应值，一对一映射
   示例："根据员工ID查找部门"、"根据产品编号查价格"

2. priority_keyword_match（优先级+关键词匹配）：源单元格可能包含多个关键词（逗号或其他分隔符），需要检测所有关键词，按优先级返回最高级别
   示例："根据问题类型（可能有多个）返回最高错误级别"、"检测包含哪些标签并返回最重要的分类"

3. keyword_match（简单关键词匹配）：检测单元格是否包含某关键词，返回对应分类
   示例："如果包含'紧急'则标记为高优先级"

4. calculation（数学计算）：基于其他列进行数学运算
   示例："销售额 * 0.1"、"单价 * 数量"

5. complex（复杂逻辑）：以上场景都不适用，需要模型生成完整公式

你的输出必须是严格的 JSON 格式：
{
  "scenario": "lookup | priority_keyword_match | keyword_match | calculation | complex",
  "sourceColumn": "源列字母",
  "targetColumn": "目标列字母（从 activeCell 解析）",
  "dictionary": {
    "sourceFile": "外部文件名（如 tt.xlsx）或 null",
    "sheetName": "工作表名",
    "keyColumn": "键列字母",
    "valueColumn": "值列字母",
    "startRow": 2,
    "endRow": null
  },
  "priorities": ["优先级列表，从高到低"],
  "defaultValue": "默认值",
  "explanation": "场景判断说明"
}

判断要点：
- 如果用户描述提到"多个"、"逗号分隔"、"最高"、"优先级"，极可能是 priority_keyword_match
- 如果字典表存在且是简单的一对一映射，是 lookup
- 只返回 JSON，不要生成公式文本
"""

async def analyze_formula_scenario(request: GenerateFormulaRequest) -> FormulaAnalysis:
    """调用模型分析场景"""
    # 构建用户消息
    lines = [
        f"目标单元格：{request.activeCell}",
        f"需求：{request.description}",
    ]
    
    if request.headers and request.columns:
        mapping = "、".join(f"{col}列={h}" for col, h in zip(request.columns, request.headers))
        lines.append(f"列头映射：{mapping}")
    
    if request.dictionary:
        lines.append(f"\n当前工作簿内有字典表「{request.dictionary.name}」")
    
    if request.extraSheets:
        for extra in request.extraSheets:
            lines.append(
                f"\n外部工作簿「{extra.sourceFile}」工作表「{extra.sheetName}」"
                f"（列头：{', '.join(f'{c}={h}' for c, h in zip(extra.columns, extra.headers))}）"
            )
    
    lines.append("\n请分析这是什么场景，返回 JSON。")
    
    messages = [
        {"role": "system", "content": SCENARIO_ANALYSIS_PROMPT},
        {"role": "user", "content": "\n".join(lines)},
    ]
    
    # 调用模型
    config = formula_model_config() or selected_model_config(request.modelId)
    timeout = capability_float("llm", "timeoutSeconds")
    
    async with OpenAICompatibleClient(config, timeout=timeout) as client:
        response = await client.chat_completions(messages=messages, max_tokens=1000)
    
    content = response["choices"][0]["message"].get("content", "")
    
    # 解析 JSON
    import json, re
    json_match = re.search(r"\{.*\}", content, re.DOTALL)
    if not json_match:
        raise ValueError("模型未返回有效 JSON")
    
    data = json.loads(json_match.group(0))
    return FormulaAnalysis(**data)
```

---

## 实施步骤

### 阶段 1：基础架构（2小时）
1. 创建 `server/app/formula_analyzer.py`
2. 创建 `server/app/formula_generator.py`
3. 实现 `lookup` 和 `priority_keyword_match` 两个生成器

### 阶段 2：集成测试（1小时）
1. 修改 `server/app/rule_generator.py`，增加两级架构入口
2. 用真实场景测试（评分表 + tt.xlsx）
3. 验证生成的公式正确性

### 阶段 3：扩展场景（可选）
1. 实现 `keyword_match` 生成器
2. 实现 `calculation` 生成器
3. 添加更多边界情况处理

---

## 预期效果

### 当前版本问题
- ✗ 硬编码关键词（漏项）
- ✗ "/" 被替换成 ""
- ✗ 引用语法错误
- ✗ 成功率 ~70%

### 改造后效果
- ✓ 引用字典表完整区域
- ✓ "/" 等字面值 100% 保留
- ✓ 公式语法 100% 正确
- ✓ 结构化场景成功率 100%
- ✓ 复杂场景回退，不比现在差

---

## 给 Codex 的实施指令

请按以下顺序实施：

1. 创建 `server/app/formula_analyzer.py`，实现场景分析逻辑（包含 `SCENARIO_ANALYSIS_PROMPT` 和 `analyze_formula_scenario` 函数）

2. 创建 `server/app/formula_generator.py`，实现：
   - `_build_dict_range` 工具函数
   - `generate_lookup_formula` 函数
   - `generate_priority_keyword_formula` 函数
   - `generate_formula_from_analysis` 函数

3. 修改 `server/app/rule_generator.py` 的 `generate_formula` 函数，增加两级架构入口（在现有逻辑之前尝试场景分析）

4. 测试：用评分表场景（`tt.xlsx` 字典表 + 优先级判断）验证生成的公式是否正确

注意事项：
- 保持现有代码兼容，失败则回退
- 公式中的文件名、表名、列字母都要从 `request` 对象中精确提取，不要硬编码
- `priorities` 需要从字典表的值列中自动识别（去重+排序），或让模型在场景分析时提取

"""
/function 短链：AI 生成原生 Excel 公式
"""

import json
import re
from pydantic import BaseModel, Field

from .llm import formula_model_config, selected_model_config
from .llm.client import OpenAICompatibleClient
from .capabilities import capability_float, capability_int


# ---------------------------------------------------------------------------
# /function 短链：AI 生成原生 Excel 公式（主路径）
# ---------------------------------------------------------------------------


class DictionarySheet(BaseModel):
    """数据字典表上下文：名称 + 内容行（含表头）"""
    name: str = Field(min_length=1, max_length=100)
    rows: list[list[str]] = Field(default_factory=list, max_length=200)


class GenerateFormulaRequest(BaseModel):
    """生成原生公式请求"""
    description: str = Field(min_length=1, max_length=500, description="用户对公式的自然语言描述")
    activeCell: str = Field(min_length=1, max_length=32, description="公式将填入的起始单元格，如 E2")
    headers: list[str] = Field(default_factory=list, max_length=100, description="选区/活动表列头")
    columns: list[str] = Field(default_factory=list, max_length=100, description="列字母，与 headers 一一对应，如 [A,B,C]")
    sampleRows: list[list[str]] = Field(default_factory=list, max_length=10, description="选区样本行")
    dictionary: DictionarySheet | None = Field(default=None, description="数据字典表内容（若存在）")
    modelId: str | None = Field(default=None, description="模型 ID")


class GenerateFormulaResponse(BaseModel):
    """生成原生公式响应：一次返回现代版 + 兼容版两条公式，前端预览可切换"""
    modernFormula: str = Field(description="现代版公式（Excel 365/2021，可用 XLOOKUP/LET 等新函数），以 = 开头")
    modernExplanation: str = Field(default="", description="现代版公式的中文解释")
    compatFormula: str = Field(description="兼容版公式（Excel 2016/2019 通用函数），以 = 开头")
    compatExplanation: str = Field(default="", description="兼容版公式的中文解释")


def _build_formula_generation_prompt(request: GenerateFormulaRequest) -> list[dict]:
    """构建原生公式生成提示词（短链，单发，不接 planner）"""

    active_col = re.match(r"[A-Za-z]+", request.activeCell)
    active_col_letter = active_col.group(0).upper() if active_col else ""

    system_prompt = f"""你是 Excel 原生公式生成器。根据用户描述和提供的表格上下文，一次生成两条能直接填入指定单元格的原生 Excel 公式：一条「现代版」，一条「兼容版」。两条公式要产生相同结果，只是用不同的函数写法。

要求：
1. 两条公式都以 = 开头，使用相对引用对齐到目标单元格所在行（如目标是 E2 就引用 D2、Sheet1!A:B 等）。
2. 【最重要·循环引用】公式的引用区域绝对不能把目标单元格 {request.activeCell} 本身圈进去（否则循环引用、结果全错）。注意：这条只禁「引用范围包含目标格自己」，**不是**禁引用目标格所在的「{active_col_letter}」列。要分清两种常见场景：
   (a) 逐行输出（如目标 E2，往下每行一个结果）：读的是同一行其它列（D2、C2 等），此时引用 E 列会循环，务必避开。
   (b) 汇总/统计到一个格（如目标 B21，对上方 B2:B20 计数或求和）：这时**必须**引用目标格所在列的数据区（B2:B20），因为要统计的数据就在那一列——只要范围不含 B21 自己就不循环，这是完全正确的写法，不要因为"看见目标格在 B 列"就改去引用别的列。
   判断要统计/读取的数据到底在哪一列，严格按"列头→列字母"映射，别被目标格位置带偏。
3. 【现代版】面向 Excel 365/2021，可自由使用 XLOOKUP、LET、TEXTSPLIT、TEXTJOIN、FILTER、IFS、SWITCH 等新函数，写出最简洁清晰的公式。
4. 【兼容版】面向 Excel 2016/2019，只用这些老版本通用函数：IF、AND、OR、NOT、VLOOKUP、HLOOKUP、INDEX、MATCH、SUMPRODUCT、SUMIF(S)、COUNTIF(S)、SEARCH、FIND、ISNUMBER、IFERROR、LEFT、RIGHT、MID、LEN、TRIM、SUBSTITUTE、& 拼接、数组常量；不要用 LET/XLOOKUP/TEXTSPLIT/TEXTJOIN/CONCAT/IFS/SWITCH/FILTER/SORT/UNIQUE/LAMBDA 等新函数或动态数组溢出。多值判断改用 SUMPRODUCT+ISNUMBER+SEARCH 或嵌套 IF/COUNTIF 实现。
5. 若涉及跨表查表，直接在公式里引用那张表的区域（如 数据字典!$A$2:$B$7）。
6. 【易错点·必须遵守】COUNTIF/COUNTIFS/SUMIF/SUMIFS 的区域参数必须是真实的工作表单元格区域，绝不能传入 TEXTSPLIT/XLOOKUP/FILTER 等算出来的内存数组或 LET 变量（会 #VALUE! 报错）。要在内存数组里按条件计数，请改用 SUMPRODUCT((条件)*(条件))。
7. 【常见模式·首选】判断"某单元格是否含关键词清单中的任一项、并按其归类取值"时，两版都优先用 SUMPRODUCT(ISNUMBER(SEARCH(关键词区域, 目标格))*(级别区域=某值))>0 这种整块区域运算，不要先 TEXTSPLIT 切分再逐个查表（切分+数组查表最容易踩函数组合坑）。
8. 【回退·严格受限】仅当需求恰好是这两个已存在的预制规则能解决的复杂文本场景时，才回退用它们：清理隐形/零宽/控制字符 → =EB("隐形字符清理", 引用)；检测伪空白单元格 → =EB("空白检测", 引用)。这是仅有的两个预制规则，**绝不可编造任何其它规则名**。
9. 【做不了就明说】若原生公式和上述两个预制规则都无法表达该需求，请把 modernFormula 和 compatFormula 都返回空字符串 ""，并在 explanation 里写明「原生公式无法表达此需求，建议直接在对话中描述需求，走常规处理流程」。不要硬凑公式、不要编造规则。
10. 不要编造上下文里不存在的列或表名。
11. 只返回 JSON：{{"modernFormula": "=...", "modernExplanation": "现代版怎么工作", "compatFormula": "=...", "compatExplanation": "兼容版怎么工作"}}"""

    lines: list[str] = [
        f"目标单元格：{request.activeCell}（引用范围不得圈入此格本身；但若是对该列数据做汇总/统计，可正常引用同列其它单元格区域）",
        f"需求：{request.description}",
    ]

    if request.headers:
        lines.append("")
        if request.columns and len(request.columns) == len(request.headers):
            mapping = "、".join(
                f"{col}列={header}"
                for col, header in zip(request.columns, request.headers)
            )
            lines.append(f"列头→列字母映射：{mapping}")
        else:
            lines.append(f"当前表列头：{json.dumps(request.headers, ensure_ascii=False)}")

    if request.sampleRows:
        lines.append("样本行（前几行数据）：")
        for row in request.sampleRows:
            lines.append(f"  {json.dumps(row, ensure_ascii=False)}")

    if request.dictionary and request.dictionary.rows:
        lines.append("")
        lines.append(f"数据字典表「{request.dictionary.name}」内容（可在公式中引用此表）：")
        for row in request.dictionary.rows:
            lines.append(f"  {json.dumps(row, ensure_ascii=False)}")

    lines.append("")
    lines.append('请只返回 JSON：{"modernFormula": "=...", "modernExplanation": "...", "compatFormula": "=...", "compatExplanation": "..."}')

    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": "\n".join(lines)},
    ]


async def generate_formula(request: GenerateFormulaRequest) -> GenerateFormulaResponse:
    """短链生成原生 Excel 公式：单发调用模型，不经过 planner"""

    # 优先用「公式专用」连接（在模型配置里勾选），与聊天窗口全局选择脱钩，
    # 避免有人把全局切成推理模型时 /function 被 reasoning 拖到 length 截断。
    # 没有勾选任何专用连接时，才回退到全局选择逻辑。
    config = formula_model_config() or selected_model_config(request.modelId)
    if not config:
        raise ValueError("未配置模型")

    messages = _build_formula_generation_prompt(request)

    timeout = capability_float("llm", "timeoutSeconds")
    async with OpenAICompatibleClient(config, timeout=timeout) as client:
        # 凡是会输出 reasoning_content 的模型（DeepSeek v4/R1、kimi-k2 等，含 flash
        # 变体）都会把大量 token 花在思考上，预算太小会导致思考没结束就
        # finish_reason=length 截断，真正的 JSON（content）根本没开始写。
        # 走 capabilities 配置，别硬编码。
        max_tokens = capability_int("formula", "maxOutputTokens")
        response = await client.chat_completions(messages=messages, max_tokens=max_tokens)

    choice = response["choices"][0]
    message = choice.get("message", {})
    # content 可能为空：思考型模型把输出放在 reasoning_content。
    content = message.get("content") or message.get("reasoning_content") or ""

    modern_formula = ""
    modern_explanation = ""
    compat_formula = ""
    compat_explanation = ""

    # 去掉 markdown 代码围栏，再从整段里摳出首个 {...} JSON 对象
    # （模型常在 JSON 前后加说明文字，整段 json.loads 会失败）。
    cleaned = re.sub(r"```(?:json)?\s*", "", content)
    cleaned = re.sub(r"```", "", cleaned)
    json_match = re.search(r"\{.*\}", cleaned, re.DOTALL)
    if json_match:
        try:
            result = json.loads(json_match.group(0))
            modern_formula = str(result.get("modernFormula", "")).strip()
            modern_explanation = str(result.get("modernExplanation", "")).strip()
            compat_formula = str(result.get("compatFormula", "")).strip()
            compat_explanation = str(result.get("compatExplanation", "")).strip()
        except (json.JSONDecodeError, KeyError, TypeError):
            pass

    # 回退：JSON 解析失败（模型常忘转义公式内的引号）。逐字段用锚点抓值，
    # 兼容公式内含未转义的 " —— 非贪婪 + 后向锚点强制回溯到正确的收尾引号。
    def _field(key: str, next_keys: tuple[str, ...]) -> str:
        anchors = "".join(rf'|\s*,\s*"{k}"' for k in next_keys)
        pattern = rf'"{key}"\s*:\s*"(.*?)"\s*(?:{anchors[1:]}|\}})'
        m = re.search(pattern, cleaned, re.DOTALL)
        if not m:
            return ""
        raw = m.group(1)
        # 抓到的是 JSON 字符串体：模型转义正确的引号会留下字面 \" \\ \n 等，
        # 用 json.loads 反转义还原（如 \"子鹏\" → "子鹏"），失败则退回原文。
        try:
            return str(json.loads(f'"{raw}"')).strip()
        except (json.JSONDecodeError, ValueError):
            return raw.strip()

    if not modern_formula:
        modern_formula = _field(
            "modernFormula", ("modernExplanation", "compatFormula", "compatExplanation")
        )
    if not modern_explanation:
        modern_explanation = _field(
            "modernExplanation", ("compatFormula", "compatExplanation")
        )
    if not compat_formula:
        compat_formula = _field("compatFormula", ("compatExplanation",))
    if not compat_explanation:
        compat_explanation = _field("compatExplanation", ())

    # 兜底：只要有一版就用它补另一版，避免前端切换时出现空公式。
    if not modern_formula and compat_formula:
        modern_formula = compat_formula
        modern_explanation = compat_explanation
    if not compat_formula and modern_formula:
        compat_formula = modern_formula
        compat_explanation = modern_explanation

    # "做不了"是合法结果：模型可主动返回空公式 + 说明原生无法表达。
    # 只要拿到了任一说明文字，就原样透传空公式，让前端提示走常规对话流程。
    explanation_text = modern_explanation or compat_explanation
    if not modern_formula and not compat_formula and explanation_text:
        return GenerateFormulaResponse(
            modernFormula="",
            modernExplanation=explanation_text,
            compatFormula="",
            compatExplanation=explanation_text,
        )

    # 回退：整段无 JSON 结构，取首个以 = 开头的整行填两版。
    if not modern_formula and not compat_formula:
        for line in content.splitlines():
            stripped = line.strip()
            if stripped.startswith("="):
                modern_formula = compat_formula = stripped
                break

    if not modern_formula and not compat_formula:
        snippet = content.strip()[:200]
        finish = choice.get("finish_reason", "?")
        keys = ",".join(message.keys())
        # length 截断：多为推理型模型思考吃光 token 预算，JSON 还没写就被砍。
        # 给用户可操作的建议，而不是甩一段模型的半截推理原文。
        if finish == "length":
            raise ValueError(
                "模型输出被长度截断（finish_reason=length）：当前模型在思考上消耗了"
                "过多 token，未能写出公式。建议换用非推理型模型，或调大 "
                "config/capabilities.json 里 formula.maxOutputTokens。"
            )
        raise ValueError(
            f"模型未返回有效的公式（finish_reason={finish}, 字段={keys}）。"
            f"原文：{snippet}"
        )

    if not modern_formula.startswith("="):
        modern_formula = "=" + modern_formula
    if not compat_formula.startswith("="):
        compat_formula = "=" + compat_formula

    return GenerateFormulaResponse(
        modernFormula=modern_formula,
        modernExplanation=modern_explanation,
        compatFormula=compat_formula,
        compatExplanation=compat_explanation,
    )

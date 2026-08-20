"""
将工具调用架构集成到现有的公式生成流程。

调用方（rule_generator.generate_formula）优先走这里；任何失败都会抛出
异常，由调用方回退到原有的端到端模型生成逻辑，不让工具调用失败导致接口报错。
"""

from __future__ import annotations

import json
import logging
import re

from .capabilities import capability_float, capability_int
from .formula_tools.compiler import compile_formula
from .formula_tools.model_prompt import SYSTEM_PROMPT, build_user_message
from .formula_tools.schema import FormulaToolCallResponse
from .llm import formula_model_config, selected_model_config
from .llm.client import OpenAICompatibleClient
from .safety import dangerous_formula


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
    使用工具调用架构生成公式：模型返回结构化函数调用 JSON，代码编译成公式。

    Returns:
        {
            "modernFormula": "=...",
            "modernExplanation": "...",
            "compatFormula": "=...",  # 目前与 modern 相同
            "compatExplanation": "⚠️ 此公式使用了 Excel 365+ 函数，可能不兼容旧版本。..."
        }
    """
    logger.warning("=== [工具调用] extra_sheets 参数: %s ===", extra_sheets)

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

    config = formula_model_config() or selected_model_config(model_id)
    if not config:
        raise ValueError("未配置模型")

    timeout = capability_float("llm", "timeoutSeconds")
    max_tokens = capability_int("formula", "maxOutputTokens")

    try:
        async with OpenAICompatibleClient(config, timeout=timeout) as client:
            response = await client.chat_completions(
                messages=messages,
                max_tokens=max_tokens,
                temperature=0.1,  # 低温度保证结构输出稳定
            )
    except Exception as e:
        # 部分推理型模型（如 o1）不接受 temperature 参数，会 400 报错；
        # 这类模型走回退路径反而更合理，这里只做日志分级，便于后续针对性配置。
        model_label = getattr(config, "model", None) or model_id or "未知模型"
        if "temperature" in str(e).lower():
            logger.info("模型 %s 不支持 temperature 参数，回退到原路径", model_label)
        else:
            logger.warning("工具调用失败: %s，回退到原路径", e)
        raise

    message = response["choices"][0].get("message", {})
    # 思考型模型可能把内容放在 reasoning_content；与主流程保持一致。
    content = message.get("content") or message.get("reasoning_content") or ""

    try:
        json_match = _extract_json(content)
        data = json.loads(json_match)
        parsed = FormulaToolCallResponse.model_validate(data)
    except Exception as e:
        logger.error("解析工具调用 JSON 失败: %s\n内容: %s", e, content[:2000])
        raise ValueError(f"模型返回格式错误: {e}") from e

    try:
        formula = compile_formula(parsed.formula)
    except Exception as e:
        logger.error("编译公式失败: %s\nJSON: %s", e, parsed.formula.model_dump())
        raise ValueError(f"公式编译失败: {e}") from e

    # 与主流程相同的安全护栏：外部引用必须来自用户勾选的文件集合。
    allowed_external = {s["sourceFile"] for s in (extra_sheets or [])} or None
    logger.warning("=== [工具调用] allowed_external: %s ===", allowed_external)
    logger.warning("=== [工具调用] 生成的公式: %s ===", formula[:200])
    matched = dangerous_formula(formula, allowed_external=allowed_external)
    if matched is not None:
        logger.error("工具调用生成的公式未通过安全检查: %s", matched)
        raise ValueError(f"公式包含被禁用的函数：{matched}，已拒绝写入。")

    return {
        "modernFormula": formula,
        "modernExplanation": parsed.explanation,
        "compatFormula": formula,  # 与 modern 相同；兼容版转换（XLOOKUP→VLOOKUP）留待后续
        "compatExplanation": (
            "⚠️ 此公式使用了 Excel 365+ 函数，可能不兼容旧版本。"
            f"{parsed.explanation}"
        ),
    }


def _extract_json(text: str) -> str:
    """从文本中提取 JSON（支持 ```json 代码块与裸 JSON 对象）。"""
    code_block_match = re.search(r"```(?:json)?\s*\n(.*?)\n```", text, re.DOTALL)
    if code_block_match:
        return code_block_match.group(1).strip()

    json_match = re.search(r"\{.*\}", text, re.DOTALL)
    if json_match:
        return json_match.group(0)

    raise ValueError("未找到 JSON")

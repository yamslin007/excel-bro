"""AI 辅助计划编辑服务。

用户提供当前计划和自然语言描述，AI 返回修改建议和完整的新计划 JSON。
"""

from __future__ import annotations

from typing import Any

from .llm import OpenAICompatibleClient
from .llm.config import ModelConnection
from .models import AnalysisPlan


async def chat_modify_plan(
    current_plan: AnalysisPlan,
    user_message: str,
    connection: ModelConnection,
    timeout: float,
) -> tuple[str, AnalysisPlan]:
    """AI 辅助修改计划。

    Args:
        current_plan: 当前的计划 JSON
        user_message: 用户的修改需求描述
        connection: 模型连接配置
        timeout: 超时时间（秒）

    Returns:
        (AI 回复文本, 修改后的完整计划)
    """
    system_prompt = """你是 Excel Bro 计划编辑助手。用户会给你一份当前的执行计划（JSON 格式）和他们想要的修改，你需要：

1. 理解用户的修改需求
2. 基于当前计划进行修改
3. 返回完整的修改后计划 JSON

执行计划结构说明：
- `steps`: 执行步骤数组，每个步骤包含操作类型（如 filter、sort、group 等）
- `acceptanceCriteria`: 验收条件（由服务端自动计算，你不需要修改）
- `sourceFingerprint`: 数据源指纹（由服务端自动计算，你不需要修改）
- `liveness`: 计划活性配置（参数求值时机）

常见操作类型：
- filter: 筛选数据
- sort: 排序
- group: 分组统计
- pivot: 透视表
- compute: 计算新列
- take: 取前 N 行
- select: 选择特定列

请用中文简洁解释你做了什么修改，然后提供完整的修改后的计划 JSON。

重要：请直接返回 JSON 格式的响应，包含两个字段：
- explanation: 修改说明（2-3 句话）
- modifiedPlan: 完整的修改后计划对象"""

    user_prompt = f"""当前计划：
```json
{current_plan.model_dump_json(indent=2, exclude_none=True)}
```

用户需求：{user_message}"""

    async with OpenAICompatibleClient(connection, timeout=timeout) as client:
        response, _ = await client.chat_completions(
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.3,
            tools=[{
                "type": "function",
                "function": {
                    "name": "return_modified_plan",
                    "description": "返回修改说明和完整的修改后计划",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "explanation": {
                                "type": "string",
                                "description": "修改说明，2-3 句中文"
                            },
                            "modifiedPlan": {
                                "type": "object",
                                "description": "完整的修改后计划 JSON"
                            }
                        },
                        "required": ["explanation", "modifiedPlan"]
                    }
                }
            }],
            tool_choice="auto"
        )

    # 解析工具调用结果
    import json
    message = response.get("choices", [{}])[0].get("message", {})
    tool_calls = message.get("tool_calls", [])

    if not tool_calls:
        raise ValueError("AI 未返回修改后的计划")

    arguments = json.loads(tool_calls[0]["function"]["arguments"])
    explanation = arguments.get("explanation", "已完成修改")
    modified_plan_dict = arguments.get("modifiedPlan", {})

    # 保留原有的 acceptanceCriteria 和 sourceFingerprint
    # 这些字段由服务端重新计算，AI 不应该修改
    if "acceptanceCriteria" in modified_plan_dict:
        del modified_plan_dict["acceptanceCriteria"]
    if current_plan.sourceFingerprint:
        modified_plan_dict["sourceFingerprint"] = current_plan.sourceFingerprint

    # 解析为 AnalysisPlan（会触发 Pydantic 校验）
    modified_plan = AnalysisPlan.model_validate(modified_plan_dict)

    return explanation, modified_plan

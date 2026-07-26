from __future__ import annotations

import json
import os
import re
import uuid
from typing import Any

from pydantic import ValidationError

from .capabilities import capability_int
from .llm import ModelConnection, OpenAICompatibleClient
from .models import (
    IntentCheckRequest,
    IntentCheckResponse,
    IntentClarificationResponse,
    IntentProceedResponse,
    IntentToolResponse,
)


INTENT_SYSTEM_PROMPT = """
你是 Excel Bro 的需求确认器。你只判断用户需求是否足够明确，不能分析数据、
计算结果或生成 Excel 操作计划。

仅当存在会实质改变答案或写入结果的多种合理解释时，返回 clarification。
不要因为措辞简短、语气口语化或缺少不影响结果的细节而追问。

重点检查：
- 是否明确使用当前表、用户选中的多张表，还是跨表汇总；
- 字段、分组层级、指标计算口径，尤其是占比的分子和分母；
- 输出是一张汇总表、每组一张表，还是修改原表；
- 删除、覆盖、清空等高风险动作的目标是否明确。

用户手动选中的工作表是默认且不可扩大的数据边界。不要建议读取未选中的工作表。
query_table 目前只用于 sourceMode=workbook；文件夹来源返回 proceed。
优先结合“上一轮结构化意图”和“上一轮紧凑结果”理解“那最低的呢、换成平均值、只看某地区”等承接式追问，
不得要求用户重复已经明确的信息。
用户完成一次确认后，不要重复询问同一个问题；但自定义回答仍存在新的关键歧义时，
可以继续追问。不得超过请求给出的澄清轮次上限；达到上限后，只读查询应采用最合理口径继续并明确假设；
写入需求应返回 proceed 交给后续预览，不得直接执行写入。
若提供了“上次工具失败”，请根据错误和可用字段修正工具参数，不要原样重试。
如果需求明确且属于查询、筛选、比较、统计或汇总，返回 tool_request，让 Excel
本地工具完成确定性计算。只有不需要读取数据即可规划的格式、结构或写入操作才返回
proceed。如果不明确，只问一个最能消除关键歧义的问题，提供 2 到 4 个互斥选项。
所有内容使用简体中文。

只返回 JSON，不要使用 Markdown。格式：
{"kind":"proceed","summary":"我理解的需求","confirmedPrompt":"明确化后的完整需求"}
或
{"kind":"clarification","clarification":{"id":"短标识","summary":"当前理解",
"question":"需要确认的问题","reason":"为什么需要确认","scopeLabel":"数据范围说明",
"options":[{"id":"选项标识","label":"短标签","description":"解释",
"resolution":"选择后追加到原需求的明确说明","action":"resolve|editScope"}]}}
或
{"kind":"tool_request","summary":"准备执行的本地查询",
"confirmedPrompt":"明确化后的完整需求",
"request":{"id":"短标识","tool":"query_table","arguments":{
"mode":"rows|aggregate|profile","scope":"selected|active",
"fields":["需要返回的字段"],"filters":[{"field":"字段",
"operator":"equals|notEquals|contains|greaterThan|greaterThanOrEqual|lessThan|lessThanOrEqual|isBlank|isNotBlank",
"value":"可选值"}],"groupBy":["分组字段"],"metrics":[{"operation":"countRows|countDistinct|sum|average|min|max",
"field":"countRows 可省略，其余必填","outputName":"输出字段名","ratioOutputName":"可选，占总量比例的输出名"}],
"profileField":"profile 模式的字段","sortBy":"字段或指标名",
"sortDirection":"asc|desc","limit":20}}}
""".strip()
MAX_CLARIFICATION_ROUNDS = capability_int(
    "conversation", "maxClarificationRounds"
)


def _scope_label(request: IntentCheckRequest) -> str:
    names = [sheet.name for sheet in request.scope.sheets]
    if request.scope.selectionMode == "auto":
        return f"当前工作表：{names[0]}"
    if len(names) <= 3:
        return f"已选择 {len(names)} 张工作表：{'、'.join(names)}"
    return f"已选择 {len(names)} 张工作表：{'、'.join(names[:3])} 等"


def local_intent_check(
    request: IntentCheckRequest,
) -> IntentCheckResponse | None:
    # 业务语义一律交给模型。本地层只负责协议、边界和工具参数校验。
    return None


def _parse_model_response(
    content: str, request: IntentCheckRequest
) -> IntentCheckResponse:
    text = re.sub(r"^```(?:json)?\s*", "", content.strip())
    text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("模型没有返回有效的意图判断")
    payload = json.loads(text[start : end + 1])
    if payload.get("kind") == "proceed":
        return IntentProceedResponse.model_validate(
            {**payload, "provider": "model"}
        )
    if payload.get("kind") == "clarification":
        clarification = dict(payload.get("clarification") or {})
        clarification.setdefault("id", f"intent-{uuid.uuid4().hex[:10]}")
        clarification["scopeLabel"] = _scope_label(request)
        return IntentClarificationResponse.model_validate(
            {
                "kind": "clarification",
                "provider": "model",
                "clarification": clarification,
            }
        )
    if payload.get("kind") == "tool_request":
        return IntentToolResponse.model_validate(
            {**payload, "provider": "model"}
        )
    raise ValueError("模型返回了未知的意图判断类型")


async def check_intent(
    request: IntentCheckRequest,
    *,
    config: ModelConnection | None,
) -> IntentCheckResponse:
    deterministic = local_intent_check(request)
    if deterministic is not None:
        return deterministic
    if config is None:
        return IntentProceedResponse(
            provider="local",
            summary=request.prompt,
            confirmedPrompt=request.prompt,
        )

    scope_payload: dict[str, Any] = request.scope.model_dump()
    user_content = (
        f"用户需求：{request.prompt}\n"
        f"用户是否回答过澄清问题：{request.intentConfirmed}\n"
        f"当前澄清轮次：{request.clarificationRound}/{MAX_CLARIFICATION_ROUNDS}\n"
        f"附件图片数量：{request.imageCount}\n"
        f"最近对话：{json.dumps([item.model_dump() for item in request.conversation], ensure_ascii=False)}\n"
        f"上一轮结构化意图：{json.dumps(request.priorIntent.model_dump() if request.priorIntent else None, ensure_ascii=False)}\n"
        f"上一轮紧凑结果：{json.dumps(request.priorResult.model_dump() if request.priorResult else None, ensure_ascii=False)}\n"
        f"上次工具失败：{json.dumps(request.toolFailure.model_dump() if request.toolFailure else None, ensure_ascii=False)}\n"
        f"数据范围结构摘要（不含数据行）：{json.dumps(scope_payload, ensure_ascii=False)}"
    )
    full_timeout = float(os.getenv("AI_TIMEOUT_SECONDS", "60"))
    timeout = float(
        os.getenv(
            "AI_INTENT_TIMEOUT_SECONDS",
            str(max(10.0, min(full_timeout, 60.0))),
        )
    )
    async with OpenAICompatibleClient(config, timeout=timeout) as client:
        repair_note = ""
        for _attempt in range(2):
            payload = await client.chat_completions(
                messages=[
                    {"role": "system", "content": INTENT_SYSTEM_PROMPT},
                    {
                        "role": "user",
                        "content": f"{user_content}{repair_note}",
                    },
                ],
                max_tokens=1200,
            )
            content = ""
            try:
                content = (
                    payload["choices"][0]["message"].get("content")
                    or ""
                )
                parsed = _parse_model_response(content, request)
                if (
                    parsed.kind == "clarification"
                    and request.clarificationRound
                    >= MAX_CLARIFICATION_ROUNDS
                ):
                    raise ValueError("已达到澄清轮次上限，不能继续返回 clarification")
                return parsed
            except (
                KeyError,
                IndexError,
                TypeError,
                ValueError,
                ValidationError,
                json.JSONDecodeError,
            ) as error:
                repair_note = (
                    "\n\n你上一次返回的内容未通过结构化协议校验。"
                    f"\n校验错误：{error}"
                    f"\n无效输出：{content[:1200]}"
                    "\n请只返回一个符合系统规定格式的 JSON 对象，不要解释。"
                )

    # 意图路由失败不应阻断用户。后续计划仍会先展示预览，避免直接写入。
    return IntentProceedResponse(
        provider="local",
        summary="意图路由未能生成有效结构，已按原需求进入安全预览。",
        confirmedPrompt=request.prompt,
    )

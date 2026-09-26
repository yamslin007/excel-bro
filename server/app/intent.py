from __future__ import annotations

import json
import os
import re
import uuid
from typing import Any

from pydantic import ValidationError

from .capabilities import capability_int, model_timeout_seconds
from .llm import ModelConnection, ModelNotConfiguredError, OpenAICompatibleClient
from .models import (
    IntentCheckRequest,
    IntentCheckResponse,
    IntentClarificationResponse,
    IntentConversationResponse,
    IntentProceedResponse,
    IntentToolResponse,
)
from .structural_intent import parse_bulk_worksheet_deletion


INTENT_SYSTEM_PROMPT = """
你是 Excel Bro 的对话路由器。你判断用户是在普通交谈，还是提出了工作簿任务；
不能分析数据、计算结果或生成 Excel 操作计划。

普通问候、感谢、寒暄、询问你能做什么等不要求操作或读取工作簿的消息，
返回 conversation，并直接给出自然、简短的回复。不要为这类消息返回澄清卡、
查询工具或操作计划。介绍能力时只概述可对已授权工作表进行查询与分析，
写入需先生成预览并经用户确认；不要假装已经读取表格，或承诺未验证的功能。
用户提出了具体的工作簿任务时，不要用 conversation 敷衍；按下述规则选择
tool_request、proceed 或真正必要的 clarification。

仅当存在会实质改变答案或写入结果的多种合理解释时，返回 clarification。
不要因为措辞简短、语气口语化或缺少不影响结果的细节而追问。

重点检查：
- 是否明确使用当前表、用户选中的多张表，还是跨表汇总；
- 字段、分组层级、指标计算口径，尤其是占比的分子和分母；
- 输出是一张汇总表、每组一张表，还是修改原表；
- 删除、覆盖、清空等高风险动作的目标是否明确。

拆分需求中，“输出工作表名称来自哪个字段”决定拆分维度；“每张结果表内部按
哪个字段分组”决定表内聚合维度。若用户一处说“按 A 拆分”，另一处又要求工作表
按 B 命名，不要把“B 分列显示”这类含义不清的短语作为选项，也不要预设其中一句
一定优先。应使用对称、可直接执行的选项，例如“按 B 建表，每张表内按 A 汇总”
与“按 A 建表，每张表内按 B 汇总”，并在 resolution 中完整保留已明确的指标口径。
若占比的分子是“当前拆分值在某个表内分组中的记录数”，用户所说的“该分组全部
记录数、全部项目数”应理解为分母覆盖该分组中的全部拆分值。不要提供“分母仍为
当前拆分值在该分组中的记录数”这种会让每个非空占比恒为 100% 的选项，除非用户
明确要求分子与分母使用同一数据范围。
拆分选项的 resolution 必须只指定一个工作表命名字段；辅助标识应作为表内字段，
不要擅自把多个字段拼接成拆分键。拆分字段会自动出现在结果列中，不要再把它描述为
表内分组字段。

用户手动选中的工作表是默认且不可扩大的数据边界。不要建议读取未选中的工作表。
数据范围摘要中的 worksheetNames 列出全部数据工作表名（不含系统表），仅供规划
结构操作（删除、重命名等用户明确要求的工作表）；读取数据仍以 sheets 中的已选
工作表为限。
query_table 同时支持 sourceMode=workbook 和 sourceMode=folder；文件夹模式由受控
pandas 工具读取用户已选择的文件和工作表。
query_table 的 combine.deduplicate / combine.join 仅用于文件夹模式的多表去重或关联；
单表「去重」（修改原表）不是只读查询，不要用 query_table 的 combine 表达。
当前用户需求始终是本轮唯一主任务。最近对话、上一轮结构化意图和上一轮紧凑结果只用于
补全“那最低的呢、换成平均值、只看某地区”等明显依赖前文的省略式追问。只要当前输入
自身已经包含明确的动作和目标，就必须覆盖上一轮任务，不得把旧的拆分、汇总、写入等动作
合并进来，也不得用上一轮 confirmedPrompt 替换当前输入。
用户完成一次确认后，不要重复询问同一个问题；但自定义回答仍存在新的关键歧义时，
可以继续追问。不得超过请求给出的澄清轮次上限；达到上限后，只读查询应采用最合理口径继续并明确假设；
写入需求应返回 proceed 交给后续预览，不得直接执行写入。
若提供了“上次工具失败”，请根据错误和可用字段修正工具参数，不要原样重试。
如果需求明确且属于查询、筛选、比较、统计或汇总，返回 tool_request，让 Excel
本地工具完成确定性计算。只有不需要读取数据即可规划的格式、结构或写入操作才返回
proceed。「去重」「删除重复行」「删除 XX 的行」「清空 XX 的行」都会修改原表，属于
写入操作，应返回 proceed（交给规划 Agent 生成 removeDuplicates / deleteRange /
clearRange 动作），不要返回 tool_request。如果不明确，只问一个最能消除关键歧义的
问题，提供 2 到 4 个互斥选项。每个选项必须标明 resolutionMode：
若该选项会推翻或改变原需求中的拆分维度、输出结构、指标口径或目标，使用 replace，
且 resolution 必须是可独立交给规划 Agent 的完整最终需求；只补充原需求中缺失的
参数时使用 append。
路由示例：“某个对象有多少条记录”是只读统计，必须返回 tool_request，使用名称字段
筛选并以 countRows 聚合；“统计记录数并写入汇总表”明确要求修改工作簿，返回 proceed。
所有内容使用简体中文。

只返回 JSON，不要使用 Markdown。格式：
{"kind":"conversation","message":"直接给用户的简短回复"}
或
{"kind":"proceed","summary":"我理解的需求","confirmedPrompt":"明确化后的完整需求"}
或
{"kind":"clarification","clarification":{"id":"短标识","summary":"当前理解",
"question":"需要确认的问题","reason":"为什么需要确认","scopeLabel":"数据范围说明",
"options":[{"id":"选项标识","label":"短标签","description":"解释",
"resolution":"选择后的完整或补充说明","resolutionMode":"append|replace",
"action":"resolve|editScope"}]}}
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
"sortDirection":"asc|desc","limit":20,
"combine":{"mode":"union|deduplicate|join","deduplicateBy":["去重字段"],
"leftSourceSheetId":"关联左表工作表ID","rightSourceSheetId":"关联右表工作表ID",
"leftKey":"左键","rightKey":"右键","joinHow":"inner|left|right|outer"}}}}
""".strip()
MAX_CLARIFICATION_ROUNDS = capability_int(
    "conversation", "maxClarificationRounds"
)

_READ_ONLY_QUERY_PATTERN = re.compile(
    r"(?:有多少(?:条|行|个)?(?:记录|数据)?|多少(?:条|行|个)(?:记录|数据)|"
    r"记录数|数量|查询|查找|查一下|看看|看一下|列出|哪些|哪个|最高|最低|"
    r"平均|合计|总计|占比|统计|汇总|比较|筛选)"
)
_MUTATING_INTENT_PATTERN = re.compile(
    r"(?:写入|写到|写回|填入|填写|放到|放入|贴到|导出|修改|更新|替换|覆盖|删除|删掉|清除|清空|"
    r"去重|拆分|分表|新建|创建|插入|追加|移动|复制)"
    r"|(?:生成|建立|制作|排序|设置|调整|合并).{0,16}"
    r"(?:表|工作表|单元格|区域|列|行|公式|格式|样式|颜色|图表|透视表|文件)"
)


def is_explicit_read_only_query(prompt: str) -> bool:
    """Return True when the current request clearly asks only for information.

    This is a generic safety gate, not a business-field parser. It prevents a
    model formatting error from turning an obvious query into an Excel plan;
    the model still decides which authorized fields and filters query_table
    should use.
    """

    normalized = re.sub(r"\s+", "", prompt)
    return bool(
        normalized
        and _READ_ONLY_QUERY_PATTERN.search(normalized)
        and not _MUTATING_INTENT_PATTERN.search(normalized)
    )


def _read_only_routing_fallback(
    request: IntentCheckRequest,
) -> IntentClarificationResponse:
    """Fail closed when a read-only query cannot be routed safely."""

    return IntentClarificationResponse.model_validate(
        {
            "kind": "clarification",
            "provider": "local",
            "clarification": {
                "id": f"readonly-{uuid.uuid4().hex[:10]}",
                "summary": "这是只读查询，本次不会写入 Excel。",
                "question": "意图识别未能生成有效查询，是否按当前数据范围重新查询？",
                "reason": "为避免把查询误转成写入操作，需要重新生成只读查询参数。",
                "scopeLabel": _scope_label(request),
                "options": [
                    {
                        "id": "retry-read-only",
                        "label": "按当前范围查询",
                        "description": "仅重新读取和统计，不修改工作簿。",
                        "resolution": (
                            "仅执行只读查询并直接回答，不生成或执行任何 Excel 写入计划。"
                            f"原始查询：{request.prompt}"
                        ),
                        "resolutionMode": "replace",
                        "action": "resolve",
                    },
                    {
                        "id": "edit-scope",
                        "label": "修改数据范围",
                        "description": "重新选择需要查询的工作表。",
                        "resolution": "修改本次只读查询的数据范围",
                        "resolutionMode": "append",
                        "action": "editScope",
                    },
                ],
            },
        }
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
    bulk_deletion = parse_bulk_worksheet_deletion(request.prompt)
    if bulk_deletion is not None:
        keeper_name = bulk_deletion.keeper_name
        return IntentProceedResponse(
            provider="local",
            summary=f"删除除「{keeper_name}」外的所有其他工作表。",
            confirmedPrompt=f"删除除「{keeper_name}」外的所有其他工作表",
        )
    # 其他业务语义交给模型。本地层只处理可确定识别的高风险结构命令。
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
    if payload.get("kind") == "conversation":
        return IntentConversationResponse.model_validate(
            {**payload, "provider": "model"}
        )
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
    read_only_required = is_explicit_read_only_query(request.prompt)
    if config is None:
        # 普通 Agent 请求缺少模型时必须明确失败，不能静默把原需求当作
        # “已确认”放行——那会让用户以为模型已经理解了这个需求。
        raise ModelNotConfiguredError(
            "尚未配置模型，无法理解这条需求。请点击「添加模型」，"
            "填写服务地址、模型 ID 和 API Key 后重试。"
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
        f"本地只读安全判定：{'明确只读查询，禁止返回 proceed' if read_only_required else '未锁定为只读'}\n"
        f"数据范围结构摘要（不含数据行）：{json.dumps(scope_payload, ensure_ascii=False)}"
    )
    full_timeout = model_timeout_seconds()
    timeout = float(
        os.getenv(
            "AI_INTENT_TIMEOUT_SECONDS",
            str(max(10.0, min(full_timeout, 60.0))),
        )
    )
    async with OpenAICompatibleClient(config, timeout=timeout) as client:
        repair_note = ""
        for _attempt in range(2):
            payload, _ = await client.chat_completions(
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
                if read_only_required and parsed.kind == "proceed":
                    raise ValueError(
                        "当前需求是明确只读查询，必须返回 tool_request，不能进入 Excel 规划"
                    )
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

    if read_only_required:
        return _read_only_routing_fallback(request)

    # 写入意图路由失败不应阻断用户。后续计划仍会先展示预览，避免直接写入。
    return IntentProceedResponse(
        provider="local",
        summary="意图路由未能生成有效结构，已按原需求进入安全预览。",
        confirmedPrompt=request.prompt,
    )

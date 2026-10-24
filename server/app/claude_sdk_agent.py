"""Claude Agent SDK orchestrator: structure and explicit user images only.

This adapter never gives the SDK workbook data rows, a live-workbook handle,
filesystem/command tools, or an execution tool. Its only terminal tools return
an existing EB answer, local computation rule, or validated preview plan.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import sys
import uuid
from pathlib import Path
from typing import Any, Awaitable, Callable

from .capabilities import capability_float, capability_int
from .assistant_identity import with_runtime_model_identity
from .task_context import TASK_CONTEXT_PROMPT, task_context_payload
from .agent_runtime import AGENT_DECISION_PROMPT
from .plan_tools import PLAN_COMPILERS, compile_plan_tool, plan_compiler_definitions
from .llm.config import DEEPSEEK_ANTHROPIC_URL, ModelConnection, ModelNotAvailableError
from .llm.errors import (
    LLMError, LLMHTTPStatusError, LLMResponseError, LLMTimeoutError,
    LLMUsageLimitError, model_usage_error_code,
)
from .models import (
    AssistantResponse,
    ImageAttachment,
    IntentCheckRequest,
    IntentCheckResponse,
    MessageUnderstandingRequest,
    MessageUnderstandingResponse,
    PlanRequest,
    QueryTableArguments,
)
from .result_submission import (
    RESULT_REF_INVALID_CODE,
    RESULT_REF_INVALID_MESSAGE,
    LocalResultIncomplete,
    ResultSubmissionError,
    build_local_result_answer,
    build_result_plan_response,
    evaluate_local_result_submission,
    local_result_card as _local_result_card,
    normalize_planning_goal,
    resolve_result_ref,
)

logger = logging.getLogger(__name__)


def _sdk() -> Any:
    try:
        import claude_agent_sdk
    except ImportError as error:
        raise ModelNotAvailableError(
            "Claude Agent SDK 尚未安装。请在服务端环境运行 "
            "python -m pip install --only-binary=claude-agent-sdk "
            "-r server/requirements-claude-sdk.txt。"
        ) from error
    return claude_agent_sdk


def _state_dir() -> Path:
    raw = os.getenv("CLAUDE_AGENT_SDK_STATE_DIR", "").strip()
    path = Path(raw)
    repository = Path(__file__).resolve().parents[2]
    if not raw or not path.is_absolute():
        raise ModelNotAvailableError("Claude SDK 会话目录必须是仓库外的绝对路径。")
    path = path.resolve()
    if path == repository or repository in path.parents:
        raise ModelNotAvailableError("Claude SDK 会话目录必须是仓库外的绝对路径。")
    path.mkdir(mode=0o700, parents=True, exist_ok=True)
    return path


def _context_confirmations(request: PlanRequest) -> list[dict[str, str]]:
    """把 `taskContext.confirmations` 投影成检查点允许的三字段结构。

    只保留"问题标识 / 字段名 / 用户回答"，不保留任何来源文本或数据行。
    上限与 `planning_context.MAX_CONFIRMATIONS` 一致；解析不了的条目直接丢弃，
    不用一个残缺的确认去冒充"用户已经确认过"。
    """
    context = request.taskContext
    if context is None:
        return []
    projected: list[dict[str, str]] = []
    for item in context.confirmations:
        question_id = (item.questionId or item.label or "").strip()
        if not question_id:
            continue
        projected.append({
            "questionId": question_id[:80],
            "label": (item.label or "").strip()[:80],
            "answerLabel": (item.value or "").strip()[:400],
        })
    return projected[:8]


_RESUME_NOTE = (
    "继续原任务尚未完成的步骤；这是新的模型调用，不是恢复上一次的模型会话。"
    "已完成的查询与 Python 分析优先复用下面的结果引用，不要重复读取或计算。"
    "若分析已完成而写入计划未完成，使用同一结果引用调用 submit_result_plan；宿主校验后执行一次，再供用户审阅。"
    "下面只给出已确认的口径与有归属的结果引用；引用可能已经过期，"
    "取不到时必须重新读取来源，不得假称结果仍然可用。"
    "不要因为这里列过某个结果就跳过授权范围与来源核对。"
)


def _resume_projection(resume_context: dict[str, Any] | None) -> dict[str, Any] | None:
    """把检查点投影成**只含白名单字符串**的续跑提示。

    任何值矩阵、计划、错误原文、路径都不会进来：只取 `questionId` / `label` /
    `answerLabel` 与 `res_[0-9a-f]{32}` 形式的引用。
    """
    if not isinstance(resume_context, dict):
        return None
    confirmations: list[dict[str, str]] = []
    for item in (resume_context.get("confirmations") or [])[:8]:
        if not isinstance(item, dict):
            continue
        question_id = str(item.get("questionId") or "").strip()
        if not question_id:
            continue
        confirmations.append({
            "questionId": question_id[:80],
            "label": str(item.get("label") or "").strip()[:80],
            "answerLabel": str(item.get("answerLabel") or "").strip()[:400],
        })
    refs: list[str] = []
    for item in (resume_context.get("resultRefs") or [])[:12]:
        ref = item.get("resultRef") if isinstance(item, dict) else item
        if (
            isinstance(ref, str)
            and ref.startswith("res_")
            and len(ref) == 36
            and all(char in "0123456789abcdef" for char in ref[4:])
        ):
            refs.append(ref)
    completed = [
        str(name)[:80]
        for name in (resume_context.get("completedTools") or [])[:32]
        if str(name).strip()
    ]
    if not confirmations and not refs and not completed:
        return None
    return {
        "mode": "continue_remaining",
        "note": _RESUME_NOTE,
        "confirmedDecisions": confirmations,
        "previousResultRefs": refs,
        "completedTools": completed,
    }


def _options(
    sdk: Any, connection: ModelConnection, system_prompt: str,
    *, max_turns: int, server: Any | None = None, allowed: list[str] | None = None,
) -> Any:
    state_dir = _state_dir()
    sdk_env = {
        "ANTHROPIC_API_KEY": connection.api_key,
        "ANTHROPIC_AUTH_TOKEN": "",
        "ANTHROPIC_BASE_URL": connection.base_url,
        "CLAUDE_CONFIG_DIR": str(state_dir),
        "CLAUDE_CODE_DISABLE_NONESSENTIAL_TRAFFIC": "1",
    }
    if connection.base_url == DEEPSEEK_ANTHROPIC_URL:
        # Claude Code can make internal model-class requests; keep those on the
        # user's selected DeepSeek model rather than implicitly selecting Pro.
        sdk_env.update({
            "ANTHROPIC_AUTH_TOKEN": connection.api_key,
            "ANTHROPIC_MODEL": connection.model,
            "ANTHROPIC_DEFAULT_OPUS_MODEL": connection.model,
            "ANTHROPIC_DEFAULT_SONNET_MODEL": connection.model,
            "ANTHROPIC_DEFAULT_HAIKU_MODEL": connection.model,
            "CLAUDE_CODE_SUBAGENT_MODEL": connection.model,
        })
    return sdk.ClaudeAgentOptions(
        model=connection.model,
        system_prompt=with_runtime_model_identity(system_prompt, connection),
        tools=[],  # No built-in Read, Bash, Edit, network or agent tools.
        disallowed_tools=["Read", "Bash", "Write", "Edit", "Glob", "Grep", "WebFetch", "WebSearch", "Agent", "Skill"],
        mcp_servers={"excel_bro": server} if server else {},
        strict_mcp_config=True,
        setting_sources=[],
        allowed_tools=allowed or [],
        permission_mode="dontAsk",
        cwd=str(state_dir),  # Do not discover repository CLAUDE.md or project files.
        max_turns=max_turns,
        max_budget_usd=capability_float("agent", "claudeSdkMaxBudgetUsd"),
        env=sdk_env,
        extra_args={"no-session-persistence": None},
    )


async def _image_prompt(text: str, images: list[ImageAttachment]):
    """User attachments only: no temporary files or filesystem tools."""
    yield {"type": "user", "message": {"role": "user", "content": [
        {"type": "text", "text": text},
        *[{"type": "image", "source": {
            "type": "base64", "media_type": image.mediaType, "data": image.data,
        }} for image in images],
    ]}}


def _sdk_failure(result: Any, assistant_error: str | None = None) -> LLMError:
    """Retain structured SDK limits without disclosing CLI text or prompts."""
    subtype_codes = {
        "error_max_budget_usd": "MODEL_BUDGET_LIMIT",
        "error_max_turns": "MODEL_TURN_LIMIT",
    }
    code = subtype_codes.get(getattr(result, "subtype", None))
    if code:
        return LLMUsageLimitError(code)
    if assistant_error == "billing_error":
        return LLMUsageLimitError("MODEL_QUOTA_EXHAUSTED")
    details = getattr(result, "errors", None) or []
    if isinstance(details, str):
        details = [details]
    for detail in [*details, getattr(result, "result", None)]:
        if isinstance(detail, str):
            code = model_usage_error_code(detail)
            if code:
                return LLMUsageLimitError(code)
    if getattr(result, "stop_reason", None) == "max_tokens":
        return LLMUsageLimitError("MODEL_OUTPUT_LIMIT")
    status = getattr(result, "api_error_status", None)
    if isinstance(status, int) and status >= 400:
        return LLMHTTPStatusError(status, "")
    if assistant_error == "rate_limit":
        return LLMHTTPStatusError(429, "")
    return LLMResponseError("Claude Agent SDK 未能完成本轮请求。")


async def _result(sdk: Any, prompt: Any, options: Any, timeout: float) -> str:
    # Uvicorn's Windows reload/multi-worker mode installs a Selector loop,
    # which has no subprocess transport. Fail before invoking the SDK rather
    # than misreporting NotImplementedError as a model-network failure.
    if sys.platform == "win32" and isinstance(
        asyncio.get_running_loop(), asyncio.SelectorEventLoop
    ):
        raise ModelNotAvailableError(
            "Agent SDK 无法在当前 Windows 热重载事件循环中启动子进程。"
            "请停止现有后端，再运行 powershell -NoProfile -ExecutionPolicy Bypass "
            "-File .\\scripts\\start_claude_sdk_server.ps1（等同 npm run dev:server）；只启动一个服务。"
        )
    final = None
    assistant_error = None
    assistant_stop_reason = None
    try:
        async with asyncio.timeout(timeout):
            async for item in sdk.query(prompt=prompt, options=options):
                if isinstance(item, sdk.ResultMessage):
                    final = item
                elif isinstance(item, getattr(sdk, "AssistantMessage", ())):
                    # Keep the last attempt only; a successful retry supersedes
                    # an earlier provider error. Never inspect normal prose.
                    assistant_error = getattr(item, "error", None)
                    assistant_stop_reason = getattr(item, "stop_reason", None)
    except TimeoutError as error:
        raise LLMTimeoutError("Claude Agent SDK 处理超时。") from error
    except sdk.CLINotFoundError as error:
        raise ModelNotAvailableError(
            "Claude Agent SDK 找不到原生 Claude Code 二进制。"
            "源码目录未附带该程序；请安装带二进制的 wheel 或原生 claude.exe。"
        ) from error
    except sdk.ResultError as error:
        raise _sdk_failure(error, assistant_error) from error
    except (sdk.CLIConnectionError, sdk.ProcessError) as error:
        if assistant_error in {"billing_error", "rate_limit"}:
            raise _sdk_failure(error, assistant_error) from error
        # Never expose raw CLI stderr: it can contain credentials or prompt data.
        raise ModelNotAvailableError(
            "Agent SDK 的本地 Claude Code 进程启动失败或提前退出。"
            "请使用 scripts/start_claude_sdk_server.ps1 启动后端并检查 SDK 安装。"
        ) from error
    except sdk.CLIJSONDecodeError as error:
        raise LLMResponseError("Claude Agent SDK 返回了无法解析的响应。") from error
    if final is None:
        if assistant_error in {"billing_error", "rate_limit"}:
            raise _sdk_failure(None, assistant_error)
        raise LLMResponseError("Claude Agent SDK 没有返回最终结果。")
    if final.is_error or final.subtype != "success" or assistant_error:
        raise _sdk_failure(final, assistant_error)
    if getattr(final, "stop_reason", None) == "max_tokens" or assistant_stop_reason == "max_tokens":
        raise LLMUsageLimitError("MODEL_OUTPUT_LIMIT")
    return final.result or ""


async def check_sdk_connection(connection: ModelConnection) -> None:
    """Probe the real SDK/MCP round trip using only a random, synthetic token."""
    from .agent_runtime import sdk_agent_connection

    connection = sdk_agent_connection(connection)
    if not connection.api_key:
        raise ModelNotAvailableError("SDK Agent 连接测试需要 API Key；不会借用其他登录身份。")
    sdk = _sdk()
    token = uuid.uuid4().hex
    received = False

    async def probe(args: dict[str, Any]) -> dict[str, Any]:
        nonlocal received
        if args.get("token") != token:
            return {"isError": True, "content": [{"type": "text", "text": "Token mismatch"}]}
        received = True
        return {"content": [{"type": "text", "text": "OK"}]}

    tool = sdk.tool("connection_probe", "Verify this connection with the supplied token.", {
        "type": "object", "properties": {"token": {"type": "string"}},
        "required": ["token"], "additionalProperties": False,
    })(probe)
    server = sdk.create_sdk_mcp_server(name="excel_bro", version="1.0.0", tools=[tool])
    options = _options(sdk, connection,
        "Call connection_probe exactly once with the user's token, then reply OK.",
        max_turns=capability_int("modelConnectionTest", "maxTurns"), server=server,
        allowed=["mcp__excel_bro__connection_probe"])
    await _result(sdk, f"Verify connection using token: {token}", options,
                  capability_float("modelConnectionTest", "timeoutSeconds"))
    if not received:
        raise LLMResponseError("模型返回了文本，但未完成 SDK 工具调用；连接尚不能用于表格 Agent。")


async def check_sdk_message_understanding(
    request: MessageUnderstandingRequest, connection: ModelConnection
) -> MessageUnderstandingResponse:
    from .intent import MESSAGE_UNDERSTANDING_PROMPT, message_understanding_payload, parse_message_understanding

    sdk = _sdk()
    if request.imageCount and not connection.supports_vision:
        raise ModelNotAvailableError("当前模型不支持图片，请选择支持视觉输入的模型。")
    prompt = json.dumps(message_understanding_payload(request), ensure_ascii=False)
    options = _options(sdk, connection, MESSAGE_UNDERSTANDING_PROMPT, max_turns=1)
    timeout = min(capability_float("llm", "timeoutSeconds"), 60.0)
    for attempt in range(2):
        content = await _result(sdk, prompt, options, timeout)
        try:
            return parse_message_understanding(content, request)
        except (ValueError, TypeError, KeyError):
            if attempt == 0:
                prompt += "\n上次输出未通过校验；只返回规定的完整 JSON。"
    raise LLMResponseError("未能理解这条消息；尚未读取工作表结构，请重试。")


async def check_sdk_intent(
    request: IntentCheckRequest, connection: ModelConnection
) -> IntentCheckResponse:
    """Current prompt, image count and authorized table structure only."""
    from .intent import (
        INTENT_SYSTEM_PROMPT, MAX_CLARIFICATION_ROUNDS,
        _parse_model_response, _read_only_routing_fallback,
        is_explicit_read_only_query,
    )

    sdk = _sdk()
    if request.imageCount and not connection.supports_vision:
        raise ModelNotAvailableError("当前模型不支持图片，请选择支持视觉输入的模型。")
    read_only = is_explicit_read_only_query(request.prompt)
    scope = request.scope
    prompt = json.dumps({
        "prompt": request.prompt,
        "imageCount": request.imageCount,
        "intentConfirmed": request.intentConfirmed,
        "clarificationRound": request.clarificationRound,
        "maxClarificationRounds": MAX_CLARIFICATION_ROUNDS,
        "readOnlyRequired": read_only,
        "taskContext": task_context_payload(request),
        # Preserve follow-up semantics without sending historical result rows
        # or free-form host error messages into the intent model.
        "priorIntent": request.priorIntent.model_dump(mode="json") if request.priorIntent else None,
        "priorResult": ({"title": request.priorResult.title,
                         "headers": request.priorResult.headers,
                         "sourceSheets": request.priorResult.sourceSheets}
                        if request.priorResult else None),
        "toolFailure": ({"retryable": request.toolFailure.retryable,
                         "availableFields": request.toolFailure.availableFields,
                         "request": request.toolFailure.request.model_dump(mode="json")}
                        if request.toolFailure else None),
        "scope": {
            "workbookName": scope.workbookName,
            "sourceMode": scope.sourceMode,
            "selectionMode": scope.selectionMode,
            "activeWorksheet": scope.activeWorksheet,
            "selectedRange": scope.selectedRange,
            "worksheetNames": scope.worksheetNames,
            "sheets": [{
                "name": sheet.name,
                "usedRange": sheet.usedRange,
                "rowCount": sheet.rowCount,
                "columnCount": sheet.columnCount,
                "headers": sheet.headers,
            } for sheet in scope.sheets],
        },
    }, ensure_ascii=False)
    options = _options(sdk, connection, INTENT_SYSTEM_PROMPT + TASK_CONTEXT_PROMPT, max_turns=1)
    timeout = min(capability_float("llm", "timeoutSeconds"), 60.0)
    for attempt in range(2):
        content = await _result(sdk, prompt, options, timeout)
        try:
            parsed = _parse_model_response(content, request)
            if parsed.kind == "clarification" and request.clarificationRound >= MAX_CLARIFICATION_ROUNDS:
                raise ValueError("澄清轮次已满")
            if read_only and parsed.kind == "proceed":
                raise ValueError("只读查询不能进入写入规划")
            return parsed
        except (ValueError, TypeError, KeyError):
            if attempt == 0:
                prompt += "\n上次输出未通过结构校验；只返回符合格式的 JSON。"
    if read_only:
        return _read_only_routing_fallback(request)
    raise LLMResponseError("Claude Agent SDK 未能生成有效的意图判断。")


# SDK 后端明确批准下发的非终结工具白名单。登记与投影统一在
# excel_agent.NON_TERMINAL_TOOL_PROJECTIONS / agent_tools 完成；这里只决定
# "哪些非终结结果允许经出站闸下发"，其余 dict 结果一律维持"非终结工具结果
# 不可用"。
_APPROVED_NON_TERMINAL_TOOLS = frozenset(
    {"run_python_analysis", "describe_authorized_tables", "query_table", "prepare_python_data"}
)


async def run_sdk_plan(
    request: PlanRequest, connection: ModelConnection,
    *, on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    on_checkpoint: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    resume_context: dict[str, Any] | None = None,
) -> AssistantResponse:
    """Header-only structure; data values stay on the trusted local side.

    Local ``query_table`` calls do not read the workbook in Python: the SDK
    handler hands the request to the volatile task bridge, the host adapter
    (Office.js / WPS) executes the existing deterministic query locally and
    posts a receipt back. The model only ever sees a typed receipt (opaque
    result ref + allowed structure), never values.

    ``on_checkpoint`` 只在**可信边界**回调（模型开跑前 / 每个工具成功回执后），
    交给调用方盖章（任务身份、授权范围、来源指纹）并由共享契约校验后落盘。
    它**不是** SDK 会话检查点：没有任何模型上下文被保存，续跑是按现场重新分析。

    ``resume_context`` 是上一次中断时留下的受控投影（已确认澄清 + 有归属的结果引用）。
    它可能已过期：引用取不到时必须重新读取，绝不允许把过期引用当可用现场。
    """
    if request.planningReadPolicy != "none" or request.dataResults:
        # 保留原数据行拒绝守卫：任何试图把值矩阵塞进 HTTP 请求的调用都拒绝。
        # 新流程通过本机工具桥 / 结果引用进入，不在这里放开任意 dataResults。
        raise ModelNotAvailableError(
            "Claude Agent SDK 当前仅支持不上传数据行的表结构规划。"
            "需要数据行的请求请使用本机工具桥和有归属的结果引用；不会切换到旧 Agent。"
        )
    if request.images and not connection.supports_vision:
        raise ModelNotAvailableError("当前模型不支持图片，请选择支持视觉输入的模型。")
    from .excel_agent import (
        AGENT_SYSTEM_PROMPT, _execute_tool, _non_terminal_tool_content, _parse_content,
    )
    from .agent_tools import (
        DESCRIBE_TABLES_TOOL,
        PREPARE_PYTHON_DATA_TOOL,
        QUERY_TABLE_TOOL,
        SUBMIT_LOCAL_RESULT_TOOL,
        SUBMIT_RESULT_PLAN_TOOL,
        authorized_structure_payload,
        host_pruned_terminal_definitions,
        local_query_tool_definitions,
        local_terminal_tool_definitions,
        prepare_python_data_tool_definition,
    )
    from .python_analysis import (
        python_analysis_tool_definition, python_error_response_content,
        python_execution_enabled,
    )
    from .task_bridge import BridgeBinding, bridge_step_id, task_tool_bridge

    sdk = _sdk()
    terminal: AssistantResponse | None = None
    unsupported_query_message: str | None = None
    task_id = request.turnId or f"sdk-{uuid.uuid4().hex[:12]}"
    binding = BridgeBinding.from_request(request, task_id=task_id)

    # -- 受控规划现场 ------------------------------------------------------
    # 只累计"可信进展"：工具名、有归属的结果引用、已确认澄清。**不含任何值**。
    # 每次回调都把完整快照交给调用方盖章（调用方才知道 taskId/authorization），
    # 因此这里不需要自增序号，也不会出现"部分更新"。
    resume_projection = _resume_projection(resume_context)
    completed_tools: list[str] = list((resume_projection or {}).get("completedTools", []))
    result_refs: list[dict[str, Any]] = [
        {"resultRef": ref, "sourceFingerprint": None}
        for ref in (resume_projection or {}).get("previousResultRefs", [])
    ]
    failed_tool: dict[str, str] | None = None
    confirmations = _context_confirmations(request)

    def _note_step(name: str, payload: dict[str, Any] | None = None) -> None:
        nonlocal failed_tool
        if isinstance(payload, dict) and payload.get("status") == "error":
            failed_tool = {
                "tool": name,
                "errorCode": str(payload.get("errorCode") or "tool_error")[:80],
            }
            return
        failed_tool = None
        if name not in completed_tools:
            completed_tools.append(name)
        ref = payload.get("resultRef") if isinstance(payload, dict) else None
        if (
            isinstance(ref, str)
            and ref.startswith("res_")
            and len(ref) == 36
            and all(char in "0123456789abcdef" for char in ref[4:])
            and all(item["resultRef"] != ref for item in result_refs)
        ):
            result_refs.append({"resultRef": ref, "sourceFingerprint": None})

    async def _save_checkpoint(stage: str) -> None:
        if on_checkpoint is None:
            return
        try:
            await on_checkpoint({
                "stage": stage,
                "confirmations": [dict(item) for item in confirmations],
                "completedTools": list(completed_tools),
                "resultRefs": [dict(item) for item in result_refs],
                "failedTool": dict(failed_tool) if failed_tool else None,
            })
        except Exception:
            # 落盘失败只降低"能不能继续"的能力，不能让本轮规划崩掉，
            # 更不能因此把任务判成失败。
            logger.warning("无法保存规划现场 stage=%s", stage, exc_info=True)

    async def on_bridge_request(
        operation_id: str, tool: str, arguments: dict[str, Any], fingerprint: str
    ) -> None:
        # 只发一个不透明指针事件；请求参数留在易失桥里由宿主适配器拉取，
        # 不写进会落盘的进度事件/账本。
        if on_event is not None:
            await on_event({
                "phase": "planning",
                # **每次调用一个身份**：同一个工具连续调两次是两个步骤，不是一步。
                # 曾经这里是 `bridge.{tool}`，于是第二次调用的 started 覆盖了第一次
                # 的步骤，界面上"同一张表被读了两次"看起来像只读了一次。
                # 重放同一次调用时 operation_id 不变，因此仍然只更新、不新增。
                "stepId": bridge_step_id(tool, operation_id),
                "status": "started",
                "title": "正在本机执行数据查询",
                "toolName": tool,
                "bridgeTool": {"operationId": operation_id, "tool": tool},
            })

    async def on_bridge_complete(
        operation_id: str, tool: str, receipt: dict[str, Any]
    ) -> None:
        """一次桥调用的**真实结束**：成功、固定错误码、超时、取消都在这里发声。

        没有这条事件时，界面上这一步永远停在"正在执行"，而下一步已经开始了 ——
        前后矛盾的界面只能靠"下一步出现就说明上一步完成"来遮掩，那是猜测。
        """
        if on_event is None:
            return
        status = str(receipt.get("status") or "")
        code = str(receipt.get("errorCode") or "")
        # 取消与超时是"没做成"，不是"失败"：前者是用户停的，后者是现场不允许。
        terminal = (
            "cancelled" if code == "cancelled"
            else "warning" if code in {"timeout", "busy", "budget_exhausted", "source_changed"}
            else "completed" if status == "ok"
            else "failed"
        )
        await on_event({
            "phase": "planning",
            "stepId": bridge_step_id(tool, operation_id),
            "status": terminal,
            "title": "本机数据查询已完成" if terminal == "completed" else "本机数据查询未完成",
            "toolName": tool,
        })

    session = task_tool_bridge.open(
        binding, on_request=on_bridge_request, on_complete=on_bridge_complete
    )
    tools = []
    allowed = []

    async def _emit_submitted(tool_name: str) -> None:
        if on_event is not None:
            await on_event({
                "phase": "planning", "stepId": "planning.submit", "status": "completed",
                "title": "已生成受控结果", "toolName": tool_name,
            })

    definitions = [
        *host_pruned_terminal_definitions(request),
        *local_terminal_tool_definitions(),
        *local_query_tool_definitions(),
        *plan_compiler_definitions(request),
    ]
    # 隔离 Python 分析：总开关默认关闭；开启后追加到注册表尾部。
    # 无执行器时工具仍失败关闭，只返回固定不可用回执。
    if python_execution_enabled():
        definitions.append(prepare_python_data_tool_definition())
        definitions.append(python_analysis_tool_definition())
    try:
        for definition in definitions:
            function = definition["function"]
            name = function["name"]

            async def handle(args: dict[str, Any], tool_name: str = name) -> dict[str, Any]:
                nonlocal terminal, unsupported_query_message
                if terminal is not None:
                    return {"content": [{"type": "text", "text": "本轮结果已提交。"}], "is_error": True}
                if tool_name == DESCRIBE_TABLES_TOOL:
                    payload = {"tables": authorized_structure_payload(request)}
                    return {
                        "content": [{"type": "text", "text": _non_terminal_tool_content(tool_name, payload)}],
                    }
                if tool_name == QUERY_TABLE_TOOL:
                    # Reject unavailable semantics before registering any host read.
                    combine = args.get("combine")
                    if isinstance(combine, dict) and combine.get("mode") in {"join", "deduplicate"}:
                        unsupported_query_message = (
                            "当前工作簿查询暂不支持关联或去重，未执行这次查询。"
                            "如需两表记录合计，请明确要求纵向合并、保留重复记录后统计；"
                            "系统不会自动把关联改为合并。"
                        )
                        return {"content": [{"type": "text", "text": _non_terminal_tool_content(
                            tool_name, {"status": "error", "errorCode": "invalid_arguments"},
                        )}], "is_error": True}
                    try:
                        normalized = QueryTableArguments.model_validate(args).model_dump(mode="json")
                    except (ValueError, TypeError):
                        receipt = {"status": "error", "errorCode": "invalid_arguments"}
                        return {
                            "content": [{"type": "text", "text": _non_terminal_tool_content(tool_name, receipt)}],
                        }
                    unsupported_query_message = None
                    receipt = await session.request(tool_name, normalized)
                    _note_step(tool_name, receipt)
                    if receipt.get("status") == "ok":
                        await _save_checkpoint("after_tool")
                    return {
                        "content": [{"type": "text", "text": _non_terminal_tool_content(tool_name, receipt)}],
                        "is_error": receipt.get("status") == "error",
                    }
                if tool_name == PREPARE_PYTHON_DATA_TOOL:
                    # 数据登记走桥：宿主读授权表 → 值回执 → 服务端登记数据集并
                    # 建立协调器绑定。模型只拿到 dataset:// 句柄，值从不进提示。
                    sources = args.get("sources")
                    if not isinstance(sources, list) or not sources:
                        return {
                            "content": [{"type": "text", "text": _non_terminal_tool_content(
                                tool_name, {"status": "error", "errorCode": "invalid_arguments"},
                            )}],
                            "is_error": True,
                        }
                    normalized = {"sources": sources}
                    receipt = await session.request(tool_name, normalized)
                    _note_step(tool_name, receipt)
                    if receipt.get("status") == "ok":
                        await _save_checkpoint("after_tool")
                    return {
                        "content": [{"type": "text", "text": _non_terminal_tool_content(tool_name, receipt)}],
                        "is_error": receipt.get("status") == "error",
                    }
                if tool_name == SUBMIT_LOCAL_RESULT_TOOL:
                    # 归属校验与结果卡构造收拢在 result_submission：两后端同一份语义。
                    try:
                        item = resolve_result_ref(task_id, args.get("resultRef"))
                    except ResultSubmissionError:
                        return {
                            "content": [{"type": "text", "text": RESULT_REF_INVALID_MESSAGE}],
                            "is_error": True,
                        }
                    # 完成判据（共享层）：声明目标是"生成写入预览"时，只交结果卡
                    # 不算完成。**这里必须 return is_error 而不是设 terminal** ——
                    # 一旦置终局，下面的守卫会锁死后续所有工具，模型再也无法改用
                    # submit_result_plan（这正是缺陷 3 的表现）。
                    try:
                        evaluate_local_result_submission(
                            normalize_planning_goal(args.get("planningGoal")),
                            bound_goal=request.planningGoal,
                        )
                    except LocalResultIncomplete as error:
                        return {
                            "content": [{"type": "text", "text": json.dumps(
                                {"status": "incomplete", "errorCode": error.code,
                                 "message": str(error)[:600]},
                                ensure_ascii=False,
                            )}],
                            "is_error": True,
                        }
                    terminal = build_local_result_answer(item, message=args.get("message") or "")
                    await _emit_submitted(tool_name)
                    return {"content": [{"type": "text", "text": "结果已通过本端校验，请结束本轮。"}]}
                if tool_name == SUBMIT_RESULT_PLAN_TOOL:
                    try:
                        item = resolve_result_ref(task_id, args.get("resultRef"))
                        prepared = build_result_plan_response(
                            item, request,
                            target=args.get("target"),
                            include_header=bool(args.get("includeHeader", True)),
                        )
                    except ResultSubmissionError as error:
                        if error.code == RESULT_REF_INVALID_CODE:
                            return {
                                "content": [{"type": "text", "text": RESULT_REF_INVALID_MESSAGE}],
                                "is_error": True,
                            }
                        # 固定错误码 + 本地校验文案（只含工作表名/区域等非敏感信息）。
                        # **必须置 is_error**：编译被拒（尺寸不符 / 未授权 / 来源已变）
                        # 却回一个"看起来成功"的回执，会让模型据此收尾并谎称完成。
                        return {
                            "content": [{"type": "text", "text": json.dumps(
                                {"status": "error", "errorCode": error.code,
                                 "message": str(error)[:200]},
                                ensure_ascii=False,
                            )}],
                            "is_error": True,
                        }
                    terminal = prepared
                    await _emit_submitted(tool_name)
                    return {"content": [{"type": "text", "text": "计划已通过本端校验并交给宿主执行流程，请结束本轮；提交成功不代表写入已完成。"}]}
                try:
                    if tool_name == "run_python_analysis":
                        # 隔离 Python 是一次同步容器运行（最长 maxWallTimeSeconds）。
                        # 直接同步调用会把本轮 SDK 的流式事件、工具桥回执与取消
                        # 一起冻住 —— 必须走 offload 版本（取消会真实传导到容器）。
                        from .python_coordinator import execute_python_analysis_async

                        result = await execute_python_analysis_async(request, args, on_event=on_event)
                    elif tool_name in PLAN_COMPILERS:
                        result = compile_plan_tool(request, tool_name, args)
                    else:
                        result = _execute_tool(request, tool_name, args)
                except (ValueError, TypeError, KeyError) as error:
                    # 缺陷 3：run_python_analysis 的执行路径异常统一映射为固定安全
                    # 回执，绝不把 str(error)（可能携带参数/代码字面量）回显给模型。
                    if tool_name == "run_python_analysis":
                        return {
                            "content": [{"type": "text", "text": python_error_response_content()}],
                            "is_error": True,
                        }
                    return {
                        "content": [{"type": "text", "text": f"规则或计划未通过校验：{str(error)[:300]}"}],
                        "is_error": True,
                    }
                if isinstance(result, dict):
                    if tool_name in _APPROVED_NON_TERMINAL_TOOLS:
                        # Python's receipt uses resultHandle (the bridge uses resultRef).
                        # Save it before returning to the model so another pause does
                        # not discard an analysis whose result is already on screen.
                        if tool_name == "run_python_analysis" and result.get("status") == "completed":
                            _note_step(tool_name, {"status": "ok", "resultRef": result.get("resultHandle")})
                            await _save_checkpoint("after_tool")
                        return {"content": [{"type": "text", "text": _non_terminal_tool_content(tool_name, result)}]}
                    return {"content": [{"type": "text", "text": "非终结工具结果不可用。"}], "is_error": True}
                terminal = result
                await _emit_submitted(tool_name)
                return {"content": [{"type": "text", "text": "结果已通过本端校验，请结束本轮。"}]}

            tools.append(sdk.tool(name, function["description"], function["parameters"])(handle))
            allowed.append(f"mcp__excel_bro__{name}")
        server = sdk.create_sdk_mcp_server(name="excel_bro", version="1.0.0", tools=tools)
        # 续跑提示：**这是"按当前现场重新分析"，不是恢复模型会话。** 只带入
        # 已确认的澄清（避免重复追问）与有归属的结果引用（可能已过期）。
        resume_projection = _resume_projection(resume_context)
        # Deliberately construct a fresh header-only payload. Never serialize the
        # request's snapshots, lastResult or dataResults into the SDK prompt.
        prompt = json.dumps({
            "prompt": request.prompt,
            "planningGoal": request.planningGoal,
            "taskContext": task_context_payload(request),
            "workbook": {
                "name": request.workbook.name,
                "activeWorksheet": request.workbook.activeWorksheet,
                "selectedRange": request.workbook.selectedRange,
                "worksheets": [{
                    "name": sheet.name,
                    "usedRange": sheet.usedRange,
                    "rowCount": sheet.rowCount,
                    "columnCount": sheet.columnCount,
                    "headers": sheet.headers,
                } for sheet in request.workbook.worksheets],
            },
            # 已登记结果的引用：模型只知道句柄与结构，拿不到值。
            "registeredResultRefs": [ref for ref in request.localResultRefs],
            "planningReadPolicy": "none",
            **({"resumeContext": resume_projection} if resume_projection else {}),
        }, ensure_ascii=False)
        if on_event:
            await on_event({
                "phase": "planning", "stepId": "planning.model", "status": "started",
                "title": "正在生成操作方案",
            })
        options = _options(
            sdk, connection,
            AGENT_SYSTEM_PROMPT + TASK_CONTEXT_PROMPT + AGENT_DECISION_PROMPT
            + "\n本轮只有表结构，没有源值；只能调用 excel_bro 白名单工具。"
            "需要数据计算时调用 mcp__excel_bro__query_table：本地会在真实宿主上"
            "执行确定性查询，回执只给出结果引用与结构，**不含任何数值**。"
            "回执里没有数据值时，不要声称你比较过数值或识别出异常。"
            "同一任务内可以多次调用工具，可按回执中的固定错误码修正参数后重试。"
            "只读任务用 mcp__excel_bro__submit_local_result 提交结果引用让本地"
            "结果卡展示；需要写入时用 mcp__excel_bro__submit_result_plan 提交"
            "结果引用与明确目标，本地会编译为白名单计划；提交成功不等于写入完成。"
            "写入任务由宿主复核授权、来源与恢复准备后执行一次，再回读供审阅，不等待执行前确认。确认变更只接受结果并清除标记，不重复执行；放弃变更恢复原状。仅请求查询或预览时不得擅自写入。"
            "判断本轮目标：若用户要的是「把结果写进某张表 / 生成写入预览」，"
            "展示结果卡只是中间成果，必须再用 submit_result_plan 产出白名单写入计划"
            "才算完成；若目标区域或范围不明确，先向用户澄清；"
            "若确实做不到（工具不支持、编译失败、预算耗尽），如实说明"
            "「分析完成，写入计划未完成」，不要声称任务已成功。"
            "\n终局必须提交受控结果，优先实际调用以下一个工具："
            + "、".join(allowed)
            + "。不能只在文字中描述工具调用或声称已经生成计划；"
            "信息不足时必须调用 mcp__excel_bro__submit_clarification，"
            "将每个缺少的信息放进 fields 供用户分别填写；不要用 submit_answer 追问。",
            max_turns=capability_int("agent", "maxTurns"),
            server=server, allowed=allowed,
        )
        # 模型开跑前先落一次 before_model：这是"已经接住任务、尚未产生任何
        # 可信进展"的最小现场。中断在这里也能如实告诉用户"没做完，可继续"。
        await _save_checkpoint("before_model")
        try:
            content = await _result(
                sdk, _image_prompt(prompt, request.images) if request.images else prompt, options,
                capability_float("taskRecovery", "planningHardWallTimeSeconds"),
            )
        except LLMError as error:
            # 模型这一轮**确实结束了**（失败）：这一步必须如实收尾，不能挂在
            # "正在生成操作方案"上。父阶段的终态与子工具步骤的终态是两件事。
            if on_event is not None:
                await on_event({
                    "phase": "planning", "stepId": "planning.model", "status": "failed",
                    "title": "模型未完成本轮规划",
                })
            # Exhausting turns after an unsupported tool call should retain
            # that concrete capability diagnosis. Never mask a provider quota,
            # context, output or budget limit with an earlier tool error.
            capability_failure = type(error) is LLMResponseError or (
                isinstance(error, LLMUsageLimitError) and error.code == "MODEL_TURN_LIMIT"
            )
            if unsupported_query_message and terminal is None and capability_failure:
                raise LLMResponseError(unsupported_query_message) from error
            raise
        if on_event is not None:
            # 父阶段（模型规划）到此结束：它覆盖了本轮所有子工具调用。
            # 以前没有任何事件关它，界面上只能靠"后面又出现了步骤"去猜；
            # 那个猜法在"模型最后一个动作是子工具"时会把父阶段永远留在进行中。
            await on_event({
                "phase": "planning", "stepId": "planning.model", "status": "completed",
                "title": "模型已给出本轮结果",
            })
        if terminal is None:
            # The shared planning prompt permits structured JSON for providers
            # that return it instead of a tool call. Preserve that contract, but
            # route it through the same terminal handler and capability checks.
            # Plain prose and unvalidated plans never become an executable result.
            try:
                parsed = _parse_content(content)
                if parsed.kind == "plan":
                    await handle({"plan": parsed.plan.model_dump(mode="json")}, "submit_plan")
                elif parsed.kind == "answer":
                    await handle({
                        "message": parsed.message, "resultContext": parsed.resultContext,
                    }, "submit_answer")
                elif parsed.kind == "clarification":
                    await handle({"clarification": parsed.clarification.model_dump(mode="json")}, "submit_clarification")
            except (ValueError, TypeError, KeyError):
                pass
        if terminal is None:
            raise LLMResponseError(unsupported_query_message or "Claude Agent SDK 未提交有效的受控结果。")
        return terminal
    finally:
        task_tool_bridge.close(task_id)

"""Opt-in, header-only Claude Agent SDK adapter.

This adapter never gives the SDK workbook data rows, a live-workbook handle,
filesystem/command tools, or an execution tool. Its only terminal tools return
an existing EB answer, local computation rule, or validated preview plan.
"""

from __future__ import annotations

import asyncio
import json
import os
import sys
from pathlib import Path
from typing import Any, Awaitable, Callable

from .capabilities import capability_float, capability_int
from .llm.config import DEEPSEEK_ANTHROPIC_URL, ModelConnection, ModelNotAvailableError
from .llm.errors import LLMHTTPStatusError, LLMResponseError, LLMTimeoutError
from .models import AssistantResponse, IntentCheckRequest, IntentCheckResponse, PlanRequest


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
        system_prompt=system_prompt,
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
    )


async def _result(sdk: Any, prompt: str, options: Any, timeout: float) -> str:
    # Uvicorn's Windows reload/multi-worker mode installs a Selector loop,
    # which has no subprocess transport. Fail before invoking the SDK rather
    # than misreporting NotImplementedError as a model-network failure.
    if sys.platform == "win32" and isinstance(
        asyncio.get_running_loop(), asyncio.SelectorEventLoop
    ):
        raise ModelNotAvailableError(
            "Agent SDK 无法在当前 Windows 热重载事件循环中启动子进程。"
            "请停止现有后端，再运行 powershell -NoProfile -ExecutionPolicy Bypass "
            "-File .\\scripts\\start_claude_sdk_server.ps1；不要同时运行 npm run dev:server。"
        )
    final = None
    try:
        async with asyncio.timeout(timeout):
            async for item in sdk.query(prompt=prompt, options=options):
                if isinstance(item, sdk.ResultMessage):
                    final = item
    except TimeoutError as error:
        raise LLMTimeoutError("Claude Agent SDK 处理超时。") from error
    except sdk.CLINotFoundError as error:
        raise ModelNotAvailableError(
            "Claude Agent SDK 找不到原生 Claude Code 二进制。"
            "源码目录未附带该程序；请安装带二进制的 wheel 或原生 claude.exe。"
        ) from error
    except sdk.ResultError as error:
        status = error.api_error_status
        if isinstance(status, int) and status >= 400:
            raise LLMHTTPStatusError(status, "") from error
        raise LLMResponseError("Claude Agent SDK 未能完成本轮请求。") from error
    except (sdk.CLIConnectionError, sdk.ProcessError) as error:
        # Never expose raw CLI stderr: it can contain credentials or prompt data.
        raise ModelNotAvailableError(
            "Agent SDK 的本地 Claude Code 进程启动失败或提前退出。"
            "请使用 scripts/start_claude_sdk_server.ps1 启动后端并检查 SDK 安装。"
        ) from error
    except sdk.CLIJSONDecodeError as error:
        raise LLMResponseError("Claude Agent SDK 返回了无法解析的响应。") from error
    if final is None:
        raise LLMResponseError("Claude Agent SDK 没有返回最终结果。")
    if final.is_error or final.subtype != "success":
        status = final.api_error_status
        if isinstance(status, int) and status >= 400:
            raise LLMHTTPStatusError(status, "")
        raise LLMResponseError("Claude Agent SDK 未能完成本轮请求。")
    return final.result or ""


async def check_sdk_intent(
    request: IntentCheckRequest, connection: ModelConnection
) -> IntentCheckResponse:
    """Only the current prompt and authorized table *structure* leave EB."""
    from .intent import (
        INTENT_SYSTEM_PROMPT, MAX_CLARIFICATION_ROUNDS,
        _parse_model_response, _read_only_routing_fallback,
        is_explicit_read_only_query,
    )

    sdk = _sdk()
    if request.imageCount:
        raise ModelNotAvailableError("Claude Agent SDK 当前未接入图片输入，请选择原模型连接。")
    read_only = is_explicit_read_only_query(request.prompt)
    scope = request.scope
    prompt = json.dumps({
        "prompt": request.prompt,
        "intentConfirmed": request.intentConfirmed,
        "clarificationRound": request.clarificationRound,
        "maxClarificationRounds": MAX_CLARIFICATION_ROUNDS,
        "readOnlyRequired": read_only,
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
    options = _options(sdk, connection, INTENT_SYSTEM_PROMPT, max_turns=1)
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


async def run_sdk_plan(
    request: PlanRequest, connection: ModelConnection,
    *, on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
) -> AssistantResponse:
    """Header-only planning; data-bearing requests fail without fallback."""
    if request.planningReadPolicy != "none" or request.dataResults or request.images:
        raise ModelNotAvailableError(
            "Claude Agent SDK 当前仅支持不上传数据行的表结构规划。"
            "请切换到原模型连接处理需要数据行或图片的请求。"
        )
    from .excel_agent import (
        AGENT_SYSTEM_PROMPT, AGENT_TOOLS_WITH_DATA, _execute_tool, _parse_content,
    )

    sdk = _sdk()
    terminal: AssistantResponse | None = None
    tools = []
    allowed = []
    for definition in AGENT_TOOLS_WITH_DATA:
        function = definition["function"]
        name = function["name"]

        async def handle(args: dict[str, Any], tool_name: str = name) -> dict[str, Any]:
            nonlocal terminal
            if terminal is not None:
                return {"content": [{"type": "text", "text": "本轮结果已提交。"}], "is_error": True}
            try:
                result = _execute_tool(request, tool_name, args)
            except (ValueError, TypeError, KeyError) as error:
                return {
                    "content": [{"type": "text", "text": f"规则或计划未通过校验：{str(error)[:300]}"}],
                    "is_error": True,
                }
            if isinstance(result, dict):
                return {"content": [{"type": "text", "text": "非终结工具结果不可用。"}], "is_error": True}
            terminal = result
            if on_event:
                await on_event({
                    "phase": "planning", "stepId": "planning.submit", "status": "completed",
                    "title": "已生成受控结果", "toolName": tool_name,
                })
            return {"content": [{"type": "text", "text": "结果已通过本端校验，请结束本轮。"}]}

        tools.append(sdk.tool(name, function["description"], function["parameters"])(handle))
        allowed.append(f"mcp__excel_bro__{name}")
    server = sdk.create_sdk_mcp_server(name="excel_bro", version="1.0.0", tools=tools)
    # Deliberately construct a fresh header-only payload. Never serialize the
    # request's snapshots, lastResult or dataResults into the SDK prompt.
    prompt = json.dumps({
        "prompt": request.prompt,
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
        "planningReadPolicy": "none",
    }, ensure_ascii=False)
    if on_event:
        await on_event({
            "phase": "planning", "stepId": "planning.model", "status": "started",
            "title": "正在生成操作方案",
        })
    options = _options(
        sdk, connection,
        AGENT_SYSTEM_PROMPT + "\n本轮只有表结构，没有源值；只能调用 excel_bro 白名单工具。"
        "提交计划只会创建预览，不会执行写入。"
        "\n终局必须提交受控结果，优先实际调用以下一个工具："
        + "、".join(allowed)
        + "。不能只在文字中描述工具调用或声称已经生成计划；"
        "信息不足时必须调用 mcp__excel_bro__submit_clarification，"
        "将每个缺少的信息放进 fields 供用户分别填写；不要用 submit_answer 追问。",
        max_turns=capability_int("agent", "maxTurns"),
        server=server, allowed=allowed,
    )
    content = await _result(
        sdk, prompt, options,
        capability_float("taskRecovery", "planningHardWallTimeSeconds"),
    )
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
        raise LLMResponseError("Claude Agent SDK 未提交有效的受控结果。")
    return terminal

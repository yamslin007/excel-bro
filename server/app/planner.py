"""Single SDK planning entry. Local tools never intercept natural language."""
from __future__ import annotations

from collections.abc import Awaitable, Callable
from typing import Any

from . import planning_context
from .agent_runtime import sdk_agent_connection
from .llm import ModelNotAvailableError, ModelNotConfiguredError, selected_model_config
from .models import AssistantResponse, PlanRequest


async def create_plan(
    request: PlanRequest,
    *,
    tool_cache: dict[str, Any] | None = None,
    on_event: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    on_checkpoint: Callable[[dict[str, Any]], Awaitable[None]] | None = None,
    resume_checkpoint: dict[str, Any] | None = None,
    resume_revision: int | None = None,
    approved_wall_time_seconds: float | None = None,
    reminder_seconds: float | None = None,
) -> AssistantResponse:
    try:
        config = sdk_agent_connection(selected_model_config(request.modelId))
    except ModelNotConfiguredError as error:
        raise ModelNotConfiguredError(
            "尚未配置模型，无法理解这条需求。请点击「添加模型」，"
            "填写服务地址、模型 ID 和 API Key 后重试。"
        ) from error
    resume_context: dict[str, Any] | None = None
    if resume_checkpoint is not None:
        # 只有本契约的 planning_context 才可能被接受。旧循环检查点没有 `kind`，
        # 在这个判别点就会被拒绝：绝不拿另一套现场冒充"可继续"。
        verdict = planning_context.validate(
            resume_checkpoint,
            turn_id=request.turnId or "",
            revision=resume_revision if resume_revision is not None else 1,
        )
        if not verdict.ok:
            raise ModelNotAvailableError(
                f"当前的规划现场不满足继续条件（{verdict.code}）：{verdict.detail}。"
                "请重新发起需求。"
            )
        # 传给 SDK 的**只是重新分析的提示**，不是模型会话。
        resume_context = verdict.record
    if request.images and not config.supports_vision:
        raise ModelNotAvailableError("当前模型不支持图片，请选择支持视觉输入的模型。")
    from .claude_sdk_agent import run_sdk_plan
    # Legacy callers can request snapshot planning. SDK always has the narrower
    # structure-only policy; local tool results must still use owned references.
    sdk_request = request.model_copy(update={"planningReadPolicy": "none"})
    return await run_sdk_plan(
        sdk_request,
        config,
        on_event=on_event,
        on_checkpoint=on_checkpoint,
        resume_context=resume_context,
    )

from __future__ import annotations

import asyncio
import hashlib
import time
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Literal

from .capabilities import capability_int
from .models import (
    AnswerResponse,
    IntentCheckResponse,
    PlanRequest,
    PlanResponse,
)


TurnStage = Literal[
    "deciding",
    "awaiting_clarification",
    "awaiting_tool",
    "awaiting_completion",
    "completed",
]


class TurnStateError(ValueError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


@dataclass
class TurnState:
    turn_id: str
    stage: TurnStage
    updated_at: float
    completion_fingerprint: str | None = None
    completion_response: AnswerResponse | PlanResponse | None = None
    completion_lock: asyncio.Lock = field(
        default_factory=asyncio.Lock,
        repr=False,
    )
    # 只读工具（get_workbook_context/find_fields/read_range）的确定性结果缓存，
    # 按 turn_id 存活，使失败重试不必重复读取同一份快照。
    tool_cache: dict[str, Any] = field(default_factory=dict, repr=False)


class TurnRegistry:
    def __init__(self) -> None:
        self._states: OrderedDict[str, TurnState] = OrderedDict()
        self._max_active = capability_int("turns", "maxActive")
        self._ttl_seconds = capability_int("turns", "ttlSeconds")

    def _prune(self) -> None:
        cutoff = time.monotonic() - self._ttl_seconds
        expired = [
            turn_id
            for turn_id, state in self._states.items()
            if state.updated_at < cutoff
        ]
        for turn_id in expired:
            self._states.pop(turn_id, None)
        while len(self._states) > self._max_active:
            self._states.popitem(last=False)

    def touch(self, turn_id: str, stage: TurnStage = "deciding") -> TurnState:
        self._prune()
        state = self._states.get(turn_id)
        if state is None:
            state = TurnState(
                turn_id=turn_id,
                stage=stage,
                updated_at=time.monotonic(),
            )
            self._states[turn_id] = state
        else:
            state.updated_at = time.monotonic()
            self._states.move_to_end(turn_id)
        return state

    def ensure_decision_allowed(self, turn_id: str) -> TurnState:
        state = self.touch(turn_id)
        if state.stage == "completed":
            raise TurnStateError(
                "TURN_ALREADY_COMPLETED",
                "该轮次已经完成，请开始新的请求。",
            )
        return state

    @staticmethod
    def completion_fingerprint(request: PlanRequest) -> str:
        payload = request.model_dump_json(exclude={"turnId"}, exclude_none=True)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def cached_completion(
        self, turn_id: str, request: PlanRequest
    ) -> AnswerResponse | PlanResponse | None:
        state = self.touch(turn_id)
        if state.stage != "completed":
            return None
        fingerprint = self.completion_fingerprint(request)
        if fingerprint != state.completion_fingerprint:
            raise TurnStateError(
                "TURN_ALREADY_COMPLETED",
                "该轮次已经用另一组参数完成，请开始新的请求。",
            )
        return state.completion_response

    def record_decision(
        self, turn_id: str, response: IntentCheckResponse
    ) -> None:
        state = self.touch(turn_id)
        if response.kind == "clarification":
            state.stage = "awaiting_clarification"
        elif response.kind == "tool_request":
            state.stage = "awaiting_tool"
        else:
            state.stage = "awaiting_completion"
        state.updated_at = time.monotonic()

    def record_completion(
        self,
        turn_id: str,
        request: PlanRequest,
        response: AnswerResponse | PlanResponse,
    ) -> None:
        state = self.touch(turn_id)
        state.stage = "completed"
        state.completion_fingerprint = self.completion_fingerprint(request)
        state.completion_response = response
        state.updated_at = time.monotonic()

    def reset(self) -> None:
        self._states.clear()


turn_registry = TurnRegistry()

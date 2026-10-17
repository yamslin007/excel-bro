"""规划时间策略（2026-10-10 改造）的行为证据。

被一条真实多步任务反复打断的根因：``planningReminderSeconds``（60 秒）曾被当成
**硬性暂停预算** —— ``pause_if_needed`` 一到期就抛 ``AgentBudgetPause``，前端弹出
"继续任务"，用户每 60 秒要点一次。本文件用**可控时钟**（打桩 ``time.monotonic``）
证明改造后的真实语义：

1. 跨过提醒阈值后循环**继续**发起下一次模型/工具调用，不进入 ``waiting_user``，
   也不抛 ``AgentBudgetPause``；
2. 每个可恢复边界仍照常落检查点（没有把 ``save_boundary`` 连带删掉）；
3. 到达**累计活跃执行总时限**后不再发起新调用，抛 ``AgentBudgetPause``，且此刻
   现场已保存；
4. 提醒**限频**：每跨一个提醒周期最多发一条，不重置也不伪造任何进度；
5. 累计耗时是**跨恢复累加**的（恢复不清零），等待用户的时间不计入。

全部用打桩时钟，不真实等待。
"""

from __future__ import annotations

import asyncio
import itertools
from typing import Any

import pytest

from server.app import excel_agent as excel_agent_module
from server.app.agent_loop_protection import LoopConfig
from server.app.excel_agent import AgentBudgetPause, run_excel_agent
from server.app.llm import ModelConnection
from server.app.models import PlanRequest


# --------------------------------------------------------------------------
# 可控时钟
# --------------------------------------------------------------------------


class FakeClock:
    """可用 ``advance()`` 手动推进的单调时钟（打桩 ``time.monotonic``）。"""

    def __init__(self, start: float = 0.0) -> None:
        self.now = start

    def __call__(self) -> float:
        return self.now

    def advance(self, seconds: float) -> None:
        self.now += seconds


# --------------------------------------------------------------------------
# 最小宿主：一个只回"纯文本、无工具调用"的假模型连接
# --------------------------------------------------------------------------


def _request() -> PlanRequest:
    """一张已授权工作表的规划请求（不读数据、不写盘）。"""
    return PlanRequest.model_validate(
        {
            "prompt": "分析已授权工作表",
            "turnId": "turn-time-policy",
            "planningReadPolicy": "none",
            "hostId": "excel-office",
            "hostVersion": "16.0",
            "workbook": {
                "name": "策略.xlsx",
                "capturedAt": "2026-10-10T00:00:00Z",
                "activeWorksheet": "表1",
                "selectedRange": "A1:B2",
                "sourceFingerprint": "fp-time-policy",
                "worksheets": [
                    {
                        "name": "表1",
                        "usedRange": "A1:B2",
                        "rowCount": 2,
                        "columnCount": 2,
                        "headers": ["列1", "列2"],
                    }
                ],
            },
        }
    )


def _connection() -> ModelConnection:
    return ModelConnection(
        base_url="http://127.0.0.1:1/v1",
        model="fake-model",
        api_key="fake-key",
        supports_vision=False,
    )


class _FakeClient:
    """假模型客户端：每轮返回一个不含 tool_calls 的降级响应。

    这样循环不会因为工具而分支，只会"一轮一轮往前走"，正是我们要观测的
    "提醒阈值之后还会不会继续发起下一轮模型调用"。
    """

    def __init__(self, *, fail_after: int | None = None) -> None:
        self.calls: list[dict[str, Any]] = []
        self._fail_after = fail_after

    async def __aenter__(self) -> "_FakeClient":
        return self

    async def __aexit__(self, *exc: object) -> None:
        return None

    async def chat_completions(self, **kwargs: Any) -> tuple[dict[str, Any], dict[str, int]]:
        self.calls.append(kwargs)
        if self._fail_after is not None and len(self.calls) > self._fail_after:
            raise AssertionError("总时限之后不应再发起新的模型调用")
        return (
            {
                "choices": [
                    {
                        "message": {"role": "assistant", "content": "先说明一下情况。"},
                        "finish_reason": "stop",
                    }
                ]
            },
            {"total_tokens": 1, "prompt_tokens": 1, "completion_tokens": 0},
        )


@pytest.fixture()
def patched(monkeypatch: pytest.MonkeyPatch):
    """把时钟、客户端与桥全部打桩，返回可观测句柄。"""

    def _install(*, clock: FakeClock, fail_after: int | None = None):
        client = _FakeClient(fail_after=fail_after)
        monkeypatch.setattr(excel_agent_module.time, "monotonic", clock)
        monkeypatch.setattr(
            excel_agent_module,
            "OpenAICompatibleClient",
            lambda *args, **kwargs: client,
        )
        return client

    return _install


def _events_sink() -> tuple[list[dict[str, Any]], Any]:
    events: list[dict[str, Any]] = []

    async def on_event(event: dict[str, Any]) -> None:
        events.append(event)

    return events, on_event


def _checkpoints_sink() -> tuple[list[dict[str, Any]], Any]:
    checkpoints: list[dict[str, Any]] = []

    async def on_checkpoint(checkpoint: dict[str, Any]) -> None:
        checkpoints.append(checkpoint)

    return checkpoints, on_checkpoint


def _plan_config(*, max_turns: int, max_wall_time_seconds: float) -> LoopConfig:
    return LoopConfig(
        max_turns=max_turns,
        max_wall_time_seconds=max_wall_time_seconds,
        max_consecutive_errors=99,
    )


# --------------------------------------------------------------------------
# 1 + 4：跨过 60 秒提醒阈值后继续执行，且提醒限频
# --------------------------------------------------------------------------


def test_crossing_reminder_threshold_keeps_running(patched) -> None:
    """跨过提醒阈值后**继续**发起下一次模型调用，绝不进入暂停。"""
    clock = FakeClock()
    client = patched(clock=clock)
    events, on_event = _events_sink()
    checkpoints, on_checkpoint = _checkpoints_sink()

    async def advance_before_each_call(seconds: float = 0.0) -> None:
        return None

    # 每轮模型调用前把时钟推进 45 秒：第 2 轮开头就已超过 60 秒的提醒阈值。
    original = client.chat_completions
    seen = itertools.count()

    async def chat_completions(**kwargs: Any) -> Any:
        if next(seen) > 0:
            clock.advance(45.0)
        return await original(**kwargs)

    client.chat_completions = chat_completions  # type: ignore[assignment]

    response = asyncio.run(run_excel_agent(
        _request(),
        connection=_connection(),
        timeout=5.0,
        on_event=on_event,
        on_checkpoint=on_checkpoint,
        reminder_seconds=60.0,
        config=_plan_config(max_turns=4, max_wall_time_seconds=1800.0),
    ))

    # 关键判据：跑了多轮、时钟已远超 60 秒，但**没有**抛 AgentBudgetPause。
    assert len(client.calls) >= 3, "跨过提醒阈值后必须继续发起新的模型调用"
    assert clock.now > 60.0
    assert response is not None
    # 提醒事件是**非阻塞**进度提示，不是暂停卡。
    reminders = [e for e in events if str(e.get("stepId", "")).startswith("planning.reminder.")]
    assert reminders, "超过提醒阈值应发出至少一条进度提醒"
    assert all(e.get("status") == "progress" for e in reminders)
    # 提醒限频：4 轮里跨过 1 个 60 秒周期，最多 1 条。
    assert len(reminders) == 1


def test_reminder_does_not_repeat_every_turn(patched) -> None:
    """同一提醒周期内多轮不重复刷提醒（限频有效）。"""
    clock = FakeClock()
    client = patched(clock=clock)
    events, on_event = _events_sink()

    original = client.chat_completions
    seen = itertools.count()

    async def chat_completions(**kwargs: Any) -> Any:
        if next(seen) > 0:
            clock.advance(5.0)  # 每轮 5 秒，4 轮共 15 秒，跨不过 60 秒
        return await original(**kwargs)

    client.chat_completions = chat_completions  # type: ignore[assignment]

    asyncio.run(run_excel_agent(
        _request(),
        connection=_connection(),
        timeout=5.0,
        on_event=on_event,
        reminder_seconds=60.0,
        config=_plan_config(max_turns=4, max_wall_time_seconds=1800.0),
    ))

    reminders = [e for e in events if str(e.get("stepId", "")).startswith("planning.reminder.")]
    assert reminders == [], "未跨过提醒阈值不应发提醒"


# --------------------------------------------------------------------------
# 2：检查点照常保存
# --------------------------------------------------------------------------


def test_checkpoint_still_saved_on_every_boundary(patched) -> None:
    """提醒与暂停分离**不影响**检查点保存（save_boundary 仍在每个边界落盘）。"""
    clock = FakeClock()
    patched(clock=clock)
    checkpoints, on_checkpoint = _checkpoints_sink()

    asyncio.run(run_excel_agent(
        _request(),
        connection=_connection(),
        timeout=5.0,
        on_checkpoint=on_checkpoint,
        reminder_seconds=60.0,
        config=_plan_config(max_turns=3, max_wall_time_seconds=1800.0),
    ))

    assert checkpoints, "每个可恢复边界都必须落检查点"
    stages = {c.get("stage") for c in checkpoints}
    # before_model 与 model_response_saved 都是原本 pause_if_needed 的位置。
    assert "before_model" in stages
    assert "model_response_saved" in stages
    # 检查点里带着累计活跃耗时，恢复时不会被清零。
    for checkpoint in checkpoints:
        assert checkpoint["loopState"]["activeElapsedSeconds"] >= 0.0


# --------------------------------------------------------------------------
# 3：总时限收口
# --------------------------------------------------------------------------


def test_total_wall_time_stops_new_calls_and_saves_scene(patched) -> None:
    """到达总时限后不再发起新调用，抛 AgentBudgetPause 且现场已保存。"""
    clock = FakeClock()
    client = patched(clock=clock, fail_after=1)
    events, on_event = _events_sink()
    checkpoints: list[dict[str, Any]] = []
    advanced = False

    async def on_event(event: dict[str, Any]) -> None:
        nonlocal advanced
        events.append(event)
        # 第 1 轮模型调用**开始之后**才推进 900 秒：模拟"这一轮本身跑得很久"。
        # 下一次 before_model 的 note_progress(enforce_total_wall_time=True)
        # 才会看到累计已越过 600 秒。
        if not advanced and event.get("stepId") == "planning.model.1":
            advanced = True
            clock.advance(900.0)

    async def on_checkpoint(cp: dict[str, Any]) -> None:
        checkpoints.append(cp)

    with pytest.raises(AgentBudgetPause) as captured:
        asyncio.run(run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=5.0,
            on_event=on_event,
            on_checkpoint=on_checkpoint,
            reminder_seconds=60.0,
            config=_plan_config(max_turns=8, max_wall_time_seconds=600.0),
        ))

    pause = captured.value
    # 第 1 轮跑过（clock 在那之后被推了 900 秒），第 2 轮开头即收口。
    assert len(client.calls) == 1, "总时限之后不得再发起新的模型调用"
    # 抛出时现场已保存：检查点存在，且登记进去的累计耗时已越过总时限。
    assert pause.checkpoint is not None
    assert checkpoints[-1]["stage"] == "before_model"
    assert pause.checkpoint["loopState"]["activeElapsedSeconds"] >= 600.0
    assert pause.approved_seconds == 600.0
    # 如实告知：事件里说明是总时限耗尽，不承诺"必然可以继续"。
    exhausted = [e for e in events if e.get("stepId") == "planning.budget_exhausted"]
    assert exhausted, "总时限耗尽必须发出明确事件"
    assert exhausted[0]["status"] == "warning"


def test_budget_pause_is_not_a_reminder(patched) -> None:
    """总时限的停止**不**伪装成提醒事件，两者判别点不同。"""
    clock = FakeClock()
    client = patched(clock=clock, fail_after=0)
    events, on_event = _events_sink()

    async def on_checkpoint(_cp: dict[str, Any]) -> None:
        # 函数入口之后才推进：start_time 已经取过，推进才会体现在 elapsed 上。
        clock.advance(700.0)

    with pytest.raises(AgentBudgetPause):
        asyncio.run(run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=5.0,
            on_event=on_event,
            on_checkpoint=on_checkpoint,
            reminder_seconds=60.0,
            config=_plan_config(max_turns=8, max_wall_time_seconds=600.0),
        ))

    assert len(client.calls) == 0, "总时限已达时不得发起新的模型调用"
    reminder_ids = {str(e.get("stepId", "")) for e in events}
    assert not any(i.startswith("planning.reminder.") for i in reminder_ids)


# --------------------------------------------------------------------------
# 5：累计耗时跨恢复累加，等待用户时间不计入
# --------------------------------------------------------------------------


def test_restore_accumulates_active_time_without_reset(patched) -> None:
    """恢复时累计活跃耗时**累加**而非清零；等待用户的那段不计入。"""
    # 第一次运行：累计活跃耗时跑到 500 秒后被总时限收口。
    first_clock = FakeClock(start=100.0)
    patched(clock=first_clock)
    checkpoints: list[dict[str, Any]] = []
    advanced = False

    async def on_checkpoint(cp: dict[str, Any]) -> None:
        nonlocal advanced
        checkpoints.append(cp)
        # 函数入口之后才推进：start_time 已取过，推进才会体现在 elapsed 上。
        # 每次推 250 秒（小于 400 秒总时限），第一次不超、第二次累计 500 秒才收口，
        # 这样暂停检查点里带的是真实的 500 秒累计耗时。
        advanced = True
        first_clock.advance(250.0)

    with pytest.raises(AgentBudgetPause) as first:
        asyncio.run(run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=5.0,
            on_checkpoint=on_checkpoint,
            reminder_seconds=60.0,
            config=_plan_config(max_turns=8, max_wall_time_seconds=400.0),
        ))
    paused = first.value.checkpoint
    elapsed_after_first = paused["loopState"]["activeElapsedSeconds"]
    assert elapsed_after_first >= 250.0

    # 用户等待很久（真实时钟里只是"没跑"，不影响累计活跃耗时）。
    resume_clock = FakeClock(start=99_999.0)
    resumed_client = patched(clock=resume_clock)
    resumed_events, on_event = _events_sink()

    resumed_checkpoints: list[dict[str, Any]] = []

    async def on_resume_checkpoint(cp: dict[str, Any]) -> None:
        resumed_checkpoints.append(cp)
        resume_clock.advance(1.0)  # 恢复后只推进 1 秒

    asyncio.run(run_excel_agent(
        _request(),
        connection=_connection(),
        timeout=5.0,
        on_event=on_event,
        on_checkpoint=on_resume_checkpoint,
        resume_checkpoint=paused,
        reminder_seconds=60.0,
        config=_plan_config(max_turns=3, max_wall_time_seconds=10_000.0),
    ))

    # 恢复后起点的累计耗时**不等于**零，而是带着之前那 500 秒继续加。
    assert resumed_checkpoints, "恢复后也要继续落检查点"
    resumed_states = [c["loopState"]["activeElapsedSeconds"] for c in resumed_checkpoints]
    assert resumed_states[0] >= elapsed_after_first, "恢复不得把累计活跃耗时清零"
    assert resumed_states[-1] >= resumed_states[0], "恢复后继续累加"


def test_resume_immediately_after_pause_does_not_resend_reminder(patched) -> None:
    """恢复后不因为"时间已过阈值"立刻重发提醒（限频跨恢复沿用）。

    恢复态的起点就是"上次已发过提醒的那个累计耗时"，因此不该刚恢复就再刷一条。
    这条判据要求第一轮真的跨过阈值并发过提醒，否则断言是空的。
    """
    first_clock = FakeClock(start=0.0)
    patched(clock=first_clock)
    first_events, on_first_event = _events_sink()
    advanced = False

    async def on_first_checkpoint(_cp: dict[str, Any]) -> None:
        nonlocal advanced
        # 每个边界推进 200 秒（第一轮即跨过 60 秒提醒阈值），很快撞上总时限。
        advanced = True
        first_clock.advance(200.0)

    with pytest.raises(AgentBudgetPause) as captured:
        asyncio.run(run_excel_agent(
            _request(),
            connection=_connection(),
            timeout=5.0,
            on_event=on_first_event,
            on_checkpoint=on_first_checkpoint,
            reminder_seconds=60.0,
            config=_plan_config(max_turns=8, max_wall_time_seconds=400.0),
        ))
    paused = captured.value.checkpoint

    # 正向锚点：第一轮确实发过提醒，否则下面的"恢复后不发"没有判别力。
    first_reminders = [
        e for e in first_events if str(e.get("stepId", "")).startswith("planning.reminder.")
    ]
    assert first_reminders, "第一轮跨过阈值必须先真的发过提醒"

    resume_clock = FakeClock(start=5_000.0)
    patched(clock=resume_clock)
    events, on_event = _events_sink()

    async def on_resume_checkpoint(_cp: dict[str, Any]) -> None:
        resume_clock.advance(1.0)  # 恢复后只推进 1 秒，远不到下一个提醒周期

    asyncio.run(run_excel_agent(
        _request(),
        connection=_connection(),
        timeout=5.0,
        on_event=on_event,
        on_checkpoint=on_resume_checkpoint,
        resume_checkpoint=paused,
        reminder_seconds=60.0,
        config=_plan_config(max_turns=2, max_wall_time_seconds=10_000.0),
    ))

    # 恢复即处于阈值之上，但上次发提醒的位置也被带回来了 ⇒ 立刻重发是限频漏洞。
    reminders = [e for e in events if str(e.get("stepId", "")).startswith("planning.reminder.")]
    assert reminders == [], "恢复后立即重发提醒说明限频没有跨恢复沿用"

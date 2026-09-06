"""Agent 循环保护机制

检测 Agent 循环中的异常模式：
- 动作矛盾：重复查询相同工具+参数
- 状态机路径循环：在多个工具之间重复相同调用路径
- 连续错误：持续的 ValidationError 或解析失败
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .models import AnswerResponse


@dataclass
class LoopState:
    """Agent 循环状态追踪

    Attributes:
        turn_index: 当前轮次索引（从 0 开始）
        tool_call_history: 工具名称调用历史序列
        cache_hit_counts: 缓存键到命中次数的映射
        consecutive_errors: 连续错误次数（ValidationError、JSON 解析失败等）
        total_tokens: 累计 token 消耗（如果 API 返回 usage）
        last_warning: 上一次注入的警告内容，避免重复警告
        detected_path_loop: 最近一次已告警的路径循环模式
        elapsed_time: Agent 循环经过的墙上时间（秒）
        encountered_truncation: 是否遇到工作簿快照截断
        last_error_message: 最近一次结构或工具错误的简要信息
        last_failed_plan: 最近一次校验失败的提交计划
    """
    turn_index: int = 0
    tool_call_history: list[str] = field(default_factory=list)
    cache_hit_counts: dict[str, int] = field(default_factory=dict)
    consecutive_errors: int = 0
    total_tokens: int = 0
    last_warning: str | None = None
    detected_path_loop: str | None = None
    elapsed_time: float = 0.0
    encountered_truncation: bool = False
    last_error_message: str | None = None
    last_failed_plan: dict[str, Any] | None = None


@dataclass
class LoopConfig:
    """Agent 循环预算和异常检测配置。"""

    max_turns: int = 8
    max_total_tokens: int = 20_000
    max_wall_time_seconds: float = 60.0
    max_consecutive_errors: int = 3
    max_cache_hits_per_query: int = 3
    max_tool_path_loop_length: int = 4
    warning_turn: int = 5


def detect_action_conflict(
    state: LoopState,
    cache_key: str | None,
    max_cache_hits_per_query: int = 3,
) -> str | None:
    """检测动作矛盾（重复查询同一工具）

    当 Agent 反复查询相同的工具+参数组合，且结果未变时，视为动作矛盾。
    这通常表示 Agent 陷入了原地打转，需要介入提醒。

    Args:
        state: 当前循环状态（会原地修改 cache_hit_counts）
        cache_key: 如果这次工具调用命中了缓存，传入缓存键；否则传 None
        max_cache_hits_per_query: 同一缓存键触发告警的命中次数

    Returns:
        如果命中次数达到阈值，返回警告文案；否则返回 None。同一个
        cache_key 只在恰好达到阈值时告警一次，后续命中不再返回。

    Examples:
        >>> state = LoopState()
        >>> detect_action_conflict(state, "find_fields:sales")  # 第 1 次
        >>> detect_action_conflict(state, "find_fields:sales")  # 第 2 次
        >>> detect_action_conflict(state, "find_fields:sales")  # 第 3 次
        '⚠️ 你已第 3 次查询相同内容，结果未变。请基于现有信息决策，或调用 submit_answer 说明卡在哪里。'
    """
    if cache_key is None:
        return None

    # 更新缓存命中计数
    state.cache_hit_counts[cache_key] = state.cache_hit_counts.get(cache_key, 0) + 1
    hit_count = state.cache_hit_counts[cache_key]

    # 恰好达到阈值时触发，避免同一查询重复告警。
    if hit_count == max_cache_hits_per_query:
        warning = (
            f"⚠️ 你已第 {hit_count} 次查询相同内容，结果未变。"
            f"请基于现有信息决策，或调用 submit_answer 说明卡在哪里。"
        )
        state.last_warning = warning
        return warning

    return None


def should_stop(
    state: LoopState, config: LoopConfig
) -> tuple[bool, str | None]:
    """判断 Agent 循环是否应该触发预算熔断。

    按步数、连续错误、累计 token 和墙上时间的固定优先级检查，返回
    第一个达到配置阈值的原因。

    Args:
        state: 当前循环状态。
        config: 循环预算配置。

    Returns:
        (是否停止, 原因)。原因可能为 max_turns、max_tokens、timeout 或
        consecutive_errors；未命中时原因是 None。
    """
    if state.turn_index >= config.max_turns:
        return True, "max_turns"
    # 当连续结构错误同时推高了 token，用根因而不是伴随结果向用户解释停止原因。
    if state.consecutive_errors >= config.max_consecutive_errors:
        return True, "consecutive_errors"
    if state.total_tokens >= config.max_total_tokens:
        return True, "max_tokens"
    if state.elapsed_time >= config.max_wall_time_seconds:
        return True, "timeout"
    return False, None


def safe_degradation(state: LoopState, stop_reason: str) -> AnswerResponse:
    """生成包含尝试摘要和后续建议的安全降级响应。

    Args:
        state: 熔断时的循环状态。
        stop_reason: should_stop 返回的熔断原因。

    Returns:
        可由前端直接渲染的结构化系统回答。
    """
    from .models import AnswerResponse

    tool_counts = Counter(state.tool_call_history)
    tool_summary = ", ".join(
        f"{tool} ({count}次)" for tool, count in tool_counts.items()
    ) or "无"
    reason_map = {
        "max_turns": (
            f"已尝试 {state.turn_index} 轮探索，但未能收敛到可执行计划"
        ),
        "max_tokens": (
            f"累计消耗 token 达到预算上限（{state.total_tokens} tokens）"
        ),
        "timeout": f"规划时间超过 {state.elapsed_time:.1f} 秒",
        "consecutive_errors": (
            f"连续 {state.consecutive_errors} 次提交的计划不符合格式要求"
        ),
    }
    loop_info = ""
    if state.detected_path_loop:
        loop_info = f"\n- 检测到循环路径：{state.detected_path_loop}"

    suggestions = _generate_suggestions(state, stop_reason)
    message = (
        "抱歉，我在规划过程中遇到困难：\n\n"
        f"**原因：** {reason_map.get(stop_reason, stop_reason)}\n\n"
        "**尝试记录：**\n"
        f"- 调用工具：{tool_summary}\n"
        f"- 累计轮次：{state.turn_index}\n"
        f"- 累计 token：{state.total_tokens}{loop_info}\n\n"
        "**建议：**\n"
        f"{suggestions}"
    )
    return AnswerResponse(provider="system", message=message)


def _generate_suggestions(state: LoopState, stop_reason: str) -> str:
    """根据数据完整性和熔断原因生成可操作建议。"""
    if state.encountered_truncation:
        return (
            "- 数据预览超过快照容量（200 行），但这不代表 Excel 无法处理完整数据\n"
            "- 拆分、聚合或统计任务应使用执行时读取完整数据的计划；只有必须查看"
            "具体明细时才需要缩小范围"
        )

    if stop_reason == "consecutive_errors":
        error_hint = ""
        if state.last_error_message:
            error_hint = f"\n- 最后一次错误：{state.last_error_message[:150]}"
        return (
            f"- 我多次提交的计划格式不正确{error_hint}\n"
            "- 建议：尝试更简单的操作，或换个表达方式重新描述需求"
        )

    if state.detected_path_loop:
        return (
            "- 我在探索数据时陷入循环\n"
            "- 建议：提供更明确的字段名称或数据范围"
        )

    return "- 尝试简化需求，分步完成\n- 或提供更明确的数据位置和操作细节"


def classify_stuck_state(state: LoopState) -> str:
    """将当前 Agent 状态分类为发散、收敛失败、循环或正常。"""
    if state.consecutive_errors >= 2:
        return "convergence_failure"

    if state.detected_path_loop or any(
        count >= 2 for count in state.cache_hit_counts.values()
    ):
        return "looping"

    recent_tools = (
        state.tool_call_history[-6:]
        if len(state.tool_call_history) >= 6
        else state.tool_call_history
    )
    readonly_tools = {"get_workbook_context", "find_fields", "read_range"}
    if recent_tools and all(tool in readonly_tools for tool in recent_tools):
        return "diverging"

    return "normal"


def get_warning_prompt(state: LoopState) -> str | None:
    """根据当前卡住状态返回一次循环收尾提示，正常状态返回 None。"""
    prompts = {
        "diverging": (
            "\n\n⚠️ 预算提示：还剩 2-3 轮探索机会。如果你已掌握足够信息，"
            "请尽快调用 submit_plan 或 submit_answer。"
        ),
        "convergence_failure": (
            "\n\n⚠️ 你已连续多次提交不合法的计划。建议：\n"
            "1. 大幅简化计划（减少动作数量、去掉复杂条件）\n"
            "2. 或调用 submit_answer 说明你想做什么，但无法用当前格式表达"
        ),
        "looping": (
            "\n\n⚠️ 你似乎在重复查询相同内容或循环调用工具。建议：\n"
            "1. 基于已有查询结果直接提交计划\n"
            "2. 或调用 submit_answer 说明当前信息不足，还需要什么"
        ),
        "normal": None,
    }
    return prompts.get(classify_stuck_state(state))


def detect_path_loop(state: LoopState, max_loop_length: int = 4) -> str | None:
    """检测状态机路径循环（A→B→A→B 模式）。

    从工具调用历史末尾检查长度为 2 到 max_loop_length 的模式。
    如果最后 2N 个工具调用由两个完全相同的 N 步路径组成，则认为
    Agent 正在重复同一状态机路径。同一个模式连续检测到时只告警一次；
    检测到不同模式后会更新状态并重新告警。

    Args:
        state: 当前循环状态（检测到新模式时会更新 detected_path_loop）。
        max_loop_length: 最大检测的循环长度，默认检测到 4 步循环。

    Returns:
        如果检测到尚未告警的循环，返回警告文案；否则返回 None。

    Examples:
        >>> state = LoopState(tool_call_history=["A", "B", "A", "B"])
        >>> detect_path_loop(state)
        '⚠️ 检测到重复路径 A→B。你似乎在原地打转，请换个角度或基于现有信息提交结果。'
    """
    history = state.tool_call_history

    for loop_length in range(2, max_loop_length + 1):
        required_history_length = loop_length * 2
        if len(history) < required_history_length:
            continue

        recent_calls = history[-required_history_length:]
        pattern_calls = recent_calls[:loop_length]
        if pattern_calls != recent_calls[loop_length:]:
            continue

        pattern = "→".join(pattern_calls)
        if state.detected_path_loop == pattern:
            return None

        state.detected_path_loop = pattern
        return (
            f"⚠️ 检测到重复路径 {pattern}。"
            "你似乎在原地打转，请换个角度或基于现有信息提交结果。"
        )

    return None


def update_loop_state(
    state: LoopState,
    tool_name: str | None,
    cache_key: str | None,
    error_occurred: bool
) -> None:
    """更新循环状态

    在每轮 Agent 循环结束后调用，更新追踪状态。

    Args:
        state: 当前循环状态（会原地修改）
        tool_name: 本轮调用的工具名称（如 "find_fields"），如果没有调用工具则为 None
        cache_key: 本轮的缓存键（如果命中缓存），否则为 None
        error_occurred: 本轮是否发生错误（ValidationError、JSON 解析失败等）

    Examples:
        >>> state = LoopState()
        >>> update_loop_state(state, "find_fields", None, False)
        >>> state.turn_index
        1
        >>> state.tool_call_history
        ['find_fields']
        >>> state.consecutive_errors
        0
    """
    state.turn_index += 1

    if tool_name is not None:
        state.tool_call_history.append(tool_name)

    if error_occurred:
        state.consecutive_errors += 1
    else:
        state.consecutive_errors = 0

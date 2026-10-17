"""测试 Agent 循环保护机制"""

import pytest

from server.app.agent_loop_protection import (
    LoopConfig,
    LoopState,
    classify_stuck_state,
    detect_action_conflict,
    detect_path_loop,
    get_warning_prompt,
    safe_degradation,
    should_stop,
    update_loop_state,
)


class TestLoopState:
    """测试 LoopState 数据结构"""

    def test_default_initialization(self):
        """测试默认初始化"""
        state = LoopState()
        assert state.turn_index == 0
        assert state.tool_call_history == []
        assert state.cache_hit_counts == {}
        assert state.consecutive_errors == 0
        assert state.total_tokens == 0
        assert state.last_warning is None
        assert state.detected_path_loop is None
        assert state.elapsed_time == 0.0
        assert state.encountered_truncation is False
        assert state.last_error_message is None
        assert state.last_failed_plan is None

    def test_custom_initialization(self):
        """测试自定义初始化"""
        state = LoopState(
            turn_index=5,
            tool_call_history=["find_fields", "read_range"],
            cache_hit_counts={"key1": 2},
            consecutive_errors=1,
            total_tokens=1000,
            last_warning="test warning",
            detected_path_loop="find_fields→read_range",
            elapsed_time=12.5,
            encountered_truncation=True,
            last_error_message="invalid plan",
            last_failed_plan={"id": "failed-plan"},
        )
        assert state.turn_index == 5
        assert state.tool_call_history == ["find_fields", "read_range"]
        assert state.cache_hit_counts == {"key1": 2}
        assert state.consecutive_errors == 1
        assert state.total_tokens == 1000
        assert state.last_warning == "test warning"
        assert state.detected_path_loop == "find_fields→read_range"
        assert state.elapsed_time == 12.5
        assert state.encountered_truncation is True
        assert state.last_error_message == "invalid plan"
        assert state.last_failed_plan == {"id": "failed-plan"}


class TestLoopConfig:
    """测试循环配置默认值。"""

    def test_default_budget_and_detection_thresholds(self):
        config = LoopConfig()

        assert config.max_turns == 8
        assert config.max_total_tokens is None
        assert config.max_wall_time_seconds == 60.0
        assert config.max_consecutive_errors == 3
        assert config.max_cache_hits_per_query == 3
        assert config.max_tool_path_loop_length == 4
        assert config.warning_turn == 5

    @pytest.mark.parametrize("value", [0, -1, True, "20000", 3.5])
    def test_invalid_token_budget_is_not_silently_disabled(self, value):
        with pytest.raises(ValueError):
            LoopConfig(max_total_tokens=value)


class TestDetectActionConflict:
    """测试动作矛盾检测"""

    def test_no_cache_key_returns_none(self):
        """没有缓存键时返回 None"""
        state = LoopState()
        result = detect_action_conflict(state, None)
        assert result is None
        assert state.cache_hit_counts == {}

    def test_first_cache_hit_no_warning(self):
        """第一次缓存命中，不返回警告"""
        state = LoopState()
        result = detect_action_conflict(state, "find_fields:sales")
        assert result is None
        assert state.cache_hit_counts["find_fields:sales"] == 1

    def test_second_cache_hit_no_warning(self):
        """第二次缓存命中，仍不返回警告"""
        state = LoopState()
        detect_action_conflict(state, "find_fields:sales")
        result = detect_action_conflict(state, "find_fields:sales")
        assert result is None
        assert state.cache_hit_counts["find_fields:sales"] == 2

    def test_third_cache_hit_returns_warning(self):
        """第三次缓存命中，返回警告"""
        state = LoopState()
        detect_action_conflict(state, "find_fields:sales")
        detect_action_conflict(state, "find_fields:sales")
        result = detect_action_conflict(state, "find_fields:sales")

        assert result is not None
        assert "⚠️" in result
        assert "第 3 次查询相同内容" in result
        assert "结果未变" in result
        assert "submit_answer" in result
        assert state.cache_hit_counts["find_fields:sales"] == 3
        assert state.last_warning == result

    def test_fourth_cache_hit_no_duplicate_warning(self):
        """第四次缓存命中，不重复返回相同警告"""
        state = LoopState()
        detect_action_conflict(state, "find_fields:sales")
        detect_action_conflict(state, "find_fields:sales")
        warning1 = detect_action_conflict(state, "find_fields:sales")  # 第 3 次
        warning2 = detect_action_conflict(state, "find_fields:sales")  # 第 4 次

        assert warning1 is not None
        assert warning2 is None  # 第 4 次不重复警告
        assert state.cache_hit_counts["find_fields:sales"] == 4

    def test_different_cache_keys_tracked_separately(self):
        """不同缓存键独立追踪"""
        state = LoopState()

        # key1 命中 2 次
        detect_action_conflict(state, "find_fields:sales")
        detect_action_conflict(state, "find_fields:sales")

        # key2 命中 3 次
        detect_action_conflict(state, "read_range:A1:B10")
        detect_action_conflict(state, "read_range:A1:B10")
        warning = detect_action_conflict(state, "read_range:A1:B10")

        assert warning is not None
        assert "第 3 次" in warning
        assert state.cache_hit_counts["find_fields:sales"] == 2
        assert state.cache_hit_counts["read_range:A1:B10"] == 3

    def test_warning_count_reflects_actual_hits(self):
        """警告中显示的次数正确反映实际命中次数"""
        state = LoopState()

        # 连续命中 5 次
        for i in range(5):
            result = detect_action_conflict(state, "same_key")
            if i < 2:
                assert result is None  # 前两次无警告
            elif i == 2:
                assert result is not None  # 第 3 次有警告
                assert "第 3 次" in result
            else:
                assert result is None  # 第 4、5 次不再警告

        # 验证计数正确
        assert state.cache_hit_counts["same_key"] == 5


class TestPathLoopDetection:
    """测试状态机路径循环检测。"""

    def test_no_loop_returns_none(self):
        """互不重复的调用路径不告警。"""
        state = LoopState(tool_call_history=["A", "B", "C", "D"])

        assert detect_path_loop(state) is None
        assert state.detected_path_loop is None

    def test_detects_two_step_loop(self):
        """检测 A→B 的两步循环。"""
        state = LoopState(tool_call_history=["A", "B", "A", "B"])

        warning = detect_path_loop(state)

        assert warning == (
            "⚠️ 检测到重复路径 A→B。"
            "你似乎在原地打转，请换个角度或基于现有信息提交结果。"
        )
        assert state.detected_path_loop == "A→B"

    def test_detects_three_step_loop(self):
        """检测 A→B→C 的三步循环。"""
        state = LoopState(
            tool_call_history=["A", "B", "C", "A", "B", "C"]
        )

        warning = detect_path_loop(state)

        assert warning is not None
        assert "重复路径 A→B→C" in warning
        assert state.detected_path_loop == "A→B→C"

    def test_detects_four_step_loop(self):
        """检测 A→B→C→D 的四步循环。"""
        state = LoopState(
            tool_call_history=["A", "B", "C", "D", "A", "B", "C", "D"]
        )

        warning = detect_path_loop(state)

        assert warning is not None
        assert "重复路径 A→B→C→D" in warning
        assert state.detected_path_loop == "A→B→C→D"

    def test_insufficient_history_returns_none(self):
        """不足两轮完整两步路径时不告警。"""
        state = LoopState(tool_call_history=["A", "B", "A"])

        assert detect_path_loop(state) is None
        assert state.detected_path_loop is None

    def test_same_pattern_warns_only_once(self):
        """同一个循环模式连续检测时不重复告警。"""
        state = LoopState(tool_call_history=["A", "B", "A", "B"])

        first_warning = detect_path_loop(state)
        second_warning = detect_path_loop(state)

        assert first_warning is not None
        assert second_warning is None
        assert state.detected_path_loop == "A→B"

    def test_different_patterns_are_tracked_independently(self):
        """切换到不同循环模式时重新告警。"""
        state = LoopState(tool_call_history=["A", "B", "A", "B"])
        first_warning = detect_path_loop(state)

        state.tool_call_history.extend(["C", "D", "C", "D"])
        second_warning = detect_path_loop(state)

        assert first_warning is not None
        assert "重复路径 A→B" in first_warning
        assert second_warning is not None
        assert "重复路径 C→D" in second_warning
        assert state.detected_path_loop == "C→D"

    def test_detects_loop_at_end_of_complex_history(self):
        """只依据历史末尾检测最近的重复路径。"""
        state = LoopState(
            tool_call_history=["X", "Y", "Z", "A", "B", "A", "B"]
        )

        warning = detect_path_loop(state)

        assert warning is not None
        assert "重复路径 A→B" in warning
        assert state.detected_path_loop == "A→B"


class TestBudgetControl:
    """测试预算熔断。"""

    def test_should_stop_max_turns(self):
        state = LoopState(turn_index=8)

        assert should_stop(state, LoopConfig()) == (True, "max_turns")

    def test_should_stop_max_tokens(self):
        state = LoopState(total_tokens=20_000)

        assert should_stop(state, LoopConfig(max_total_tokens=20_000)) == (True, "max_tokens")

    def test_disabled_token_budget_keeps_accounting_without_stopping(self):
        state = LoopState(total_tokens=1_000_000)
        restored = LoopState.from_checkpoint(state.to_checkpoint())
        assert restored.total_tokens == 1_000_000
        assert should_stop(restored, LoopConfig()) == (False, None)

    @pytest.mark.parametrize("state, reason", [
        (LoopState(total_tokens=1_000_000, turn_index=8), "max_turns"),
        (LoopState(total_tokens=1_000_000, consecutive_errors=3), "consecutive_errors"),
    ])
    def test_disabled_tokens_do_not_disable_other_guards(self, state, reason):
        assert should_stop(state, LoopConfig(max_total_tokens=None)) == (True, reason)

    def test_should_stop_does_not_own_wall_time(self):
        """累计活跃耗时的收口不在 ``should_stop``。

        2026-10-10 改造：``config.max_wall_time_seconds`` 是**累计活跃执行总时限**
        （默认 1800 秒）。它必须由 ``run_excel_agent`` 的
        ``note_progress(..., enforce_total_wall_time=True)`` 判定并抛
        ``AgentBudgetPause``（现场已保存、可按真实恢复能力继续）；若留在
        ``should_stop``，会被当作普通熔断走 ``safe_degradation``，出现"一边说现场
        已保存、一边说规划超时"的双重口径。
        """
        over_budget = LoopState(elapsed_time=10_000.0)

        # 与 max_turns / consecutive_errors 等真实熔断不同，超时**不**在这里停。
        assert should_stop(over_budget, LoopConfig()) == (False, None)
        assert should_stop(
            LoopState(elapsed_time=10_000.0, turn_index=8), LoopConfig()
        ) == (True, "max_turns")

    def test_should_stop_consecutive_errors(self):
        state = LoopState(consecutive_errors=3)

        assert should_stop(state, LoopConfig()) == (
            True,
            "consecutive_errors",
        )

    def test_structure_error_is_reported_before_its_token_side_effect(self):
        state = LoopState(consecutive_errors=3, total_tokens=20_000)

        assert should_stop(state, LoopConfig(max_total_tokens=20_000)) == (
            True,
            "consecutive_errors",
        )

    def test_should_not_stop_normal(self):
        state = LoopState(
            turn_index=7,
            total_tokens=19_999,
            elapsed_time=59.9,
            consecutive_errors=2,
        )

        assert should_stop(state, LoopConfig()) == (False, None)


class TestSafeDegradation:
    """测试安全降级消息生成。"""

    def test_degradation_with_truncation(self):
        state = LoopState(encountered_truncation=True)

        response = safe_degradation(state, "max_turns")

        assert response.provider == "system"
        assert "超过快照容量" in response.message
        assert "不代表 Excel 无法处理完整数据" in response.message
        assert "执行时读取完整数据" in response.message
        assert "使用 Excel 内置的筛选功能先过滤数据" not in response.message

    def test_degradation_with_loop(self):
        state = LoopState(
            tool_call_history=["A", "B", "A", "B"],
            detected_path_loop="A→B",
        )

        response = safe_degradation(state, "max_tokens")

        assert "检测到循环路径：A→B" in response.message
        assert "陷入循环" in response.message

    def test_degradation_with_consecutive_errors(self):
        state = LoopState(
            consecutive_errors=3,
            last_error_message="动作字段缺失",
        )

        response = safe_degradation(state, "consecutive_errors")

        assert "格式不正确" in response.message
        assert "动作字段缺失" in response.message


class TestStuckStateClassification:
    """测试卡住状态分类和对应收尾提示。"""

    def test_classify_diverging(self):
        state = LoopState(
            tool_call_history=[
                "get_workbook_context",
                "find_fields",
                "read_range",
                "find_fields",
                "read_range",
                "get_workbook_context",
            ]
        )

        assert classify_stuck_state(state) == "diverging"

    def test_classify_convergence_failure(self):
        state = LoopState(consecutive_errors=2, detected_path_loop="A→B")

        assert classify_stuck_state(state) == "convergence_failure"

    def test_classify_looping(self):
        state = LoopState(detected_path_loop="A→B")

        assert classify_stuck_state(state) == "looping"
        assert "重复查询" in (get_warning_prompt(state) or "")

    def test_classify_normal(self):
        state = LoopState(tool_call_history=["submit_answer"])

        assert classify_stuck_state(state) == "normal"
        assert get_warning_prompt(state) is None


class TestUpdateLoopState:
    """测试循环状态更新"""

    def test_turn_index_increments(self):
        """每次调用都增加轮次"""
        state = LoopState()
        update_loop_state(state, None, None, False)
        assert state.turn_index == 1

        update_loop_state(state, None, None, False)
        assert state.turn_index == 2

    def test_tool_call_history_appends(self):
        """工具名称正确追加到历史"""
        state = LoopState()
        update_loop_state(state, "find_fields", None, False)
        update_loop_state(state, "read_range", None, False)
        update_loop_state(state, "find_fields", None, False)

        assert state.tool_call_history == ["find_fields", "read_range", "find_fields"]

    def test_no_tool_call_skips_history(self):
        """没有工具调用时不追加历史"""
        state = LoopState()
        update_loop_state(state, None, None, False)
        update_loop_state(state, "find_fields", None, False)
        update_loop_state(state, None, None, False)

        assert state.tool_call_history == ["find_fields"]
        assert state.turn_index == 3

    def test_consecutive_errors_increment(self):
        """连续错误计数正确增加"""
        state = LoopState()
        update_loop_state(state, "find_fields", None, True)
        assert state.consecutive_errors == 1

        update_loop_state(state, "read_range", None, True)
        assert state.consecutive_errors == 2

        update_loop_state(state, "find_fields", None, True)
        assert state.consecutive_errors == 3

    def test_consecutive_errors_reset_on_success(self):
        """成功执行时重置错误计数"""
        state = LoopState()
        update_loop_state(state, "find_fields", None, True)
        update_loop_state(state, "read_range", None, True)
        assert state.consecutive_errors == 2

        # 成功一次，计数归零
        update_loop_state(state, "find_fields", None, False)
        assert state.consecutive_errors == 0

        # 再次错误，从 1 开始
        update_loop_state(state, "read_range", None, True)
        assert state.consecutive_errors == 1

    def test_cache_key_does_not_affect_update(self):
        """cache_key 参数不影响状态更新（仅用于检测）"""
        state = LoopState()
        update_loop_state(state, "find_fields", "some_cache_key", False)

        # cache_key 在 update_loop_state 中不使用，不影响状态
        assert state.turn_index == 1
        assert state.tool_call_history == ["find_fields"]
        assert state.consecutive_errors == 0

    def test_complete_workflow(self):
        """完整工作流测试"""
        state = LoopState()

        # 第 1 轮：成功调用 find_fields
        update_loop_state(state, "find_fields", None, False)
        assert state.turn_index == 1
        assert state.tool_call_history == ["find_fields"]
        assert state.consecutive_errors == 0

        # 第 2 轮：调用 read_range，发生错误
        update_loop_state(state, "read_range", None, True)
        assert state.turn_index == 2
        assert state.tool_call_history == ["find_fields", "read_range"]
        assert state.consecutive_errors == 1

        # 第 3 轮：再次错误
        update_loop_state(state, "find_fields", None, True)
        assert state.turn_index == 3
        assert state.tool_call_history == ["find_fields", "read_range", "find_fields"]
        assert state.consecutive_errors == 2

        # 第 4 轮：成功，错误计数归零
        update_loop_state(state, "read_range", None, False)
        assert state.turn_index == 4
        assert state.tool_call_history == ["find_fields", "read_range", "find_fields", "read_range"]
        assert state.consecutive_errors == 0


class TestIntegration:
    """集成测试：检测 + 更新组合使用"""

    def test_typical_loop_with_repeated_queries(self):
        """典型场景：Agent 重复查询同一内容"""
        state = LoopState()

        # 第 1 轮：首次查询
        warning = detect_action_conflict(state, "find_fields:sales")
        assert warning is None
        update_loop_state(state, "find_fields", "find_fields:sales", False)

        # 第 2 轮：再次查询同一内容
        warning = detect_action_conflict(state, "find_fields:sales")
        assert warning is None
        update_loop_state(state, "find_fields", "find_fields:sales", False)

        # 第 3 轮：第三次查询，触发警告
        warning = detect_action_conflict(state, "find_fields:sales")
        assert warning is not None
        assert "第 3 次" in warning
        update_loop_state(state, "find_fields", "find_fields:sales", False)

        # 验证最终状态
        assert state.turn_index == 3
        assert state.tool_call_history == ["find_fields", "find_fields", "find_fields"]
        assert state.cache_hit_counts["find_fields:sales"] == 3
        assert state.consecutive_errors == 0

    def test_mixed_success_and_errors(self):
        """混合场景：部分成功、部分错误、部分缓存命中"""
        state = LoopState()

        # 成功调用
        update_loop_state(state, "find_fields", None, False)

        # 错误
        update_loop_state(state, "read_range", None, True)

        # 缓存命中（无警告）
        detect_action_conflict(state, "key1")
        update_loop_state(state, "find_fields", "key1", False)

        # 再次错误
        update_loop_state(state, "read_range", None, True)

        # 缓存命中 key1 第 2 次
        detect_action_conflict(state, "key1")
        update_loop_state(state, "find_fields", "key1", False)

        # 验证状态
        assert state.turn_index == 5
        assert state.consecutive_errors == 0  # 最后一次成功，归零
        assert state.cache_hit_counts["key1"] == 2

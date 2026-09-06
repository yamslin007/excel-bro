from __future__ import annotations

import json

from server.app.excel_agent import AGENT_SYSTEM_PROMPT, AGENT_TOOLS
from server.app.intent import INTENT_SYSTEM_PROMPT


def submit_plan_tool() -> dict:
    for tool in AGENT_TOOLS:
        if tool["function"]["name"] == "submit_plan":
            return tool
    raise AssertionError("submit_plan 工具不存在")


def test_system_prompt_ranks_live_writes_above_baked_values() -> None:
    """写入方式的优先级必须写在系统提示词里，且顺序正确。

    这段是 computeValues 能否真正被用上的关键 —— 模型默认倾向把算好的结果塞进
    writeValues，不明确排序就会退回烘焙快照。
    """
    prompt = AGENT_SYSTEM_PROMPT
    formulas = prompt.index("writeFormulas ——")
    compute = prompt.index("computeValues / splitGroupAggregate ——")
    baked = prompt.index("writeValues / writeTable ——")

    assert formulas < compute < baked


def test_system_prompt_forbids_baking_computed_results() -> None:
    assert "不要把算出来的结果塞进 writeValues.values" in AGENT_SYSTEM_PROMPT


def test_system_prompt_scopes_data_results_to_answers_not_writes() -> None:
    """本地数据工具结果只能用来回答，不能当静态值写回工作表。

    这条曾是烘焙的直接指令源（原文让模型「优先直接使用其 headers、rows」，
    没区分回答和写入）。
    """
    prompt = AGENT_SYSTEM_PROMPT
    assert "本地数据工具结果" in prompt
    assert "不要把这些 rows 当静态值塞进 writeValues" in prompt


def test_submit_plan_schema_repeats_the_priority() -> None:
    """双管齐下：单靠系统提示词在不同模型上服从度差异大。"""
    description = submit_plan_tool()["function"]["description"]

    assert "computeValues" in description
    assert "writeFormulas" in description
    assert description.index("writeFormulas") < description.index("computeValues")


def test_submit_plan_schema_exposes_compute_values_action() -> None:
    """模型看不到动作定义就不可能选它。"""
    schema = submit_plan_tool()["function"]["parameters"]
    serialized = json.dumps(schema, ensure_ascii=False)

    assert "computeValues" in serialized


def test_intent_prompt_distinguishes_sheet_split_from_inner_grouping() -> None:
    assert "输出工作表名称来自哪个字段" in INTENT_SYSTEM_PROMPT
    assert "每张结果表内部按" in INTENT_SYSTEM_PROMPT
    assert "只指定一个工作表命名字段" in INTENT_SYSTEM_PROMPT
    assert "不要再把它描述为" in INTENT_SYSTEM_PROMPT
    assert "不要把“B 分列显示”这类含义不清的短语作为选项" in INTENT_SYSTEM_PROMPT
    assert "按 B 建表，每张表内按 A 汇总" in INTENT_SYSTEM_PROMPT
    assert "resolutionMode" in INTENT_SYSTEM_PROMPT
    assert "使用 replace" in INTENT_SYSTEM_PROMPT


def test_split_aggregate_prompt_and_schema_forbid_redundant_group_field() -> None:
    assert "绝不能为了让该字段显示在结果中" in AGENT_SYSTEM_PROMPT
    assert "不要提交 null" in AGENT_SYSTEM_PROMPT
    description = submit_plan_tool()["function"]["description"]
    assert "splitBy 会自动成为结果列" in description
    assert "groupBy 不得重复包含 splitBy" in description
    assert "不要传 null" in description


def test_intent_prompt_rejects_a_degenerate_ratio_denominator() -> None:
    assert "分母覆盖该分组中的全部拆分值" in INTENT_SYSTEM_PROMPT
    assert "每个非空占比恒为 100%" in INTENT_SYSTEM_PROMPT


def test_intent_prompt_makes_explicit_new_commands_override_prior_tasks() -> None:
    assert "当前用户需求始终是本轮唯一主任务" in INTENT_SYSTEM_PROMPT
    assert "自身已经包含明确的动作和目标" in INTENT_SYSTEM_PROMPT
    assert "不得用上一轮 confirmedPrompt 替换当前输入" in INTENT_SYSTEM_PROMPT


def test_intent_prompt_separates_read_only_count_from_write_back() -> None:
    assert "某个对象有多少条记录" in INTENT_SYSTEM_PROMPT
    assert "必须返回 tool_request" in INTENT_SYSTEM_PROMPT
    assert "统计记录数并写入汇总表" in INTENT_SYSTEM_PROMPT


def test_split_aggregate_uses_exact_header_context_without_sampling_rows() -> None:
    prompt = AGENT_SYSTEM_PROMPT

    assert "工作簿上下文已经包含已授权工作表的准确表头" in prompt
    assert "直接\n  提交计划" in prompt
    assert "输出工作表名称来自哪个字段，该字段就是 splitBy" in prompt
    assert "不要调用 read_range 抽样计数" in prompt

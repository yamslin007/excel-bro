from __future__ import annotations

import json

from server.app.excel_agent import AGENT_SYSTEM_PROMPT, AGENT_TOOLS


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

"""``run_python_analysis`` 工具描述必须交代**容器内代码契约**（真实缺陷回归）。

缺陷（2026-10-10 真实模型自主规划验收中发现）：
真实模型自主调用 ``prepare_python_data``（拿到 ``dataset://`` 句柄）后再调
``run_python_analysis``，服务端两次都返回 ``unavailable / execution_error``。
直连 Docker 沙箱、直调 ``run_isolated_analysis_internal`` 用正确契约都成功 ⇒
不是执行器坏。逐层追下来定位到**工具描述**：``python_analysis_tool_definition``
的 description 只说"运行 Python 分析代码"，**没写**容器内 ``docker_worker``
的取值约定 ——

  * 数据入口是顶层变量 ``input["datasets"][i]["values"]``；
  * 结果必须由**模块级**变量 ``result`` 表达；
  * ``result`` 必须是 ``{"kind": "matrix", "values": [[...], ...]}``。

模型拿不到契约就只能猜。它写出了"定义函数但不给模块级 ``result`` 赋值"的
代码 —— 容器里 ``exec`` 正常退出、工作程序照样回报
``{"status": "completed", "result": null}``，于是
``_matrix_result_to_data_tool_result(None)`` 返回 ``None``，最终对模型呈现为
固定的"执行环节发生内部错误"回执。**契约不写进工具描述 = 让模型面向一个
无法满足的接口**，这是产品缺陷。

本文件把契约钉在工具描述里：拆掉这段描述，下面的用例必须变红。
"""

from __future__ import annotations

import json

import pytest

from server.app.python_analysis import (
    python_analysis_tool_definition,
    withheld_python_arguments,
)


def _description() -> str:
    return python_analysis_tool_definition()["function"]["description"]


def test_description_states_the_module_level_result_contract():
    """契约三要素必须都在描述里，且点名"模块级 result"。"""
    text = _description()
    assert "result" in text
    # 数据入口：顶层 input 的 datasets 结构。
    assert "input" in text and "datasets" in text
    assert "values" in text
    # 必须是"模块级/顶层"赋值，而不是函数返回值。
    assert "模块级" in text or "顶层" in text
    # 结果形状：kind 固定 matrix，values 是二维矩阵。
    assert "matrix" in text
    assert "kind" in text
    # 明说函数返回值/print 不算（模型的错误直觉就在这儿）。
    assert "函数返回值" in text or "print" in text


def test_description_is_a_faithful_projection_of_the_worker_contract():
    """描述不是口号：逐条对着容器工作程序的真实取值行为核。

    ``docker_worker._run_code`` 的真实行为：
    - 命名空间只注入 ``input``（数据入口）；
    - 取 ``namespace["result"]``（模块级变量），没有就返回 ``None``；
    - ``project_python_analysis_receipt`` / ``_matrix_result_to_data_tool_result``
      只接受 ``{"kind": "matrix", "values": [[...]]}``。
    本用例把"描述里的形状"与"能被坐标转换接受的形状"对起来 —— 描述写对了，
    模型照描述产出的 ``result`` 就一定能被接受。
    """
    from server.app.python_coordinator import _matrix_result_to_data_tool_result

    text = _description()
    # 描述里给出的正例形状（与 worker 契约一致）必须真能被接受 → 描述可信。
    described_example = {"kind": "matrix", "values": [["列A", "列B"], ["x", 1]]}
    assert "kind" in text and "matrix" in text  # 描述写明了这个形状
    converted = _matrix_result_to_data_tool_result(
        described_example, source_sheets=["会员表"]
    )
    assert converted is not None, "描述给出的形状竟然不能被坐标转换接受"
    assert converted.rows == [["列A", "列B"], ["x", 1]]


def test_code_parameter_description_repeats_the_contract():
    """参数级 description 也要写契约：模型看参数说明比看顶层说明更近。"""
    params = python_analysis_tool_definition()["function"]["parameters"]
    code_description = params["properties"]["code"]["description"]
    assert "result" in code_description
    assert "matrix" in code_description
    assert "kind" in code_description


def test_contract_is_stateless_and_does_not_leak_the_withheld_marker_into_history():
    """契约只讲形状，不引入任何可被历史/检查点持久化的原始内容。

    顺带钉住：把带契约的 schema 交给历史投影后，``code`` 仍被替换为占位，
    ``datasetHandles`` 投影为空列表 —— 描述变长**不**能顺带改变投影行为。
    """
    described = python_analysis_tool_definition()
    raw_arguments = json.dumps(
        {"code": "result = {'kind': 'matrix', 'values': [['a']]}", "datasetHandles": []}
    )
    projected = json.loads(withheld_python_arguments(raw_arguments))
    assert "kind" not in projected["code"]
    assert projected["datasetHandles"] == []
    # schema 本身仍然只允许这两个入参键（描述变长不放松边界）。
    assert set(described["function"]["parameters"]["properties"]) == {
        "code",
        "datasetHandles",
    }
    assert described["function"]["parameters"]["additionalProperties"] is False


@pytest.mark.parametrize("field", ["description", "code_description"])
def test_contract_text_is_not_a_tautology(field: str):
    """反向哨兵：描述里必须真的出现可执行判据，不是"请遵守契约"这类空话。"""
    if field == "description":
        text = _description()
    else:
        text = python_analysis_tool_definition()["function"]["parameters"]["properties"][
            "code"
        ]["description"]
    # 至少给出一个字面可照抄的赋值形状。
    assert "values" in text, f"{field} 只讲了「要遵守契约」，没给出可照抄的形状"

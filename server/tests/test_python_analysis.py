"""隔离 Python 分析：失败关闭、有类型回执、严格入参校验与隐私闭环。

设计依据：``docs/SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md``。
本批没有任何隔离执行器实现——全部测试钉在四条失败关闭纪律上：
1. 没有执行器就绝不运行代码、绝不回显模型入参；
2. 原始 Python 代码不进入模型历史、持久检查点、事件、日志或 SDK 转录；
3. 回执是有类型的：字段值、状态组合、固定文案全部可信侧生成；
4. 入参在实际调用入口严格校验，不依赖 schema 声明或框架兜底。
"""

from __future__ import annotations

import asyncio
import copy
import json
import sqlite3
from types import SimpleNamespace
from typing import Any

import pytest

import server.app.excel_agent as excel_agent
from server.app import python_analysis
from server.app.agent_loop_protection import LoopConfig
from server.app.capabilities import capabilities
from server.app.claude_sdk_agent import _APPROVED_NON_TERMINAL_TOOLS
from server.app.llm import ModelConnection
from server.app.models import PlanRequest
from server.app.task_ledger import FolderOperationLedger

_CODE_SENTINEL = "PY_ANALYSIS_CODE_SENTINEL_9f31"
_HANDLE_SENTINEL = "dataset://PY_HANDLE_SENTINEL_7c02"
# 与生产实现一致的固定占位（字面量引用，避免实现未落地时 ImportError）。
_CODE_WITHHELD_MARKER = "[原始代码未随会话保存：如需重试请重新提交完整代码]"
# 有类型回执的固定文案（与 python_analysis 固定文案一致）。
_DISABLED_MESSAGE = "隔离 Python 未启用：功能保持关闭，本轮未运行任何代码。"
_RUNNER_MISSING_MESSAGE = "隔离 Python 不可用：尚未配置已验证的隔离执行器，本轮未运行任何代码。"
_CODE_NOT_PRESERVED_MESSAGE = "原始代码未随会话保存，无法继续上次的 Python 调用；请重新提供完整分析代码。"
_REJECTED_MESSAGE = "分析请求未通过校验：代码或数据句柄格式不正确，本轮未运行任何代码。"
_PROJECTION_FAILED_CONTENT = json.dumps({
    "status": "unavailable",
    "reasonCategory": "projection_failed",
    "message": "工具结果未通过出站校验，本轮内容不予下发。",
}, ensure_ascii=False)


def _request() -> PlanRequest:
    """构造只授权一个工作表的最小规划请求（与 test_agent_integration 同形）。"""
    return PlanRequest.model_validate(
        {
            "prompt": "请用隔离 Python 分析得分分布并给出计划",
            "workbook": {
                "name": "scores.xlsx",
                "capturedAt": "2026-10-06T00:00:00Z",
                "activeWorksheet": "Sheet1",
                "worksheets": [
                    {
                        "name": "Sheet1",
                        "usedRange": "A1:B3",
                        "rowCount": 3,
                        "columnCount": 2,
                        "headers": ["人员", "得分"],
                        "dataRows": [["甲", 10], ["乙", 20]],
                        "truncated": False,
                    }
                ],
            },
        }
    )


def _connection() -> ModelConnection:
    return ModelConnection(
        base_url="https://api.example.test/v1",
        model="test-model",
        api_key="",
        supports_vision=False,
    )


def _enable_python(monkeypatch: pytest.MonkeyPatch) -> None:
    """两道独立的开关闸：工具表注入（excel_agent）与回执分支（python_analysis）。"""
    monkeypatch.setattr(excel_agent, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)
    # These are missing-runner/privacy tests, independent of a developer's
    # running Docker Desktop. Ready-but-unwired behavior has its own test.
    monkeypatch.setattr(python_analysis, "_runner_status", lambda: None)


# ---------------------------------------------------------------------------
# 1. 默认关闭与失败关闭
# ---------------------------------------------------------------------------


def test_default_config_disables_python_execution():
    """契约：默认关闭。config 里 enabled 必须是显式 false（不是缺失）。"""
    capabilities.cache_clear()
    section = capabilities().get("pythonExecution")
    assert isinstance(section, dict)
    assert section["enabled"] is False
    assert python_analysis.python_execution_enabled() is False


def test_disabled_switch_returns_disabled_receipt_even_for_malformed_request():
    """开关关闭时连入参校验都不进入：畸形请求也只拿固定 disabled 回执。"""
    capabilities.cache_clear()
    receipt = python_analysis.execute_python_analysis(_request(), {"code": 42})
    assert receipt["status"] == "unavailable"
    assert receipt["reasonCategory"] == "disabled"
    assert receipt["message"] == _DISABLED_MESSAGE


def test_enabled_without_runner_fails_closed_and_runs_nothing(monkeypatch):
    """即使开关被打开，没有执行器也绝不运行任何代码。

    把宿主的所有常见执行入口全部换成"一旦被调用就炸"的哨兵，再调用工具：
    回执必须是固定的 runner_missing，哨兵一个都不能被碰到。
    """
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)

    def _forbidden(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("失败关闭被绕过：宿主执行入口在无执行器时被调用")

    import os
    import subprocess

    for target in (
        (subprocess, "Popen"), (subprocess, "run"), (subprocess, "call"),
        (subprocess, "check_output"), (subprocess, "check_call"),
        (os, "system"), (os, "execv"), (os, "execve"), (os, "spawnv"),
        (os, "startfile"),
    ):
        # 注意：不 patch builtins.compile——pytest 断言重写/异常格式化本身要
        # 编译代码，会误伤测试运行器；compile/exec/eval 由源码扫描测试把守。
        monkeypatch.setattr(target[0], target[1], _forbidden)

    receipt = python_analysis.execute_python_analysis(
        _request(),
        {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_SENTINEL]},
    )
    assert receipt == {
        "status": "unavailable",
        "reasonCategory": "runner_missing",
        "message": _RUNNER_MISSING_MESSAGE,
    }
    # 回执绝不回显模型入参（代码与句柄都不出站）。
    assert _CODE_SENTINEL not in json.dumps(receipt, ensure_ascii=False)
    assert _HANDLE_SENTINEL not in json.dumps(receipt, ensure_ascii=False)


def test_coordinator_source_contains_no_execution_tokens():
    """源码级守卫：可信协调器本体不允许出现任何执行/逃逸能力词。

    执行器将来必须落在独立模块并通过实测健康检查接入；谁往本模块里加
    子进程、求值、网络或导入逃逸，这条测试立刻变红。
    """
    source = python_analysis.__doc__ or ""
    module_file = python_analysis.__file__
    assert module_file
    text = open(module_file, encoding="utf-8").read()
    for token in (
        "subprocess", "Popen", "os.system", "os.exec", "os.spawn", "startfile",
        "exec(", "eval(", "compile(", "__import__", "importlib",
        "socket", "urllib", "requests", "httpx", "pty.",
        "multiprocessing", "win32", "ctypes",
    ):
        assert token not in text, f"协调器源码出现禁止的执行能力词：{token}"
    assert source  # 模块必须有边界声明文档


# ---------------------------------------------------------------------------
# 2. 缺陷 3：实际入参严格校验（不依赖 schema 声明或框架兜底）
# ---------------------------------------------------------------------------


def test_extra_fields_rejected_at_actual_entry(monkeypatch):
    """额外 path/apiKey/command 字段必须在真实调用入口被拒绝，且不回显。"""
    _enable_python(monkeypatch)
    receipt = excel_agent._execute_tool(
        _request(),
        "run_python_analysis",
        {
            "code": "print(1)",
            "path": "C:/PRIVATE_PATH_SENTINEL",
            "apiKey": "SK_SECRET_KEY_SENTINEL",
            "command": "rm -rf /",
        },
    )
    assert isinstance(receipt, dict)
    assert receipt["status"] == "rejected"
    assert receipt["reasonCategory"] == "invalid_request"
    dumped = json.dumps(receipt, ensure_ascii=False)
    assert "PRIVATE_PATH_SENTINEL" not in dumped
    assert "SK_SECRET_KEY_SENTINEL" not in dumped
    assert "path" not in receipt and "apiKey" not in receipt and "command" not in receipt


def _limits() -> dict[str, int]:
    return python_analysis.python_execution_limits()


@pytest.mark.parametrize(
    "arguments",
    [
        {},
        {"code": ""},
        {"code": "   "},
        {"code": 42},
        {"code": None},
        {"code": True},
        {"code": ["print(1)"]},
        {"code": {"code": "print(1)"}},
        {"code": "x" * 40000},  # 超长（真实上限由 capabilities 决定，必小于 40000）
        {"code": "print(1)", "datasetHandles": "dataset://x"},
        {"code": "print(1)", "datasetHandles": 123},
        {"code": "print(1)", "datasetHandles": [123]},
        {"code": "print(1)", "datasetHandles": [True]},
        {"code": "print(1)", "datasetHandles": [None]},
        {"code": "print(1)", "datasetHandles": [["nested"]]},
        {"code": "print(1)", "datasetHandles": [""]},
        {"code": "print(1)", "datasetHandles": ["ok", "  "]},
        {"code": "print(1)", "datasetHandles": ["h" * 4096]},  # 单项超长
        {"code": "print(1)", "datasetHandles": [f"h{i}" for i in range(64)]},  # 数量超限
    ],
)
def test_invalid_arguments_are_rejected_when_enabled(monkeypatch, arguments):
    _enable_python(monkeypatch)
    receipt = python_analysis.execute_python_analysis(_request(), arguments)
    assert receipt == {
        "status": "rejected",
        "reasonCategory": "invalid_request",
        "message": _REJECTED_MESSAGE,
    }


def test_boundary_lengths_follow_capabilities(monkeypatch):
    """边界长度必须严格按 capabilities 执行：上限内放行（失败关闭到 runner_missing），
    超限一位即拒绝；句柄数量与单项长度同理。"""
    _enable_python(monkeypatch)
    limits = _limits()
    assert limits["maxCodeChars"] < 40000  # 保证上面的超长参数确实越界

    ok_code = "x" * limits["maxCodeChars"]
    receipt = python_analysis.execute_python_analysis(_request(), {"code": ok_code})
    assert receipt["status"] == "unavailable"

    over_code = "x" * (limits["maxCodeChars"] + 1)
    receipt = python_analysis.execute_python_analysis(_request(), {"code": over_code})
    assert receipt["status"] == "rejected"

    ok_handles = [f"handle-{i}" for i in range(limits["maxDatasetHandles"])]
    receipt = python_analysis.execute_python_analysis(
        _request(), {"code": "print(1)", "datasetHandles": ok_handles}
    )
    assert receipt["status"] == "unavailable"

    many_handles = [f"handle-{i}" for i in range(limits["maxDatasetHandles"] + 1)]
    receipt = python_analysis.execute_python_analysis(
        _request(), {"code": "print(1)", "datasetHandles": many_handles}
    )
    assert receipt["status"] == "rejected"

    ok_item = "h" * limits["maxDatasetHandleChars"]
    receipt = python_analysis.execute_python_analysis(
        _request(), {"code": "print(1)", "datasetHandles": [ok_item]}
    )
    assert receipt["status"] == "unavailable"

    over_item = "h" * (limits["maxDatasetHandleChars"] + 1)
    receipt = python_analysis.execute_python_analysis(
        _request(), {"code": "print(1)", "datasetHandles": [over_item]}
    )
    assert receipt["status"] == "rejected"


def test_rejection_receipt_never_echoes_arguments(monkeypatch):
    """拒绝回执只含固定文案：字段值、代码、路径、密钥一概不出站。"""
    _enable_python(monkeypatch)
    receipt = python_analysis.execute_python_analysis(
        _request(),
        {"code": f"import os\nos.system('{_CODE_SENTINEL}')", "path": _CODE_SENTINEL},
    )
    dumped = json.dumps(receipt, ensure_ascii=False)
    assert _CODE_SENTINEL not in dumped
    assert receipt["message"] == _REJECTED_MESSAGE


def test_valid_arguments_still_fail_closed_when_enabled(monkeypatch):
    """形状合法也到不了执行：runner_missing 是启用态唯一可达回执。"""
    _enable_python(monkeypatch)
    receipt = python_analysis.execute_python_analysis(
        _request(), {"code": "print(1)", "datasetHandles": [_HANDLE_SENTINEL]}
    )
    assert receipt["status"] == "unavailable"
    assert receipt["reasonCategory"] == "runner_missing"


def test_sdk_rejects_extra_fields_identically(monkeypatch, tmp_path):
    """SDK 后端与普通后端共用同一严格校验：额外字段同样被拒。"""
    from server.app.claude_sdk_agent import run_sdk_plan

    _enable_python(monkeypatch)
    fake = _FakeSdk(
        response=json.dumps({"kind": "answer", "message": "收到"}),
        invoke=("run_python_analysis", {
            "code": "print(1)",
            "path": "C:/PRIVATE_PATH_SENTINEL",
            "apiKey": "SK_SECRET_KEY_SENTINEL",
        }),
    )
    _install_fake_sdk(monkeypatch, tmp_path, fake)
    asyncio.run(run_sdk_plan(_sdk_plan_request(), _connection()))
    assert fake.tool_result is not None
    assert fake.tool_result.get("is_error") is not True
    parsed = json.loads(fake.tool_result["content"][0]["text"])
    assert parsed == {
        "status": "rejected",
        "reasonCategory": "invalid_request",
        "message": _REJECTED_MESSAGE,
    }
    assert "PRIVATE_PATH_SENTINEL" not in fake.tool_result["content"][0]["text"]


# ---------------------------------------------------------------------------
# 3. 缺陷 2：有类型回执（字段值、状态组合、固定文案全部可信侧生成）
# ---------------------------------------------------------------------------


def test_receipt_message_regenerated_from_trusted_state():
    """message 按可信状态生成固定文案：自由文本（含哨兵）一律被替换。"""
    for (status, reason), fixed_message in {
        ("unavailable", "disabled"): _DISABLED_MESSAGE,
        ("unavailable", "runner_missing"): _RUNNER_MISSING_MESSAGE,
        ("unavailable", "code_not_preserved"): _CODE_NOT_PRESERVED_MESSAGE,
        ("rejected", "invalid_request"): _REJECTED_MESSAGE,
    }.items():
        receipt = {"status": status, "reasonCategory": reason, "message": f"自由文本 {_CODE_SENTINEL}"}
        projected = python_analysis.project_python_analysis_receipt(receipt)
        assert projected == {"status": status, "reasonCategory": reason, "message": fixed_message}
        assert _CODE_SENTINEL not in json.dumps(projected, ensure_ascii=False)


def test_receipt_rejects_unknown_status_and_combinations():
    """未知状态、非字符串状态、错误状态组合、缺 reasonCategory 一律失败关闭。"""
    bad_receipts = [
        {"status": "completed", "reasonCategory": "runner_missing"},
        {"status": 123, "reasonCategory": "runner_missing"},
        {"status": None, "reasonCategory": "runner_missing"},
        {"status": "unavailable", "reasonCategory": "invalid_request"},  # 组合错误
        {"status": "rejected", "reasonCategory": "runner_missing"},  # 组合错误
        {"status": "rejected", "reasonCategory": "code_not_preserved"},  # 组合错误
        {"status": "unavailable"},
        {"reasonCategory": "runner_missing"},
    ]
    for receipt in bad_receipts:
        with pytest.raises(ValueError):
            python_analysis.project_python_analysis_receipt(receipt)
    # 合法组合单独确认
    projected = python_analysis.project_python_analysis_receipt(
        {"status": "unavailable", "reasonCategory": "code_not_preserved"}
    )
    assert projected["message"] == _CODE_NOT_PRESERVED_MESSAGE


def test_receipt_rejects_forged_ids_handles_and_unknown_fields():
    """没有运行器与结果登记：runId/resultHandle 出现即伪造；白名单外字段拒绝。"""
    bad_receipts = [
        {"status": "unavailable", "reasonCategory": "runner_missing", "runId": "py-run-1"},
        {"status": "unavailable", "reasonCategory": "runner_missing", "resultHandle": "result://ok"},
        {"status": "unavailable", "reasonCategory": "runner_missing", "runId": _CODE_SENTINEL},
        {"status": "unavailable", "reasonCategory": "runner_missing", "stdout": "STDOUT_SENTINEL"},
        {"status": "unavailable", "reasonCategory": "runner_missing", "rows": [["数据行哨兵"]]},
        {"status": "unavailable", "reasonCategory": "runner_missing", "traceback": "TRACEBACK_SENTINEL"},
        {"status": "unavailable", "reasonCategory": "runner_missing", "code": _CODE_SENTINEL},
    ]
    for receipt in bad_receipts:
        with pytest.raises(ValueError):
            python_analysis.project_python_analysis_receipt(receipt)


def test_receipt_projection_rejects_non_objects():
    with pytest.raises(ValueError, match="必须是对象"):
        python_analysis.project_python_analysis_receipt("not a dict")  # type: ignore[arg-type]
    with pytest.raises(ValueError, match="必须是对象"):
        python_analysis.project_python_analysis_receipt(["unavailable"])  # type: ignore[arg-type]


def test_outbound_gate_blocks_unregistered_non_terminal_tools():
    """出站闸：没有登记投影的工具，结果一律不给模型。"""
    with pytest.raises(ValueError, match="没有登记模型出站投影"):
        excel_agent._non_terminal_tool_content("map_column_values", {"any": "thing"})
    with pytest.raises(ValueError, match="没有登记模型出站投影"):
        excel_agent._non_terminal_tool_content("totally_unknown_tool", {"any": "thing"})


def test_outbound_gate_passes_python_receipt_and_read_range(monkeypatch):
    # 开启双闸后走真实路径：合法入参通过校验 → 无隔离运行器 → runner_missing 回执。
    _enable_python(monkeypatch)
    receipt = python_analysis.execute_python_analysis(
        _request(), {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_SENTINEL]}
    )
    content = excel_agent._non_terminal_tool_content("run_python_analysis", receipt)
    parsed = json.loads(content)
    assert parsed == {
        "status": "unavailable",
        "reasonCategory": "runner_missing",
        "message": _RUNNER_MISSING_MESSAGE,
    }
    assert _CODE_SENTINEL not in content
    assert _HANDLE_SENTINEL not in content
    # 既有行为保持：read_range 走原样透传投影。
    assert excel_agent._non_terminal_tool_content(
        "read_range", {"values": [["x"]]}
    ) == json.dumps({"values": [["x"]]}, ensure_ascii=False)


def test_projection_failure_fails_closed_at_both_outlets(monkeypatch):
    """投影失败（任何 ValueError）时两个出口都返回固定安全内容，绝不 str(error)。"""
    error_with_sentinel = ValueError(f"boom {_CODE_SENTINEL}")

    def _broken_projector(_receipt: dict[str, Any]) -> dict[str, Any]:
        raise error_with_sentinel

    monkeypatch.setitem(
        excel_agent.NON_TERMINAL_TOOL_PROJECTIONS,
        "run_python_analysis",
        _broken_projector,
    )
    # 普通后端出站闸
    content = excel_agent._non_terminal_tool_content(
        "run_python_analysis", {"status": "unavailable", "reasonCategory": "runner_missing"}
    )
    assert content == _PROJECTION_FAILED_CONTENT
    assert _CODE_SENTINEL not in content


def test_projection_failure_fails_closed_in_real_loop(monkeypatch):
    """真实主循环里投影失败：role=tool 内容是固定安全文案，循环继续。"""
    error_with_sentinel = ValueError(f"boom {_CODE_SENTINEL}")

    def _broken_projector(_receipt: dict[str, Any]) -> dict[str, Any]:
        raise error_with_sentinel

    monkeypatch.setitem(
        excel_agent.NON_TERMINAL_TOOL_PROJECTIONS,
        "run_python_analysis",
        _broken_projector,
    )
    _enable_python(monkeypatch)
    plan = {"id": "p3", "title": "清空内容", "summary": "清空已授权区域", "actions": [
        {"type": "clearRange", "sheet": "Sheet1", "range": "B2:B3", "applyTo": "contents"}
    ]}

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {"code": "print(1)", "datasetHandles": []},
                    "py1",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": plan}, "plan"), _usage())

    client = _FakeClient(responder)
    _install_client(monkeypatch, client)
    result = asyncio.run(
        excel_agent.run_excel_agent(
            _request(), connection=_connection(), timeout=5,
            config=LoopConfig(max_wall_time_seconds=3600),
        )
    )
    assert result.kind == "plan"
    second_messages = client.message_snapshots[1]
    tool_message = next(m for m in second_messages if m.get("role") == "tool")
    assert tool_message["content"] == _PROJECTION_FAILED_CONTENT
    assert _CODE_SENTINEL not in tool_message["content"]


def test_projection_failure_fails_closed_in_sdk(monkeypatch, tmp_path):
    """SDK 出口同样失败关闭：固定安全内容，不泄露错误文本。"""
    from server.app.claude_sdk_agent import run_sdk_plan

    error_with_sentinel = ValueError(f"boom {_CODE_SENTINEL}")

    def _broken_projector(_receipt: dict[str, Any]) -> dict[str, Any]:
        raise error_with_sentinel

    monkeypatch.setitem(
        excel_agent.NON_TERMINAL_TOOL_PROJECTIONS,
        "run_python_analysis",
        _broken_projector,
    )
    _enable_python(monkeypatch)
    fake = _FakeSdk(
        response=json.dumps({"kind": "answer", "message": "收到"}),
        invoke=("run_python_analysis", {"code": "print(1)", "datasetHandles": []}),
    )
    _install_fake_sdk(monkeypatch, tmp_path, fake)
    asyncio.run(run_sdk_plan(_sdk_plan_request(), _connection()))
    assert fake.tool_result is not None
    assert fake.tool_result["content"][0]["text"] == _PROJECTION_FAILED_CONTENT
    assert _CODE_SENTINEL not in fake.tool_result["content"][0]["text"]


# ---------------------------------------------------------------------------
# 4. 缺陷 1：原始代码不进入模型历史 / 持久检查点 / 事件 / SDK 转录
# ---------------------------------------------------------------------------


class _FakeClient:
    """按回调生成每轮模型响应，并保留传入消息快照。"""

    def __init__(self, responder):
        self.responder = responder
        self.call_count = 0
        self.message_snapshots: list[list[dict[str, Any]]] = []
        self.call_kwargs: list[dict[str, Any]] = []

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc, tb):
        return False

    async def chat_completions(
        self, *, messages: list[dict[str, Any]], **_kwargs: Any
    ) -> tuple[dict[str, Any], dict[str, int]]:
        self.call_count += 1
        self.message_snapshots.append(copy.deepcopy(messages))
        self.call_kwargs.append(copy.deepcopy(_kwargs))
        return self.responder(self.call_count)


def _tool_payload(name: str, arguments: dict[str, Any], call_id: str) -> dict[str, Any]:
    return {
        "choices": [
            {
                "message": {
                    "content": None,
                    "tool_calls": [
                        {
                            "id": call_id,
                            "type": "function",
                            "function": {
                                "name": name,
                                "arguments": json.dumps(arguments, ensure_ascii=False),
                            },
                        }
                    ],
                }
            }
        ]
    }


def _content_payload(content: str) -> dict[str, Any]:
    return {"choices": [{"message": {"content": content, "tool_calls": []}}]}


def _usage(total_tokens: int = 0) -> dict[str, int]:
    return {"prompt_tokens": total_tokens, "completion_tokens": 0, "total_tokens": total_tokens}


def _install_client(monkeypatch: pytest.MonkeyPatch, client: _FakeClient) -> None:
    monkeypatch.setattr(
        excel_agent, "OpenAICompatibleClient", lambda *_args, **_kwargs: client
    )


_PLAN = {"id": "p1", "title": "清空内容", "summary": "清空已授权区域", "actions": [
    {"type": "clearRange", "sheet": "Sheet1", "range": "B2:B3", "applyTo": "contents"}
]}


def _python_then_plan_responder():
    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_SENTINEL]},
                    "py1",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    return responder


def _run_agent_with_client(monkeypatch, responder, *, resume_checkpoint=None, on_checkpoint=None, on_event=None):
    client = _FakeClient(responder)
    _install_client(monkeypatch, client)
    result = asyncio.run(
        excel_agent.run_excel_agent(
            _request(), connection=_connection(), timeout=5,
            config=LoopConfig(max_wall_time_seconds=3600),
            resume_checkpoint=resume_checkpoint,
            on_checkpoint=on_checkpoint,
            on_event=on_event,
        )
    )
    return client, result


def test_python_tool_not_offered_by_default(monkeypatch):
    """默认关闭时工具绝不出现在模型工具表里（真实配置，无 monkeypatch 开关）。"""
    capabilities.cache_clear()
    client, result = _run_agent_with_client(
        monkeypatch, lambda _: (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())
    )
    assert result.kind == "plan"
    names = {tool["function"]["name"] for tool in client.call_kwargs[0]["tools"]}
    assert "run_python_analysis" not in names


def test_python_code_never_reenters_model_requests(monkeypatch):
    """缺陷 1（模型出口）：完整模型请求里不出现原始代码——不只检查 role=tool。

    捕获每一轮完整请求；assistant 消息的 run_python_analysis arguments 必须
    只含固定占位（协议配对保留，datasetHandles 收口为空列表）。
    """
    _enable_python(monkeypatch)
    client, result = _run_agent_with_client(monkeypatch, _python_then_plan_responder())
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _CODE_SENTINEL not in blob, f"第 {turn_index + 1} 轮模型请求携带原始代码"
    # 占位语义与协议配对：第二轮请求里的 assistant 工具调用参数可解析、
    # code 是固定占位、datasetHandles 收口为空列表（未验证句柄不保留原值）、
    # tool_call_id 与回执配对。
    second_messages = client.message_snapshots[1]
    assistant_message = next(
        m for m in second_messages
        if m.get("role") == "assistant" and m.get("tool_calls")
    )
    call = assistant_message["tool_calls"][0]
    assert call["id"] == "py1"
    arguments = json.loads(call["function"]["arguments"])
    assert arguments["code"] == _CODE_WITHHELD_MARKER
    assert arguments["datasetHandles"] == []
    tool_message = next(m for m in second_messages if m.get("role") == "tool")
    assert tool_message["tool_call_id"] == "py1"
    # 回执内容依旧不含代码。
    assert _CODE_SENTINEL not in tool_message["content"]


def test_python_code_never_enters_sqlite_persistence(monkeypatch, tmp_path):
    """缺陷 1（持久化出口）：真实账本链路下 SQLite 全字段无原始代码。

    复刻 main.py 的真实接线：on_checkpoint → save_planning_checkpoint、
    on_event → append_planning_event，跑完再 pause_planning_task；
    原始行读回 planning_tasks 与 planning_task_events 全部列扫描。
    """
    _enable_python(monkeypatch)
    db_path = tmp_path / "planning-ledger.sqlite3"
    ledger = FolderOperationLedger(lambda: db_path)
    turn_id = "turn-python-privacy-1"
    ledger.register_planning_task(
        turn_id, "fingerprint-py-privacy", {"prompt": "synthetic"},
        approved_wall_time_seconds=3600,
    )
    checkpoints: list[dict[str, Any]] = []
    event_sequence = {"n": 0}

    async def on_checkpoint(checkpoint: dict[str, Any]) -> None:
        checkpoints.append(copy.deepcopy(checkpoint))
        ledger.save_planning_checkpoint(
            turn_id, checkpoint, active_elapsed_seconds=1.0
        )

    async def on_event(event: dict[str, Any]) -> None:
        event_sequence["n"] += 1
        ledger.append_planning_event(turn_id, event_sequence["n"], event)

    client, result = _run_agent_with_client(
        monkeypatch, _python_then_plan_responder(),
        on_checkpoint=on_checkpoint, on_event=on_event,
    )
    assert result.kind == "plan"
    assert checkpoints, "真实链路必须产生至少一个检查点"
    # 覆盖暂停（waiting_user）现场：pause_reason / error_json / checkpoint_json。
    ledger.pause_planning_task(
        turn_id, checkpoints[-1], active_elapsed_seconds=1.0,
        reason="budget-test", error=None,
    )

    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    task_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_tasks WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    event_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_task_events WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    connection.close()
    assert task_rows and event_rows
    for row in task_rows + event_rows:
        blob = json.dumps(row, ensure_ascii=False)
        assert _CODE_SENTINEL not in blob, f"持久化行携带原始代码：{row.keys()}"
    # 检查点仍保留协议配对（可恢复性不因隐私投影而破坏）。
    checkpoint_blob = json.dumps(checkpoints, ensure_ascii=False)
    assert '"tool_call_id": "py1"' in checkpoint_blob.replace(" ", "") or '"py1"' in checkpoint_blob


def test_resume_python_call_reports_code_not_preserved(monkeypatch):
    """缺陷 1（恢复现场）：重启后 Python 调用不可重建——明确 code_not_preserved，
    不冒充断点续跑；恢复后的模型请求同样不携带原始代码。"""
    _enable_python(monkeypatch)
    checkpoints: list[dict[str, Any]] = []

    async def on_checkpoint(checkpoint: dict[str, Any]) -> None:
        checkpoints.append(copy.deepcopy(checkpoint))

    client_first, result_first = _run_agent_with_client(
        monkeypatch, _python_then_plan_responder(), on_checkpoint=on_checkpoint
    )
    assert result_first.kind == "plan"
    pending = next(
        cp for cp in checkpoints
        if cp.get("stage") == "model_response_saved"
        and any(
            (call.get("function") or {}).get("name") == "run_python_analysis"
            for call in (cp.get("pendingAssistantMessage") or {}).get("tool_calls", [])
        )
    )
    assert pending["nextToolIndex"] == 0

    # 用第一轮真实产生的检查点恢复：新客户端第一轮直接 submit_plan。
    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client_resumed, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=pending
    )
    assert result.kind == "plan"
    for snapshot in client_resumed.message_snapshots:
        assert _CODE_SENTINEL not in json.dumps(snapshot, ensure_ascii=False)
    resumed_messages = client_resumed.message_snapshots[0]
    tool_message = next(m for m in resumed_messages if m.get("role") == "tool")
    receipt = json.loads(tool_message["content"])
    assert receipt["status"] == "unavailable"
    assert receipt["reasonCategory"] == "code_not_preserved"
    assert receipt["message"] == _CODE_NOT_PRESERVED_MESSAGE
    assert tool_message["tool_call_id"] == "py1"


# ---------------------------------------------------------------------------
# 5. SDK 后端：批准集与条件注册
# ---------------------------------------------------------------------------


class _FakeResult:
    def __init__(self, result: str = "") -> None:
        self.result = result
        self.is_error = False
        self.subtype = "success"
        self.api_error_status = None


class _FakeSdk:
    ResultMessage = _FakeResult
    CLINotFoundError = type("CLINotFoundError", (Exception,), {})
    CLIConnectionError = type("CLIConnectionError", (Exception,), {})
    ProcessError = type("ProcessError", (Exception,), {})
    ResultError = type("ResultError", (Exception,), {})
    CLIJSONDecodeError = type("CLIJSONDecodeError", (Exception,), {})

    def __init__(self, *, response: str = "", invoke: tuple[str, dict] | None = None) -> None:
        self.response = response
        self.invoke = invoke
        self.calls: list[tuple[str, Any]] = []
        self.tool_result: dict | None = None
        self.schemas: dict[str, dict] = {}

    def ClaudeAgentOptions(self, **kwargs: object):
        return SimpleNamespace(**kwargs)

    def tool(self, name: str, _description: str, _schema: dict):
        self.schemas[name] = _schema
        return lambda handler: SimpleNamespace(name=name, handler=handler)

    def create_sdk_mcp_server(self, *, name: str, version: str, tools: list):
        assert name == "excel_bro" and version
        return {tool.name: tool for tool in tools}

    async def query(self, *, prompt, options):
        if not isinstance(prompt, str):
            prompt = [message async for message in prompt]
        self.calls.append((prompt, options))
        if self.invoke is not None:
            name, args = self.invoke
            tool = options.mcp_servers["excel_bro"][name]
            self.tool_result = await tool.handler(args)
        yield _FakeResult(self.response)


def _install_fake_sdk(monkeypatch: pytest.MonkeyPatch, tmp_path, fake: _FakeSdk) -> None:
    from server.app import claude_sdk_agent

    monkeypatch.setattr(claude_sdk_agent, "_sdk", lambda: fake)
    monkeypatch.setattr(claude_sdk_agent, "_state_dir", lambda: tmp_path)


def _sdk_plan_request():
    from server.app.llm.config import CLAUDE_SDK_CATALOG_ID

    return PlanRequest.model_validate({
        "prompt": "请按类别填写等级",
        "planningReadPolicy": "none",
        "modelId": CLAUDE_SDK_CATALOG_ID,
        "workbook": {
            "name": "测试.xlsx", "capturedAt": "2026-10-01T00:00:00Z",
            "activeWorksheet": "Sheet1",
            "worksheets": [{
                "name": "Sheet1", "sourceFile": "PRIVATE_FULL_PATH",
                "usedRange": "A1:B3", "rowCount": 3, "columnCount": 2,
                "headers": ["类别", "等级"],
                "dataRows": [["PRIVATE_DATA_ROW", None], ["另一个值", None]],
            }],
        },
    })


def test_sdk_registers_python_tool_only_when_enabled(monkeypatch, tmp_path):
    from server.app.claude_sdk_agent import run_sdk_plan

    # 关闭（真实配置）：不注册、不进 allowed_tools。
    capabilities.cache_clear()
    fake_off = _FakeSdk(response=json.dumps({"kind": "answer", "message": "ok"}))
    _install_fake_sdk(monkeypatch, tmp_path, fake_off)
    asyncio.run(run_sdk_plan(_sdk_plan_request(), _connection()))
    assert "run_python_analysis" not in fake_off.schemas
    _off_prompt, off_options = fake_off.calls[-1]
    assert "mcp__excel_bro__run_python_analysis" not in off_options.allowed_tools

    # 开启：注册且进入 allowed_tools。
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)
    fake_on = _FakeSdk(response=json.dumps({"kind": "answer", "message": "ok"}))
    _install_fake_sdk(monkeypatch, tmp_path, fake_on)
    asyncio.run(run_sdk_plan(_sdk_plan_request(), _connection()))
    assert "run_python_analysis" in fake_on.schemas
    _prompt, options = fake_on.calls[-1]
    assert "mcp__excel_bro__run_python_analysis" in options.allowed_tools


def test_sdk_delivers_projected_receipt_not_refusal(monkeypatch, tmp_path):
    from server.app.claude_sdk_agent import run_sdk_plan

    _enable_python(monkeypatch)
    fake = _FakeSdk(
        response=json.dumps({"kind": "answer", "message": "收到"}),
        invoke=("run_python_analysis", {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_SENTINEL]}),
    )
    _install_fake_sdk(monkeypatch, tmp_path, fake)
    asyncio.run(run_sdk_plan(_sdk_plan_request(), _connection()))
    assert fake.tool_result is not None
    assert fake.tool_result.get("is_error") is not True
    text = fake.tool_result["content"][0]["text"]
    parsed = json.loads(text)
    assert parsed == {
        "status": "unavailable",
        "reasonCategory": "runner_missing",
        "message": _RUNNER_MISSING_MESSAGE,
    }
    assert "非终结工具结果不可用" not in text
    # 出站闸唯一入口：批准集里的工具也必须经投影，不允许旁路。
    assert _CODE_SENTINEL not in text
    assert _HANDLE_SENTINEL not in text


def test_sdk_still_refuses_unapproved_non_terminal_results(monkeypatch, tmp_path):
    """批准集之外的 dict 结果维持原拒绝行为，不因新工具而放开。"""
    from server.app.claude_sdk_agent import run_sdk_plan

    fake = _FakeSdk(
        response=json.dumps({"kind": "answer", "message": "ok"}),
        invoke=("get_workbook_context", {}),
    )
    _install_fake_sdk(monkeypatch, tmp_path, fake)

    # run_sdk_plan 在函数体内 from .excel_agent import _execute_tool / AGENT_TOOLS_WITH_DATA，
    # 因此替身必须挂在 excel_agent 模块属性上（每次调用都会重新绑定）。
    import server.app.excel_agent as excel_agent_module

    original_execute = excel_agent_module._execute_tool

    def _fake_execute(request, name, arguments):
        if name == "get_workbook_context":  # 未列入批准集的非终结工具
            return {"workbook": {"sheets": [{"name": "PRIVATE_SHEET_SENTINEL"}]}}
        return original_execute(request, name, arguments)

    fake_def = {
        "type": "function",
        "function": {
            "name": "get_workbook_context",
            "description": "probe",
            "parameters": {"type": "object", "properties": {}, "additionalProperties": False},
        },
    }
    monkeypatch.setattr(excel_agent_module, "_execute_tool", _fake_execute)
    monkeypatch.setattr(
        excel_agent_module, "AGENT_TOOLS_WITH_DATA",
        [*excel_agent_module.AGENT_TOOLS_WITH_DATA, fake_def],
    )
    asyncio.run(run_sdk_plan(_sdk_plan_request(), _connection()))
    assert fake.tool_result is not None
    assert fake.tool_result.get("is_error") is True
    assert "非终结工具结果不可用" in fake.tool_result["content"][0]["text"]
    assert "PRIVATE_SHEET_SENTINEL" not in fake.tool_result["content"][0]["text"]


def test_approved_non_terminal_allowlist_matches_projection_registry():
    """双保险契约：SDK 批准集必须是出站投影登记表的子集。"""
    assert _APPROVED_NON_TERMINAL_TOOLS <= set(excel_agent.NON_TERMINAL_TOOL_PROJECTIONS)
    assert "run_python_analysis" in _APPROVED_NON_TERMINAL_TOOLS


# ---------------------------------------------------------------------------
# 6. 缺陷①：被拒绝的参数仍进入消息与检查点（入参校验 ≠ 历史投影）
# ---------------------------------------------------------------------------
# 复现：withheld_python_arguments 只删顶层 code，其余参数（apiKey / 嵌套对象
# datasetHandles / 畸形 JSON）原样保留，即使入口判定 invalid_request，这些
# 敏感原值仍进入下一轮模型请求与持久检查点。


def _assert_no_unknown_arguments_anywhere(blob: str, unknowns: list[str], where: str) -> None:
    for value in unknowns:
        assert value not in blob, f"{where} 泄露被拒绝参数原值：{value}"


def test_rejected_extra_fields_never_enter_model_requests(monkeypatch):
    """缺陷①（模型出口）：携带 apiKey/path/command 的非法请求被入口拒绝后，
    这些未知字段原值不得进入任何一轮完整模型请求——历史投影必须采用可信侧
    最小白名单，而不是只删 code 键。"""
    _enable_python(monkeypatch)
    secrets = ["SK_API_KEY_REJECTED_SENTINEL_1", "C:/PRIVATE_PATH_SENTINEL_1", "cmd.exe"]

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {
                        "code": _CODE_SENTINEL,
                        "apiKey": secrets[0],
                        "path": secrets[1],
                        "command": secrets[2],
                    },
                    "py_reject",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(monkeypatch, responder)
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        _assert_no_unknown_arguments_anywhere(
            blob, secrets, f"第 {turn_index + 1} 轮模型请求"
        )
        assert _CODE_SENTINEL not in blob, f"第 {turn_index + 1} 轮模型请求携带原始代码"


def test_rejected_nested_handles_never_enter_model_requests(monkeypatch):
    """缺陷①（模型出口）：datasetHandles 内嵌对象/自由文本也被入口拒绝，
    投影不得保留畸形嵌套形状（不把模型自报字符串当已授权句柄）。"""
    _enable_python(monkeypatch)
    nested = {"__object__": "NESTED_HANDLE_OBJECT_SENTINEL"}

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {"code": _CODE_SENTINEL, "datasetHandles": [nested, "ok_handle"]},
                    "py_nested",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(monkeypatch, responder)
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert "NESTED_HANDLE_OBJECT_SENTINEL" not in blob, (
            f"第 {turn_index + 1} 轮模型请求泄露嵌套句柄对象"
        )
        assert _CODE_SENTINEL not in blob


def test_rejected_arguments_never_enter_sqlite(monkeypatch, tmp_path):
    """缺陷①（持久化出口）：被拒绝的 apiKey 与畸形句柄不进入 SQLite 账本全列。

    复刻 main.py 真实接线（register → save_planning_checkpoint → pause →
    原始行读回扫描），证明"被拒绝"不等于"可以原样落盘"。"""
    _enable_python(monkeypatch)
    secret = "SK_REJECTED_PERSIST_SENTINEL_7f"
    db_path = tmp_path / "planning-rejected.sqlite3"
    ledger = FolderOperationLedger(lambda: db_path)
    turn_id = "turn-rejected-privacy-1"
    ledger.register_planning_task(
        turn_id, "fingerprint-rejected", {"prompt": "synthetic"},
        approved_wall_time_seconds=3600,
    )
    checkpoints: list[dict[str, Any]] = []
    event_seq = {"n": 0}

    async def on_checkpoint(checkpoint: dict[str, Any]) -> None:
        checkpoints.append(copy.deepcopy(checkpoint))
        ledger.save_planning_checkpoint(turn_id, checkpoint, active_elapsed_seconds=1.0)

    async def on_event(event: dict[str, Any]) -> None:
        event_seq["n"] += 1
        ledger.append_planning_event(turn_id, event_seq["n"], event)

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {"code": _CODE_SENTINEL, "apiKey": secret, "datasetHandles": [{"x": secret}]},
                    "py_reject",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, responder, on_checkpoint=on_checkpoint, on_event=on_event
    )
    assert result.kind == "plan"
    assert checkpoints
    ledger.pause_planning_task(
        turn_id, checkpoints[-1], active_elapsed_seconds=1.0,
        reason="budget-test", error=None,
    )
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    task_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_tasks WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    event_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_task_events WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    connection.close()
    assert task_rows and event_rows
    for row in task_rows + event_rows:
        blob = json.dumps(row, ensure_ascii=False)
        assert secret not in blob, f"持久化行泄露被拒绝参数：{row.keys()}"
        assert _CODE_SENTINEL not in blob


def test_malformed_json_arguments_never_enter_model_requests(monkeypatch):
    """缺陷①：畸形 JSON 的 arguments 解析失败时，不得把原字符串塞回历史。"""
    _enable_python(monkeypatch)
    # 直接构造一个 arguments 是畸形 JSON 的 run_python_analysis 调用。
    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            payload = {
                "choices": [{
                    "message": {
                        "content": None,
                        "tool_calls": [{
                            "id": "py_malformed",
                            "type": "function",
                            "function": {
                                "name": "run_python_analysis",
                                "arguments": "{not valid json {{ MALFORMED_SENTINEL",
                            },
                        }],
                    }
                }]
            }
            return (payload, _usage())
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(monkeypatch, responder)
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert "MALFORMED_SENTINEL" not in blob, (
            f"第 {turn_index + 1} 轮模型请求泄露畸形 JSON 原值"
        )


def test_multi_tool_batch_non_python_tools_unchanged(monkeypatch):
    """缺陷①：多工具批次里非 Python 工具参数保持原样，只有 Python 调用投影。"""
    _enable_python(monkeypatch)
    secret = "SK_MULTI_TOOL_REJECTED_SENTINEL_3"

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                {
                    "choices": [{
                        "message": {
                            "content": None,
                            "tool_calls": [
                                {
                                    "id": "py_multi",
                                    "type": "function",
                                    "function": {
                                        "name": "run_python_analysis",
                                        "arguments": json.dumps(
                                            {"code": _CODE_SENTINEL, "apiKey": secret},
                                            ensure_ascii=False,
                                        ),
                                    },
                                },
                                {
                                    "id": "ctx_multi",
                                    "type": "function",
                                    "function": {
                                        "name": "get_workbook_context",
                                        "arguments": json.dumps({"x": "keep-me"}, ensure_ascii=False),
                                    },
                                },
                            ],
                        }
                    }]
                },
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(monkeypatch, responder)
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert secret not in blob, f"第 {turn_index + 1} 轮泄露被拒 apiKey"
        assert _CODE_SENTINEL not in blob
    # 非 Python 工具参数保持原样（不被投影误伤）。
    second_messages = client.message_snapshots[1]
    ctx_call = next(
        c for m in second_messages if m.get("role") == "assistant" and m.get("tool_calls")
        for c in m["tool_calls"] if c["function"]["name"] == "get_workbook_context"
    )
    assert json.loads(ctx_call["function"]["arguments"]) == {"x": "keep-me"}


# ---------------------------------------------------------------------------
# 7. 缺陷②：旧检查点仍携带原始代码（恢复现场未收口）
# ---------------------------------------------------------------------------


def _old_checkpoint_with_python_code() -> dict[str, Any]:
    """构造修复前版本的旧检查点形状：messages 与 pendingAssistantMessage
    均含原始 Python code（尚未执行 + 已完成工具历史两类现场）。"""
    assistant_with_python = {
        "role": "assistant",
        "content": None,
        "tool_calls": [{
            "id": "py_old",
            "type": "function",
            "function": {
                "name": "run_python_analysis",
                "arguments": json.dumps(
                    {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_SENTINEL]},
                    ensure_ascii=False,
                ),
            },
        }],
    }
    return {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "model_response_saved",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            assistant_with_python,
            {
                "role": "tool",
                "tool_call_id": "py_old",
                "name": "run_python_analysis",
                "content": json.dumps({
                    "status": "unavailable",
                    "reasonCategory": "runner_missing",
                    "message": _RUNNER_MISSING_MESSAGE,
                }, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {
            "turnIndex": 1,
            "elapsedTime": 1.0,
            "totalTokens": 0,
            "consecutiveErrors": 0,
        },
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": {
            "role": "assistant",
            "content": None,
            "tool_calls": [{
                "id": "py_pending",
                "type": "function",
                "function": {
                    "name": "run_python_analysis",
                    "arguments": json.dumps(
                        {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_SENTINEL]},
                        ensure_ascii=False,
                    ),
                },
            }],
        },
        "nextToolIndex": 0,
        "toolBatchState": {},
    }


def test_old_checkpoint_python_code_not_sent_or_executed(monkeypatch):
    """缺陷②：旧检查点里 messages 与 pendingAssistantMessage 的原始 code，
    恢复后既不得进入模型请求，也不得进入本轮原始执行参数通道——
    恢复轮应得到 code_not_preserved（不是 runner_missing），且明确要求重新提供。"""
    _enable_python(monkeypatch)
    checkpoint = _old_checkpoint_with_python_code()

    captured_exec_arguments: list[dict[str, Any]] = []
    # ``run_python_analysis`` 改走 offload 入口（同步容器运行不能占住事件循环），
    # 因此取证点挂在这条**当前唯一**的执行入口上，而不是 ``_execute_tool``。
    from server.app import python_coordinator as _pyc

    original_async = _pyc.execute_python_analysis_async

    async def _spy_async(request, arguments, *, cancel_check=None, on_event=None):
        captured_exec_arguments.append(copy.deepcopy(arguments))
        return await original_async(request, arguments, cancel_check=cancel_check, on_event=on_event)

    monkeypatch.setattr(_pyc, "execute_python_analysis_async", _spy_async)

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    # 恢复后的每一轮模型请求都不含原始代码。datasetHandles 收口为空列表
    # （未验证句柄不保留原值，句柄登记仍未实现）。
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _CODE_SENTINEL not in blob, f"恢复后第 {turn_index + 1} 轮携带旧检查点原始代码"
    # 恢复现场经投影后，执行入口拿到的不是原始代码（应是占位，协调器判 code_not_preserved）。
    assert captured_exec_arguments, "恢复轮应调用 run_python_analysis"
    for args in captured_exec_arguments:
        assert args.get("code") != _CODE_SENTINEL, "旧检查点原始代码进入本轮执行参数通道"
    # 真实恢复结果：code_not_preserved（明确要求重新提供），不是 runner_missing。
    # 旧历史里已有已完成工具的 runner_missing 回执，因此取**最后一个** tool 消息
    # （恢复轮重新执行 pending 调用产生的新回执）。
    resumed_messages = client.message_snapshots[0]
    tool_messages = [m for m in resumed_messages if m.get("role") == "tool"]
    receipt = json.loads(tool_messages[-1]["content"])
    assert receipt["reasonCategory"] == "code_not_preserved"


def test_old_checkpoint_pending_message_code_not_preserved(monkeypatch):
    """缺陷②：旧 pendingAssistantMessage 含原始 code 时，恢复后必须判定
    code_not_preserved，而不是把原始 code 当成本轮新响应去执行（那会得到
    runner_missing 的错误语义）。"""
    _enable_python(monkeypatch)
    checkpoint = _old_checkpoint_with_python_code()
    # 让 messages 里不含 Python（已完成工具历史现场为空），只有 pending 含 code。
    checkpoint["messages"] = [
        {"role": "system", "content": "sys"},
        {"role": "user", "content": "用户需求"},
    ]

    captured_exec_arguments: list[dict[str, Any]] = []
    # ``run_python_analysis`` 改走 offload 入口（同步容器运行不能占住事件循环），
    # 因此取证点挂在这条**当前唯一**的执行入口上，而不是 ``_execute_tool``。
    from server.app import python_coordinator as _pyc

    original_async = _pyc.execute_python_analysis_async

    async def _spy_async(request, arguments, *, cancel_check=None, on_event=None):
        captured_exec_arguments.append(copy.deepcopy(arguments))
        return await original_async(request, arguments, cancel_check=cancel_check, on_event=on_event)

    monkeypatch.setattr(_pyc, "execute_python_analysis_async", _spy_async)

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    for args in captured_exec_arguments:
        assert args.get("code") != _CODE_SENTINEL
    tool_message = next(
        m for m in client.message_snapshots[0] if m.get("role") == "tool"
    )
    receipt = json.loads(tool_message["content"])
    assert receipt["reasonCategory"] == "code_not_preserved"


def test_old_checkpoint_without_python_resumes_normally(monkeypatch):
    """缺陷②边界：不含 Python 的旧检查点恢复不受影响（不破坏既有恢复）。"""
    _enable_python(monkeypatch)
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "before_model",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 0, "elapsedTime": 0.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 0,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"


# ---------------------------------------------------------------------------
# 8. 缺陷③：执行入口异常仍原样回显（str(error) 出口）
# ---------------------------------------------------------------------------


def test_execution_entry_exception_fails_closed_in_loop(monkeypatch):
    """缺陷③（普通后端）：_execute_tool / 协调器抛含哨兵异常时，模型回执、
    last_error_message、事件、持久记录都不含原始异常文本。"""
    _enable_python(monkeypatch)
    error_with_sentinel = ValueError(f"EXEC_ENTRY_EXPLOSION_{_CODE_SENTINEL}")
    # 执行入口异常：取证/注入点挂在 offload 入口（run_python_analysis 当前唯一
    # 的执行分派缝）上。
    from server.app import python_coordinator as _pyc

    async def _explode_async(request, arguments, *, cancel_check=None, on_event=None):
        raise error_with_sentinel

    monkeypatch.setattr(_pyc, "execute_python_analysis_async", _explode_async)
    events: list[dict[str, Any]] = []

    async def on_event(event: dict[str, Any]) -> None:
        events.append(copy.deepcopy(event))

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload("run_python_analysis", {"code": _CODE_SENTINEL, "datasetHandles": []}, "py_exec"),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, responder, on_event=on_event
    )
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert "EXEC_ENTRY_EXPLOSION" not in blob, f"第 {turn_index + 1} 轮回显执行异常"
        assert _CODE_SENTINEL not in blob
    # 事件 trace 不含原始异常文本。
    for event in events:
        blob = json.dumps(event, ensure_ascii=False)
        assert "EXEC_ENTRY_EXPLOSION" not in blob, "事件泄露执行异常文本"


def test_execution_entry_exception_fails_closed_in_sdk(monkeypatch, tmp_path):
    """缺陷③（SDK 后端）：执行入口抛含哨兵异常时，SDK 工具结果不含原始异常。"""
    from server.app.claude_sdk_agent import run_sdk_plan

    _enable_python(monkeypatch)
    error_with_sentinel = ValueError(f"SDK_EXEC_EXPLOSION_{_CODE_SENTINEL}")
    # 执行入口异常：取证/注入点挂在 offload 入口（run_python_analysis 当前唯一
    # 的执行分派缝）上。
    from server.app import python_coordinator as _pyc

    async def _explode_async(request, arguments, *, cancel_check=None, on_event=None):
        raise error_with_sentinel

    monkeypatch.setattr(_pyc, "execute_python_analysis_async", _explode_async)
    fake = _FakeSdk(
        response=json.dumps({"kind": "answer", "message": "收到"}),
        invoke=("run_python_analysis", {"code": _CODE_SENTINEL, "datasetHandles": []}),
    )
    _install_fake_sdk(monkeypatch, tmp_path, fake)
    asyncio.run(run_sdk_plan(_sdk_plan_request(), _connection()))
    assert fake.tool_result is not None
    text = fake.tool_result["content"][0]["text"]
    assert "SDK_EXEC_EXPLOSION" not in text
    assert _CODE_SENTINEL not in text
    # 固定安全回执（非 str(error)、非原样 dict）。
    parsed = json.loads(text)
    assert parsed["status"] == "unavailable"
    assert parsed["reasonCategory"] == "execution_error"
    assert parsed["message"] == (
        "隔离 Python 分析未能完成，未发布可用结果；是否开始执行以执行确认记录为准。"
    )


def test_cancelled_error_still_propagates(monkeypatch):
    """缺陷③边界：asyncio.CancelledError 不得被无差别吞掉伪装成成功/固定回执。"""
    _enable_python(monkeypatch)
    from server.app import python_coordinator as _pyc

    async def _cancel_async(request, arguments, *, cancel_check=None, on_event=None):
        raise asyncio.CancelledError()

    monkeypatch.setattr(_pyc, "execute_python_analysis_async", _cancel_async)

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (_tool_payload("run_python_analysis", {"code": "print(1)"}, "py_cancel"), _usage())
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client = _FakeClient(responder)
    _install_client(monkeypatch, client)
    with pytest.raises(asyncio.CancelledError):
        asyncio.run(
            excel_agent.run_excel_agent(
                _request(), connection=_connection(), timeout=5,
                config=LoopConfig(max_wall_time_seconds=3600),
            )
        )


# ---------------------------------------------------------------------------
# 9. 未登记句柄收口：历史/检查点投影不得保留模型自报的原始 datasetHandles
# ---------------------------------------------------------------------------
# 句柄登记核对尚未实现，因此「形状合法」的字符串句柄（包括看起来像正常
# dataset:// 句柄的字符串）也不能证明是可信侧签发的——投影必须统一收口为
# 空列表（固定、不含原值的兼容表示），清楚表达「未验证」，不冒充已授权。

_HANDLE_LIKE_SENTINEL = "dataset://auth/VALID_LOOKING_4f9a"


def test_unverified_handles_projected_to_empty_list(monkeypatch):
    """形状合法的 dataset:// 句柄不得以原值保留进历史投影——投影为空列表。"""
    _enable_python(monkeypatch)
    projected = json.loads(python_analysis.withheld_python_arguments(
        json.dumps({
            "code": _CODE_SENTINEL,
            "datasetHandles": [_HANDLE_LIKE_SENTINEL, "dataset://another/ok"],
        }, ensure_ascii=False)
    ))
    assert projected["code"] == _CODE_WITHHELD_MARKER
    assert projected["datasetHandles"] == []
    assert _HANDLE_LIKE_SENTINEL not in json.dumps(projected, ensure_ascii=False)


def test_unverified_handles_never_enter_model_requests(monkeypatch):
    """形状合法的句柄原值不得进入任何一轮完整模型请求。"""
    _enable_python(monkeypatch)

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_LIKE_SENTINEL]},
                    "py_h",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(monkeypatch, responder)
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _HANDLE_LIKE_SENTINEL not in blob, (
            f"第 {turn_index + 1} 轮模型请求泄露未验证句柄原值"
        )
        assert _CODE_SENTINEL not in blob
    # 投影后的 assistant 工具调用参数：datasetHandles 是空列表，不是原值。
    second_messages = client.message_snapshots[1]
    assistant_message = next(
        m for m in second_messages
        if m.get("role") == "assistant" and m.get("tool_calls")
    )
    call = assistant_message["tool_calls"][0]
    arguments = json.loads(call["function"]["arguments"])
    assert arguments["datasetHandles"] == []


def test_unverified_handles_rejected_request_still_closed(monkeypatch):
    """因额外字段被拒的请求：形状合法的句柄同样不得保留原值（两道独立边界）。"""
    _enable_python(monkeypatch)
    secret = "SK_EXTRA_REJECTED_HANDLE_SENTINEL_9c"

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {
                        "code": _CODE_SENTINEL,
                        "datasetHandles": [_HANDLE_LIKE_SENTINEL],
                        "apiKey": secret,
                    },
                    "py_reject_h",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(monkeypatch, responder)
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _HANDLE_LIKE_SENTINEL not in blob, (
            f"第 {turn_index + 1} 轮泄露被拒请求里的句柄原值"
        )
        assert secret not in blob
        assert _CODE_SENTINEL not in blob


def test_unverified_handles_never_enter_sqlite(monkeypatch, tmp_path):
    """形状合法的句柄原值不进 SQLite 账本全列（真实 FolderOperationLedger）。"""
    _enable_python(monkeypatch)
    db_path = tmp_path / "planning-handles.sqlite3"
    ledger = FolderOperationLedger(lambda: db_path)
    turn_id = "turn-handles-privacy-1"
    ledger.register_planning_task(
        turn_id, "fingerprint-handles", {"prompt": "synthetic"},
        approved_wall_time_seconds=3600,
    )
    checkpoints: list[dict[str, Any]] = []
    event_seq = {"n": 0}

    async def on_checkpoint(checkpoint: dict[str, Any]) -> None:
        checkpoints.append(copy.deepcopy(checkpoint))
        ledger.save_planning_checkpoint(turn_id, checkpoint, active_elapsed_seconds=1.0)

    async def on_event(event: dict[str, Any]) -> None:
        event_seq["n"] += 1
        ledger.append_planning_event(turn_id, event_seq["n"], event)

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        if turn == 1:
            return (
                _tool_payload(
                    "run_python_analysis",
                    {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_LIKE_SENTINEL]},
                    "py_h_db",
                ),
                _usage(),
            )
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, responder, on_checkpoint=on_checkpoint, on_event=on_event
    )
    assert result.kind == "plan"
    assert checkpoints
    ledger.pause_planning_task(
        turn_id, checkpoints[-1], active_elapsed_seconds=1.0,
        reason="budget-test", error=None,
    )
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    task_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_tasks WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    event_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_task_events WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    connection.close()
    assert task_rows and event_rows
    for row in task_rows + event_rows:
        blob = json.dumps(row, ensure_ascii=False)
        assert _HANDLE_LIKE_SENTINEL not in blob, f"持久化行泄露未验证句柄原值：{row.keys()}"
        assert _CODE_SENTINEL not in blob


def test_unverified_handles_closed_in_resume_field(monkeypatch):
    """旧检查点 messages / pendingAssistantMessage 里的形状合法句柄原值，
    恢复后不得进入模型请求——投影统一收口为空列表。"""
    _enable_python(monkeypatch)
    assistant_with_handle = {
        "role": "assistant",
        "content": None,
        "tool_calls": [{
            "id": "py_old_h",
            "type": "function",
            "function": {
                "name": "run_python_analysis",
                "arguments": json.dumps(
                    {"code": _CODE_SENTINEL, "datasetHandles": [_HANDLE_LIKE_SENTINEL]},
                    ensure_ascii=False,
                ),
            },
        }],
    }
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "model_response_saved",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            assistant_with_handle,
            {
                "role": "tool",
                "tool_call_id": "py_old_h",
                "name": "run_python_analysis",
                "content": json.dumps({
                    "status": "unavailable",
                    "reasonCategory": "runner_missing",
                    "message": _RUNNER_MISSING_MESSAGE,
                }, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _HANDLE_LIKE_SENTINEL not in blob, (
            f"恢复后第 {turn_index + 1} 轮泄露旧检查点句柄原值"
        )
        assert _CODE_SENTINEL not in blob


# ---------------------------------------------------------------------------
# 10. 旧检查点 Python 错误上下文收口：工具回执 / loopState 错误与警告
# ---------------------------------------------------------------------------
# 修复前版本的检查点可能仍携带：旧 run_python_analysis 工具错误回执（自由文本
# error/message/stdout/traceback）、loopState.lastErrorMessage、lastWarning。
# 这些在恢复后会被重放进模型请求、公开结果与 SQLite——必须收口。

_ERR_SENTINEL = "PY_OLD_ERR_SENTINEL_bb11"
_WARN_SENTINEL = "PY_OLD_WARN_SENTINEL_cc22"
_STDOUT_SENTINEL = "PY_OLD_STDOUT_SENTINEL_dd33"


def _old_checkpoint_with_python_error_context() -> dict[str, Any]:
    """构造修复前形状的检查点：tool 回执、lastErrorMessage、lastWarning 三处哨兵。"""
    return {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_completed",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [{
                    "id": "py_err",
                    "type": "function",
                    "function": {
                        "name": "run_python_analysis",
                        "arguments": json.dumps(
                            {"code": _CODE_SENTINEL, "datasetHandles": []},
                            ensure_ascii=False,
                        ),
                    },
                }],
            },
            {
                "role": "tool",
                "tool_call_id": "py_err",
                "name": "run_python_analysis",
                "content": json.dumps({
                    "status": "unavailable",
                    "reasonCategory": "execution_error",
                    "message": _ERR_SENTINEL,
                    "stdout": _STDOUT_SENTINEL,
                    "traceback": _ERR_SENTINEL,
                }, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {
            "turnIndex": 2,
            "toolCallHistory": ["run_python_analysis"],
            "elapsedTime": 1.0,
            "totalTokens": 0,
            "consecutiveErrors": 1,
            "lastErrorMessage": f"execution failed: {_ERR_SENTINEL}",
            "lastWarning": f"warning: {_WARN_SENTINEL}",
        },
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 2,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }


def _assert_no_python_error_sentinels(blob: str, where: str) -> None:
    for sentinel in (_ERR_SENTINEL, _WARN_SENTINEL, _STDOUT_SENTINEL, _CODE_SENTINEL):
        assert sentinel not in blob, f"{where} 泄露旧 Python 错误上下文：{sentinel}"


def test_old_checkpoint_python_error_context_not_sent_or_persisted(monkeypatch, tmp_path):
    """旧检查点三处错误哨兵：恢复后不进入模型请求、公开结果、SQLite。"""
    _enable_python(monkeypatch)
    checkpoint = _old_checkpoint_with_python_error_context()
    db_path = tmp_path / "planning-olderr.sqlite3"
    ledger = FolderOperationLedger(lambda: db_path)
    turn_id = "turn-olderr-1"
    ledger.register_planning_task(
        turn_id, "fingerprint-olderr", {"prompt": "synthetic"},
        approved_wall_time_seconds=3600,
    )
    checkpoints: list[dict[str, Any]] = []
    events: list[dict[str, Any]] = []
    event_seq = {"n": 0}

    async def on_checkpoint(cp: dict[str, Any]) -> None:
        checkpoints.append(copy.deepcopy(cp))
        ledger.save_planning_checkpoint(turn_id, cp, active_elapsed_seconds=1.0)

    async def on_event(event: dict[str, Any]) -> None:
        events.append(copy.deepcopy(event))
        event_seq["n"] += 1
        ledger.append_planning_event(turn_id, event_seq["n"], event)

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint,
        on_checkpoint=on_checkpoint, on_event=on_event,
    )
    assert result.kind == "plan"
    # 1) 模型请求（完整每轮）
    for turn_index, snapshot in enumerate(client.message_snapshots):
        _assert_no_python_error_sentinels(
            json.dumps(snapshot, ensure_ascii=False),
            f"恢复后第 {turn_index + 1} 轮模型请求",
        )
    # 2) 事件
    for event in events:
        _assert_no_python_error_sentinels(
            json.dumps(event, ensure_ascii=False), "事件"
        )
    # 3) 公开结果
    _assert_no_python_error_sentinels(
        json.dumps(result.model_dump(mode="json"), ensure_ascii=False), "公开结果"
    )
    # 4) 新 SQLite 记录
    ledger.pause_planning_task(
        turn_id, checkpoints[-1], active_elapsed_seconds=1.0,
        reason="budget-test", error=None,
    )
    connection = sqlite3.connect(db_path)
    connection.row_factory = sqlite3.Row
    task_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_tasks WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    event_rows = [dict(r) for r in connection.execute(
        "SELECT * FROM planning_task_events WHERE turn_id=?", (turn_id,)
    ).fetchall()]
    connection.close()
    assert task_rows and event_rows
    for row in task_rows + event_rows:
        _assert_no_python_error_sentinels(
            json.dumps(row, ensure_ascii=False), f"持久化行 {row.keys()}"
        )


def test_old_checkpoint_python_error_context_in_pending(monkeypatch):
    """pending 工具现场：pendingAssistantMessage 里的 Python 调用在恢复后同样
    不得重放旧错误上下文，loopState 错误/警告同样收口。"""
    _enable_python(monkeypatch)
    checkpoint = _old_checkpoint_with_python_error_context()
    checkpoint["messages"] = [
        {"role": "system", "content": "sys"},
        {"role": "user", "content": "用户需求"},
    ]
    checkpoint["pendingAssistantMessage"] = {
        "role": "assistant",
        "content": None,
        "tool_calls": [{
            "id": "py_pending_err",
            "type": "function",
            "function": {
                "name": "run_python_analysis",
                "arguments": json.dumps(
                    {"code": _CODE_SENTINEL, "datasetHandles": []},
                    ensure_ascii=False,
                ),
            },
        }],
    }
    checkpoint["stage"] = "model_response_saved"
    checkpoint["nextToolIndex"] = 0

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        _assert_no_python_error_sentinels(
            json.dumps(snapshot, ensure_ascii=False),
            f"恢复后第 {turn_index + 1} 轮模型请求",
        )
    _assert_no_python_error_sentinels(
        json.dumps(result.model_dump(mode="json"), ensure_ascii=False), "公开结果"
    )


def test_old_checkpoint_mixed_tool_history_preserves_non_python(monkeypatch):
    """混合工具历史：非 Python 工具回执保持原样，Python 回执重建为固定回执。"""
    _enable_python(monkeypatch)
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_completed",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [
                    {
                        "id": "ctx1",
                        "type": "function",
                        "function": {"name": "get_workbook_context", "arguments": "{}"},
                    },
                    {
                        "id": "py_mix",
                        "type": "function",
                        "function": {
                            "name": "run_python_analysis",
                            "arguments": json.dumps(
                                {"code": _CODE_SENTINEL, "datasetHandles": []},
                                ensure_ascii=False,
                            ),
                        },
                    },
                ],
            },
            {
                "role": "tool",
                "tool_call_id": "ctx1",
                "name": "get_workbook_context",
                "content": json.dumps({"workbook": {"sheets": [{"name": "KEEP_THIS_SHEET"}]}}, ensure_ascii=False),
            },
            {
                "role": "tool",
                "tool_call_id": "py_mix",
                "name": "run_python_analysis",
                "content": json.dumps({"message": _ERR_SENTINEL, "stdout": _STDOUT_SENTINEL}, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _ERR_SENTINEL not in blob, f"第 {turn_index + 1} 轮泄露旧 Python 错误回执"
        assert _STDOUT_SENTINEL not in blob
        assert _CODE_SENTINEL not in blob
        # 非 Python 工具回执内容保持原样，不被误伤。
        assert "KEEP_THIS_SHEET" in blob


def test_old_checkpoint_tool_receipt_no_name_by_tool_call_id(monkeypatch):
    """无 name 字段、但可由 tool_call_id 确认 Python 归属的回执，重建为固定回执。"""
    _enable_python(monkeypatch)
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_completed",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [{
                    "id": "py_noname",
                    "type": "function",
                    "function": {
                        "name": "run_python_analysis",
                        "arguments": json.dumps(
                            {"code": _CODE_SENTINEL, "datasetHandles": []},
                            ensure_ascii=False,
                        ),
                    },
                }],
            },
            {
                "role": "tool",
                "tool_call_id": "py_noname",
                "content": json.dumps({"message": _ERR_SENTINEL, "stdout": _STDOUT_SENTINEL}, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _ERR_SENTINEL not in blob, f"第 {turn_index + 1} 轮泄露无 name 的旧回执"
        assert _STDOUT_SENTINEL not in blob
        assert _CODE_SENTINEL not in blob


def test_old_checkpoint_unattributable_tool_receipt_refuses_resume(monkeypatch):
    """混合历史中无法确定归属的自由文本回执：拒绝恢复，给出固定提示。"""
    _enable_python(monkeypatch)
    # 一个 tool 消息：无 name、tool_call_id 不对应任何 assistant 调用，
    # 内容是无法可靠确认归属的自由文本（可能含 Python 敏感输出）。
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_completed",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [{
                    "id": "py_attr",
                    "type": "function",
                    "function": {
                        "name": "run_python_analysis",
                        "arguments": json.dumps(
                            {"code": _CODE_SENTINEL, "datasetHandles": []},
                            ensure_ascii=False,
                        ),
                    },
                }],
            },
            {
                "role": "tool",
                "tool_call_id": "orphan_unknown",
                "content": f"自由文本 {_ERR_SENTINEL} 无法确认归属",
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    with pytest.raises(ValueError, match="重新生成预览"):
        _run_agent_with_client(
            monkeypatch, resumed_responder, resume_checkpoint=checkpoint
        )


def test_old_checkpoint_without_python_error_context_resumes_normally(monkeypatch):
    """不含 Python 的合法检查点（含普通 lastErrorMessage）不受影响。"""
    _enable_python(monkeypatch)
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "before_model",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
        ],
        "toolCache": {},
        "loopState": {
            "turnIndex": 0, "elapsedTime": 0.0, "totalTokens": 0,
            "consecutiveErrors": 1, "lastErrorMessage": "普通校验错误，无 Python",
        },
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 0,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"


# ---------------------------------------------------------------------------
# 11. 两个恢复分支补修：pending-only loopState 漏清理 + 带 name 的孤立回执
# ---------------------------------------------------------------------------


def _pending_only_python_checkpoint() -> dict[str, Any]:
    """构造「messages 只有 system/user、Python 调用只在 pendingAssistantMessage」的检查点。

    loopState 里 lastErrorMessage / lastWarning 各含不同哨兵，模拟修复前版本
    在 Python 执行异常时写下的自由文本。这正是漏洞①：判定 Python 相关状态只查
    messages，看不到 pending 里的 Python 调用，导致这两个自由文本字段漏清理、
    被首次 model_response_saved 检查点写进 SQLite。
    """
    return {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "model_response_saved",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
        ],
        "toolCache": {},
        "loopState": {
            "turnIndex": 1,
            "toolCallHistory": ["run_python_analysis"],
            "elapsedTime": 1.0,
            "totalTokens": 0,
            "consecutiveErrors": 1,
            "lastErrorMessage": f"execution failed: {_ERR_SENTINEL}",
            "lastWarning": f"path loop warning: {_WARN_SENTINEL}",
        },
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": {
            "role": "assistant",
            "content": None,
            "tool_calls": [{
                "id": "py_pending_only",
                "type": "function",
                "function": {
                    "name": "run_python_analysis",
                    "arguments": json.dumps(
                        {"code": _CODE_SENTINEL, "datasetHandles": []},
                        ensure_ascii=False,
                    ),
                },
            }],
        },
        "nextToolIndex": 0,
        "toolBatchState": {},
    }


def test_pending_only_python_loop_state_never_persisted(monkeypatch, tmp_path):
    """漏洞①：messages 无 Python、仅 pending 含 Python 时，loopState 的
    lastErrorMessage / lastWarning 不得进入任何一次保存的检查点（含很快被
    后续检查点覆盖的短暂记录）。每次 save_planning_checkpoint 后立即读 SQLite
    原始行取证，不能只检查最终记录。"""
    _enable_python(monkeypatch)
    checkpoint = _pending_only_python_checkpoint()
    # 先证明原检查点确实含哨兵（反向锚点）。
    original_blob = json.dumps(checkpoint, ensure_ascii=False)
    assert _ERR_SENTINEL in original_blob and _WARN_SENTINEL in original_blob

    db_path = tmp_path / "planning-pending-only.sqlite3"
    ledger = FolderOperationLedger(lambda: db_path)
    turn_id = "turn-pending-only-1"
    ledger.register_planning_task(
        turn_id, "fingerprint-pending-only", {"prompt": "synthetic"},
        approved_wall_time_seconds=3600,
    )
    saved_rows: list[list[dict[str, Any]]] = []

    async def on_checkpoint(cp: dict[str, Any]) -> None:
        ledger.save_planning_checkpoint(turn_id, cp, active_elapsed_seconds=1.0)
        connection = sqlite3.connect(db_path)
        connection.row_factory = sqlite3.Row
        rows = [dict(r) for r in connection.execute(
            "SELECT * FROM planning_tasks WHERE turn_id=?", (turn_id,)
        ).fetchall()]
        connection.close()
        saved_rows.append(rows)

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint,
        on_checkpoint=on_checkpoint,
    )
    assert result.kind == "plan"
    assert saved_rows, "真实恢复链路必须产生至少一次检查点保存"
    for boundary_index, rows in enumerate(saved_rows):
        for row in rows:
            blob = json.dumps(row, ensure_ascii=False)
            assert _ERR_SENTINEL not in blob, (
                f"第 {boundary_index + 1} 次保存边界泄露 lastErrorMessage 哨兵"
            )
            assert _WARN_SENTINEL not in blob, (
                f"第 {boundary_index + 1} 次保存边界泄露 lastWarning 哨兵"
            )
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _ERR_SENTINEL not in blob and _WARN_SENTINEL not in blob, (
            f"恢复后第 {turn_index + 1} 轮模型请求泄露 loopState 哨兵"
        )
    _assert_no_python_error_sentinels(
        json.dumps(result.model_dump(mode="json"), ensure_ascii=False), "公开结果"
    )


def _named_orphan_receipt_checkpoint() -> dict[str, Any]:
    """构造「带合法非 Python 工具名、但 tool_call_id 孤立」的回执绕过场景。

    漏洞②：恢复投影只拒绝「无 name 且 tool_call_id 对不上」的回执；这里回执
    填了 name="get_workbook_context"，但 tool_call_id 没有对应的 assistant 调用，
    其敏感 content 被原样保留。
    """
    return {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_completed",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [{
                    "id": "ctx_real",
                    "type": "function",
                    "function": {"name": "get_workbook_context", "arguments": "{}"},
                }],
            },
            {
                "role": "tool",
                "tool_call_id": "orphan_ctx_fake",
                "name": "get_workbook_context",
                "content": json.dumps({"workbook": {"sheets": [{"name": _ERR_SENTINEL}]}}, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }


def test_named_orphan_receipt_refuses_resume(monkeypatch):
    """漏洞②：带合法非 Python 工具名、但 tool_call_id 孤立的回执，必须拒绝恢复，
    不能因 name 非空就放行其敏感 content。"""
    _enable_python(monkeypatch)
    checkpoint = _named_orphan_receipt_checkpoint()

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    with pytest.raises(ValueError, match="重新生成预览"):
        _run_agent_with_client(
            monkeypatch, resumed_responder, resume_checkpoint=checkpoint
        )


def test_named_orphan_receipt_never_persisted_or_sent(monkeypatch, tmp_path):
    """漏洞②：拒绝恢复时不得发起模型请求，也不得产生含敏感 content 的新检查点。"""
    _enable_python(monkeypatch)
    checkpoint = _named_orphan_receipt_checkpoint()
    db_path = tmp_path / "planning-named-orphan.sqlite3"
    ledger = FolderOperationLedger(lambda: db_path)
    turn_id = "turn-named-orphan-1"
    ledger.register_planning_task(
        turn_id, "fingerprint-named-orphan", {"prompt": "synthetic"},
        approved_wall_time_seconds=3600,
    )
    saved_rows: list[list[dict[str, Any]]] = []

    async def on_checkpoint(cp: dict[str, Any]) -> None:
        ledger.save_planning_checkpoint(turn_id, cp, active_elapsed_seconds=1.0)
        connection = sqlite3.connect(db_path)
        connection.row_factory = sqlite3.Row
        rows = [dict(r) for r in connection.execute(
            "SELECT * FROM planning_tasks WHERE turn_id=?", (turn_id,)
        ).fetchall()]
        connection.close()
        saved_rows.append(rows)

    client = _FakeClient(lambda _: (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage()))
    _install_client(monkeypatch, client)
    with pytest.raises(ValueError, match="重新生成预览"):
        asyncio.run(
            excel_agent.run_excel_agent(
                _request(), connection=_connection(), timeout=5,
                config=LoopConfig(max_wall_time_seconds=3600),
                resume_checkpoint=checkpoint,
                on_checkpoint=on_checkpoint,
            )
        )
    assert client.call_count == 0, "拒绝恢复不得发起模型请求"
    for rows in saved_rows:
        for row in rows:
            blob = json.dumps(row, ensure_ascii=False)
            assert _ERR_SENTINEL not in blob, "拒绝恢复仍把敏感 content 写进 SQLite"


def test_named_receipt_name_mismatch_refuses_resume(monkeypatch):
    """漏洞②：tool_call_id 有对应调用、但回执 name 与调用名不一致，必须拒绝。"""
    _enable_python(monkeypatch)
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_completed",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [{
                    "id": "ctx_mismatch",
                    "type": "function",
                    "function": {"name": "get_workbook_context", "arguments": "{}"},
                }],
            },
            {
                "role": "tool",
                "tool_call_id": "ctx_mismatch",
                "name": "read_range",
                "content": json.dumps({"values": [[_ERR_SENTINEL]]}, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    with pytest.raises(ValueError, match="重新生成预览"):
        _run_agent_with_client(
            monkeypatch, resumed_responder, resume_checkpoint=checkpoint
        )


def test_named_python_receipt_reconstructed_fixed(monkeypatch):
    """合法 Python 回执（name 与调用一致）仍用固定内容重建，不回显旧 content。"""
    _enable_python(monkeypatch)
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "tool_completed",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [{
                    "id": "py_named",
                    "type": "function",
                    "function": {
                        "name": "run_python_analysis",
                        "arguments": json.dumps(
                            {"code": _CODE_SENTINEL, "datasetHandles": []},
                            ensure_ascii=False,
                        ),
                    },
                }],
            },
            {
                "role": "tool",
                "tool_call_id": "py_named",
                "name": "run_python_analysis",
                "content": json.dumps({"message": _ERR_SENTINEL, "stdout": _STDOUT_SENTINEL}, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": None,
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _ERR_SENTINEL not in blob, f"第 {turn_index + 1} 轮泄露旧 Python 回执 content"
        assert _STDOUT_SENTINEL not in blob
        assert _CODE_SENTINEL not in blob


def test_mirror_image_in_messages_and_pending_not_rejected(monkeypatch):
    """漏洞②边界：正常检查点在 messages 与 pending 保存同一调用的镜像，
    不得把一致的镜像误判为冲突。"""
    _enable_python(monkeypatch)
    assistant_py = {
        "role": "assistant",
        "content": None,
        "tool_calls": [{
            "id": "py_mirror",
            "type": "function",
            "function": {
                "name": "run_python_analysis",
                "arguments": json.dumps(
                    {"code": _CODE_SENTINEL, "datasetHandles": []},
                    ensure_ascii=False,
                ),
            },
        }],
    }
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "model_response_saved",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            assistant_py,
            {
                "role": "tool",
                "tool_call_id": "py_mirror",
                "name": "run_python_analysis",
                "content": json.dumps({"message": _ERR_SENTINEL}, ensure_ascii=False),
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": copy.deepcopy(assistant_py),
        "nextToolIndex": 1,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    client, result = _run_agent_with_client(
        monkeypatch, resumed_responder, resume_checkpoint=checkpoint
    )
    assert result.kind == "plan"
    for turn_index, snapshot in enumerate(client.message_snapshots):
        blob = json.dumps(snapshot, ensure_ascii=False)
        assert _ERR_SENTINEL not in blob, f"第 {turn_index + 1} 轮泄露镜像回执 content"
        assert _CODE_SENTINEL not in blob


def test_true_conflicting_association_refuses_resume(monkeypatch):
    """漏洞②边界：真正不一致的映射（同一 tool_call_id 在 messages 与 pending
    被映射到不同工具名）必须拒绝，不能静默选一个。"""
    _enable_python(monkeypatch)
    checkpoint = {
        "schemaVersion": 2,
        "kind": "planning",
        "stage": "model_response_saved",
        "messages": [
            {"role": "system", "content": "sys"},
            {"role": "user", "content": "用户需求"},
            {
                "role": "assistant",
                "content": None,
                "tool_calls": [{
                    "id": "py_conflict",
                    "type": "function",
                    "function": {
                        "name": "run_python_analysis",
                        "arguments": json.dumps(
                            {"code": _CODE_SENTINEL, "datasetHandles": []},
                            ensure_ascii=False,
                        ),
                    },
                }],
            },
        ],
        "toolCache": {},
        "loopState": {"turnIndex": 1, "elapsedTime": 1.0, "totalTokens": 0, "consecutiveErrors": 0},
        "toolMode": "full",
        "repairingSubmission": False,
        "selectedActionTypes": [],
        "nextTurnIndex": 1,
        "pendingAssistantMessage": {
            "role": "assistant",
            "content": None,
            "tool_calls": [{
                "id": "py_conflict",
                "type": "function",
                "function": {"name": "read_range", "arguments": "{}"},
            }],
        },
        "nextToolIndex": 0,
        "toolBatchState": {},
    }

    def resumed_responder(turn: int) -> tuple[dict[str, Any], dict[str, int]]:
        return (_tool_payload("submit_plan", {"plan": _PLAN}, "plan2"), _usage())

    with pytest.raises(ValueError, match="重新生成预览"):
        _run_agent_with_client(
            monkeypatch, resumed_responder, resume_checkpoint=checkpoint
        )

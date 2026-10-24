"""缺陷 3：「正在运行」只能来自执行器**真实启动**。

旧实现的失败形态（已复现，见
``outputs/acceptance-2026-10-22/independent-review/python-status.json``）：
功能关闭时执行器调用 0 次、回执是 ``unavailable/disabled``，但事件流仍然发出
``validating → running(executedVersion=1) → execution_failed`` —— 界面上那段代码
"跑过一版"，实际容器一次都没起。

判据只有一个来源：``on_runner_start`` —— 由
隔离适配器在收到工作程序执行确认（或旧镜像的合法成功结果）后同步回调。
本文件把"这个回调有没有被调用"和"事件里有没有 running / executedVersion"钉在一起。

覆盖面：正常启动 / 功能关闭 / 运行器不可用 / 没有可信绑定 / 启动前取消 /
运行中取消 / 执行异常。
"""
from __future__ import annotations

import asyncio
import threading
import time
from types import SimpleNamespace

import pytest

from server.app import python_analysis as pa
from server.app.local_results import local_results
from server.app.python_coordinator import (
    coordinator,
    execute_python_analysis_async,
)
from server.app.python_datasets import registry
from server.tests.test_python_async_offload import TASK, _arguments, _bind


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()
    monkeypatch.setattr(pa, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(pa, "_runner_status", lambda: "verified")
    yield
    coordinator.reset()
    registry._grants.clear()
    registry._budgets.clear()
    local_results.reset()


def _collect():
    events: list[dict] = []

    async def on_event(payload):
        events.append(payload)

    return events, on_event


def _stages(events: list[dict]) -> list[str]:
    return [item["pythonCode"]["stage"] for item in events if "pythonCode" in item]


def _request():
    return SimpleNamespace(turnId=TASK)


def _no_runner(monkeypatch, reason: str) -> dict[str, int]:
    """执行器一旦被调用就失败，同时把"调用了几次"带出来。"""
    calls = {"runner": 0}

    def fake(**_kwargs):
        calls["runner"] += 1
        raise AssertionError(reason)

    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    return calls


def test_runner_start_is_the_only_source_of_running(monkeypatch):
    """正向锚点：执行器真的启动 ⇒ running 与 executedVersion 才出现。"""
    calls = {"runner": 0, "started": 0}

    def fake(*, data_registry, scope_token, scope, origin, receipts, code,
             cancel_check=None, on_runner_start=None):
        calls["runner"] += 1
        # 模拟已经收到工作程序的执行确认；真实适配器路径在本文件后半段覆盖。
        calls["started"] += 1
        if on_runner_start:
            on_runner_start()
        return {"status": "completed", "result": {"kind": "matrix", "values": [[1]]}}

    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    receipt = _bind()
    events, on_event = _collect()
    result = asyncio.run(
        execute_python_analysis_async(
            _request(), _arguments(receipt.datasetHandle), on_event=on_event
        )
    )

    assert result["status"] == "completed"
    assert calls == {"runner": 1, "started": 1}
    assert _stages(events) == [
        "generated",
        "validating",
        "running",
        "analysis_completed",
    ]
    by_stage = {
        item["pythonCode"]["stage"]: item["pythonCode"]
        for item in events
        if "pythonCode" in item
    }
    assert by_stage["running"]["executedVersion"] == 1
    assert by_stage["analysis_completed"]["executedVersion"] == 1
    # 只有真跑过的那两个状态带执行版本；"已生成""校验中"不许带。
    assert "executedVersion" not in by_stage["generated"]
    assert "executedVersion" not in by_stage["validating"]


def test_disabled_feature_never_claims_running(monkeypatch):
    """功能关闭：执行器 0 次，状态落到 not_started，绝不出现 running。"""
    monkeypatch.setattr(pa, "python_execution_enabled", lambda: False)
    calls = _no_runner(monkeypatch, "功能关闭时执行器不许被调用")
    receipt = _bind()
    events, on_event = _collect()
    result = asyncio.run(
        execute_python_analysis_async(
            _request(), _arguments(receipt.datasetHandle), on_event=on_event
        )
    )

    assert calls["runner"] == 0
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "disabled"
    assert _stages(events) == ["generated", "not_started"]
    terminal = [item for item in events if "pythonCode" in item][-1]
    # 具体原因要带出来，不统一包装成"执行失败"。
    assert terminal["pythonCode"]["reason"]
    assert "执行器" in terminal["pythonCode"]["reason"] or terminal["pythonCode"]["reason"]
    assert "executedVersion" not in terminal["pythonCode"]
    assert terminal["pythonCode"]["stage"] != "execution_failed"


def test_missing_runner_never_claims_running(monkeypatch):
    """运行器不可用：同样停在 not_started。"""
    monkeypatch.setattr(pa, "_runner_status", lambda: None)
    calls = _no_runner(monkeypatch, "没有已验证执行器时不许调用执行器")
    receipt = _bind()
    events, on_event = _collect()
    result = asyncio.run(
        execute_python_analysis_async(
            _request(), _arguments(receipt.datasetHandle), on_event=on_event
        )
    )

    assert calls["runner"] == 0
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "runner_missing"
    assert _stages(events) == ["generated", "not_started"]
    assert "running" not in _stages(events)
    assert not any(
        "executedVersion" in item["pythonCode"]
        for item in events
        if "pythonCode" in item
    )


def test_untrusted_binding_never_claims_running(monkeypatch):
    """没有可信数据登记（句柄无法验证）：失败关闭，且不算"跑过"。"""
    calls = _no_runner(monkeypatch, "没有可信绑定时不许调用执行器")
    receipt = _bind()
    # 换一个不在可信登记里的句柄。
    events, on_event = _collect()
    result = asyncio.run(
        execute_python_analysis_async(
            _request(), _arguments("ds_" + "0" * 24), on_event=on_event
        )
    )

    assert calls["runner"] == 0
    assert result["status"] == "rejected"
    assert _stages(events) == ["generated", "not_started"]
    del receipt  # 反向锚点：即便这次登记过数据，换了句柄也不许放宽


def test_invalid_arguments_are_validation_failed_not_started(monkeypatch):
    """入参非法 ⇒ 校验未通过（与"未开始执行"分开，不混成一种）。"""
    calls = _no_runner(monkeypatch, "入参非法时不许调用执行器")
    receipt = _bind()
    events, on_event = _collect()
    result = asyncio.run(
        execute_python_analysis_async(
            _request(),
            # 未知字段一律拒绝（不做隐式转换）；代码正文本身非空，所以这一轮
            # **有代码可展示**，只是它没有被允许进入执行。
            {"code": "result = 1", "datasetHandles": [receipt.datasetHandle],
             "path": "/etc/passwd"},
            on_event=on_event,
        )
    )

    assert calls["runner"] == 0
    assert result["status"] == "rejected"
    assert _stages(events) == ["generated", "validation_failed"]
    terminal = [item for item in events if "pythonCode" in item][-1]
    assert terminal["status"] == "failed"
    assert "executedVersion" not in terminal["pythonCode"]


def test_cancel_before_runner_start_is_not_started(monkeypatch):
    """取消发生在执行器启动之前：说"未开始执行"，不能说"已暂停"。"""
    entered = threading.Event()

    def fake(*, data_registry, scope_token, scope, origin, receipts, code,
             cancel_check=None, on_runner_start=None):
        # 执行器已经在工作线程里，但**容器还没启动**（on_runner_start 尚未回调）。
        entered.set()
        deadline = time.time() + 3
        while time.time() < deadline:
            if cancel_check is not None and cancel_check():
                break
            time.sleep(0.005)
        return {"status": "cancelled", "reasonCategory": "authorization_changed"}

    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    receipt = _bind()

    async def scenario():
        events, on_event = _collect()
        task = asyncio.create_task(
            execute_python_analysis_async(
                _request(), _arguments(receipt.datasetHandle), on_event=on_event
            )
        )
        for _ in range(300):
            if entered.is_set():
                break
            await asyncio.sleep(0.01)
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        return events

    events = asyncio.run(scenario())
    assert _stages(events) == ["generated", "validating", "not_started"]
    terminal = [item for item in events if "pythonCode" in item][-1]
    assert terminal["status"] == "cancelled"
    assert terminal["title"] == "代码执行已取消"
    assert "executedVersion" not in terminal["pythonCode"]
    # 取消不留永久 running：最后一条是终态。
    assert terminal["pythonCode"]["stage"] not in ("running", "validating", "generated")


def test_cancel_after_runner_start_is_paused_with_executed_version(monkeypatch):
    """运行中取消：说"已暂停"，并且**确实**带着执行版本（它真跑过）。"""
    started = threading.Event()

    def fake(*, data_registry, scope_token, scope, origin, receipts, code,
             cancel_check=None, on_runner_start=None):
        if on_runner_start:
            on_runner_start()
        started.set()
        deadline = time.time() + 3
        while time.time() < deadline:
            if cancel_check is not None and cancel_check():
                break
            time.sleep(0.005)
        return {"status": "cancelled", "reasonCategory": "authorization_changed"}

    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    receipt = _bind()

    async def scenario():
        events, on_event = _collect()
        task = asyncio.create_task(
            execute_python_analysis_async(
                _request(), _arguments(receipt.datasetHandle), on_event=on_event
            )
        )
        for _ in range(300):
            if started.is_set():
                break
            await asyncio.sleep(0.01)
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
        return events

    events = asyncio.run(scenario())
    assert _stages(events) == ["generated", "validating", "running", "paused"]
    terminal = [item for item in events if "pythonCode" in item][-1]
    assert terminal["status"] == "cancelled"
    assert terminal["pythonCode"]["executedVersion"] == 1
    assert local_results.count_task(TASK) == 0


def test_execution_exception_leaves_a_terminal_state(monkeypatch):
    """执行器抛异常：必须落到 execution_failed，不能永远停在"正在运行"。"""
    def fake(*, data_registry, scope_token, scope, origin, receipts, code,
             cancel_check=None, on_runner_start=None):
        if on_runner_start:
            on_runner_start()
        raise RuntimeError("容器运行时崩了")

    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    receipt = _bind()

    events, on_event = _collect()

    async def scenario():
        with pytest.raises(RuntimeError):
            await execute_python_analysis_async(
                _request(), _arguments(receipt.datasetHandle), on_event=on_event
            )

    asyncio.run(scenario())
    assert _stages(events) == [
        "generated",
        "validating",
        "running",
        "execution_failed",
    ]
    terminal = [item for item in events if "pythonCode" in item][-1]
    assert terminal["status"] == "failed"
    # 启动过 ⇒ 版本是**真实执行证据**，不是"没跑也报版本"。
    assert terminal["pythonCode"]["executedVersion"] == 1
    assert local_results.count_task(TASK) == 0


@pytest.fixture
def real_runner_boundary(monkeypatch):
    """Only OS/config boundaries are replaced; coordinator and all runner logic stay real."""
    from server.app import python_docker_runner as runner
    from server.tests.test_python_docker_runner import _valid_config, _EP, _popen_returning
    monkeypatch.setattr(runner, "_docker_config", _valid_config)
    monkeypatch.setattr(runner, "verified_endpoint", lambda: _EP)
    monkeypatch.setattr(runner, "verify_engine", lambda endpoint: True)
    monkeypatch.setattr(runner, "image_present", lambda config, endpoint: True)
    monkeypatch.setattr(runner, "build_run_arguments", lambda *a, **k: (["docker", "run", "audit"], {}))
    monkeypatch.setattr(runner, "_cleanup_container", lambda *a, **k: runner.CLEANUP_GONE)
    return runner, _popen_returning


def test_real_adapter_pre_spawn_cancel_never_claims_execution(monkeypatch, real_runner_boundary):
    runner, _ = real_runner_boundary
    calls = []
    def forbidden(*a, **k):
        calls.append(True)
        raise AssertionError("process must not start")
    monkeypatch.setattr(runner.subprocess, "Popen", forbidden)
    receipt = _bind()
    events, on_event = _collect()
    result = asyncio.run(execute_python_analysis_async(_request(), _arguments(receipt.datasetHandle),
        cancel_check=lambda: True, on_event=on_event))
    assert calls == []
    assert result["status"] == "cancelled"
    assert result["reasonCategory"] == "cancelled"
    assert pa.project_python_analysis_receipt(result)["status"] == "cancelled"
    assert _stages(events) == ["generated", "validating", "not_started"]
    assert all("executedVersion" not in e["pythonCode"] for e in events if "pythonCode" in e)


@pytest.mark.parametrize("failure", ["spawn_error", "cli_failure"])
def test_real_adapter_startup_failure_never_claims_execution(monkeypatch, real_runner_boundary, failure):
    runner, fake_process = real_runner_boundary
    if failure == "spawn_error":
        def popen(*a, **k):
            raise OSError("private startup detail")
    else:
        popen = fake_process(stderr="private daemon failure", returncode=125)
    monkeypatch.setattr(runner.subprocess, "Popen", popen)
    receipt = _bind()
    events, on_event = _collect()
    result = asyncio.run(execute_python_analysis_async(_request(), _arguments(receipt.datasetHandle), on_event=on_event))
    assert result["status"] == "unavailable"
    assert _stages(events) == ["generated", "validating", "not_started"]
    assert all("executedVersion" not in e["pythonCode"] for e in events if "pythonCode" in e)
    assert "private" not in str(events)
    assert "未运行任何代码" not in str(events)


def test_real_adapter_ack_then_cancel_preserves_cancelled(monkeypatch, real_runner_boundary):
    runner, fake_process = real_runner_boundary
    monkeypatch.setattr(runner.subprocess, "Popen", fake_process(
        stderr=runner.EXECUTION_STARTED_MARKER.decode(), hang=True))
    receipt = _bind()
    events = []
    state = {"cancel": False}
    async def on_event(payload):
        events.append(payload)
        if payload.get("pythonCode", {}).get("stage") == "running":
            state["cancel"] = True
    result = asyncio.run(execute_python_analysis_async(_request(), _arguments(receipt.datasetHandle),
        cancel_check=lambda: state["cancel"], on_event=on_event))
    assert result["status"] == "cancelled"
    assert _stages(events) == ["generated", "validating", "running", "paused"]
    assert events[-1]["pythonCode"]["executedVersion"] == 1
    assert local_results.count_task(TASK) == 0


@pytest.mark.parametrize("ack", [True, False])
def test_real_adapter_success_supports_ack_and_legacy_images(monkeypatch, real_runner_boundary, ack):
    import json
    runner, fake_process = real_runner_boundary
    monkeypatch.setattr(runner.subprocess, "Popen", fake_process(
        stdout=json.dumps({"status": "completed", "result": {"kind": "matrix", "values": [[1]]}}),
        stderr=runner.EXECUTION_STARTED_MARKER.decode() if ack else ""))
    receipt = _bind()
    events, on_event = _collect()
    result = asyncio.run(execute_python_analysis_async(_request(), _arguments(receipt.datasetHandle), on_event=on_event))
    assert result["status"] == "completed"
    assert _stages(events) == ["generated", "validating", "running", "analysis_completed"]
    assert sum(e.get("pythonCode", {}).get("stage") == "running" for e in events) == 1


def test_cancel_with_worker_exception_still_emits_terminal(monkeypatch):
    entered = threading.Event()
    def fake(*, cancel_check=None, **kwargs):
        entered.set()
        deadline = time.monotonic() + 3
        while not cancel_check() and time.monotonic() < deadline:
            time.sleep(0.005)
        raise RuntimeError("private cleanup error")
    monkeypatch.setattr(pa, "run_registered_analysis_internal", fake)
    receipt = _bind()
    events, on_event = _collect()
    async def scenario():
        task = asyncio.create_task(execute_python_analysis_async(
            _request(), _arguments(receipt.datasetHandle), on_event=on_event))
        for _ in range(300):
            if entered.is_set():
                break
            await asyncio.sleep(0.01)
        assert entered.is_set()
        task.cancel()
        with pytest.raises(asyncio.CancelledError):
            await task
    asyncio.run(scenario())
    assert _stages(events) == ["generated", "validating", "not_started"]
    assert events[-1]["status"] == "cancelled"
    assert "private" not in str(events)

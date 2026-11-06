"""验证 Python 分析完成事件能通过 on_event 回调投递到 SSE 流。"""
import asyncio
import pytest

from server.app.python_coordinator import execute_python_analysis_async
from server.app.python_coordinator import coordinator
from server.app.python_datasets import registry
from server.app.local_results import local_results


# 借用 test_python_coordinator 的辅助函数
from server.tests.test_python_coordinator import _scope, _bind


def test_python_card_survives_real_http_sse_serialization(monkeypatch, tmp_path):
    from fastapi.testclient import TestClient
    from server.app import main
    from server.app.models import AnswerResponse
    from server.app.task_ledger import FolderOperationLedger
    from server.tests.test_main import _STREAM_PAYLOAD, _parse_sse

    # 线上形状：完整性事实是**三个**字段（complete / incomplete / truncated+truncation），
    # 不是一个布尔 —— 卡片与模型回执都读它们，缺一项就只剩"猜原因"。
    card = dict(resultRef="res_" + "a" * 32, title="Python 分析结果", headers=["列1"],
                rowCount=2, columnCount=1, sourceSheets=["Sheet1"], complete=True,
                incomplete=[], truncated=False, truncation=None)

    async def plan(request, *, on_event, **kwargs):
        await on_event(dict(phase="planning", stepId="python.completed", status="completed",
                            title="隔离 Python 分析完成", pythonResult=card))
        return AnswerResponse(message="分析完成", provider="model")

    monkeypatch.setattr(main, "create_plan", plan)
    monkeypatch.setattr(main, "_folder_operation_ledger",
                        FolderOperationLedger(path_provider=lambda: tmp_path / "tasks.sqlite3"))
    response = TestClient(main.app).post("/api/turn/stream", json={
        **_STREAM_PAYLOAD, "turnId": "turn-python-event-http"})
    assert response.status_code == 200
    steps = [data for kind, data in _parse_sse(response.text) if kind == "step"]
    assert next(step for step in steps if step["stepId"] == "python.completed")["pythonResult"] == card


def test_normal_backend_passes_python_event_callback(monkeypatch):
    from server.app import excel_agent, python_coordinator
    captured = []

    async def callback(event):
        captured.append(event)

    async def execute(request, arguments, *, on_event):
        await on_event({"stepId": "python.completed"})
        return {"status": "completed"}

    monkeypatch.setattr(python_coordinator, "execute_python_analysis_async", execute)
    asyncio.run(excel_agent._execute_tool_async(None, "run_python_analysis", {}, on_event=callback))
    assert captured == [{"stepId": "python.completed"}]


def test_progress_schema_rejects_values_in_result_card():
    from pydantic import ValidationError
    from server.app.models import TaskProgressEvent
    with pytest.raises(ValidationError):
        TaskProgressEvent.model_validate(dict(schemaVersion=1, operationId="turn-python-event",
            sequence=1, phase="planning", stepId="python.completed", status="completed",
            title="完成", timestamp="2026-10-11T00:00:00Z", pythonResult=dict(
                resultRef="res_" + "a" * 32, title="结果", headers=["列1"], rowCount=1,
                columnCount=1, sourceSheets=[], complete=True, values=[["SECRET"]])))


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    """每次测试前清空登记器。"""
    from server.app import python_analysis

    registry._grants.clear()
    registry._budgets.clear()
    coordinator.reset()
    local_results.reset()
    # 启用 Python 执行
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(python_analysis, "_runner_status", lambda: "verified")
    yield
    registry._grants.clear()
    registry._budgets.clear()
    coordinator.reset()
    local_results.reset()


def test_python_completion_emits_event_with_result_handle():
    """Python 分析成功完成时，应发送包含 resultHandle 的事件。"""
    # 模拟请求对象
    class FakeRequest:
        turnId = "test-turn-123"

    # 1. 先登记数据（使用现有辅助函数）
    _, receipt = _bind(task_id="test-turn-123", values=[[1, 2], [3, 4]])

    # 2. 准备捕获事件
    emitted_events = []

    async def mock_on_event(event):
        emitted_events.append(event)

    # 3. 执行 Python 分析
    arguments = {
        "code": "result = {'kind': 'matrix', 'values': [[10, 20], [30, 40]]}",
        "datasetHandles": [receipt.datasetHandle],
    }

    result = asyncio.run(execute_python_analysis_async(
        FakeRequest(), arguments, on_event=mock_on_event
    ))

    # 4. 断言：执行成功
    assert result["status"] == "completed"
    assert "resultHandle" in result
    result_handle = result["resultHandle"]

    # 5. 断言：完成事件已发送（代码生命周期事件是另一类，不能顶替它）
    completion = [item for item in emitted_events if "pythonResult" in item]
    assert len(completion) == 1, emitted_events
    event = completion[0]
    assert len([item for item in emitted_events if item.get("status") == "completed"]) == 1
    assert event["pythonCode"]["stage"] == "analysis_completed"
    assert event["phase"] == "planning"
    # 步骤身份带本次执行的 run_id（`py_<12hex>`）：写死 `python.completed` 时，
    # 同一任务里跑两次分析会合并成同一个步骤。
    assert event["stepId"].startswith("python.analysis:py_"), event["stepId"]
    assert event["status"] == "completed"
    assert event["title"] == "隔离 Python 分析完成"
    assert "pythonResult" in event
    # pythonResult 现在是完整的结果卡
    assert event["pythonResult"]["resultRef"] == result_handle
    assert event["pythonResult"]["title"] == "Python 分析结果"
    assert event["pythonResult"]["rowCount"] > 0
    assert event["pythonResult"]["columnCount"] > 0

    # 6. 断言：代码块的状态事件**只带身份与规模**，不带正文。
    # 这是"生成的代码不进入持久化现场"的判据 —— 事件会被写进
    # planning_task_events（SQLite），正文一进去就落盘。
    lifecycle = [item for item in emitted_events if "pythonCode" in item]
    assert [item["pythonCode"]["stage"] for item in lifecycle] == [
        "generated",
        "validating",
        "running",
        "analysis_completed",
    ]
    for item in lifecycle:
        assert "code" not in item["pythonCode"]
        assert "delta" not in item["pythonCode"]
        assert item["stepId"].startswith("python.analysis:py_")
    # 同一次执行的全部生命周期事件落在**同一个**步骤上（否则代码块会变成好几块）。
    assert len({item["stepId"] for item in lifecycle}) == 1
    # 正文整段不出现在任何一条事件里。
    emitted_text = str(emitted_events)
    assert "result = {'kind': 'matrix'" not in emitted_text


def test_python_failure_does_not_emit_completion_event():
    """Python 分析失败时，不应发送完成事件，但代码块状态必须如实收尾。"""
    class FakeRequest:
        turnId = "test-turn-456"

    emitted_events = []

    async def mock_on_event(event):
        emitted_events.append(event)

    # 没有登记任何数据，直接调用应失败
    arguments = {
        "code": "result = 42",
        "datasetHandles": [],
    }

    result = asyncio.run(execute_python_analysis_async(
        FakeRequest(), arguments, on_event=mock_on_event
    ))

    # 断言：执行被拒绝
    assert result["status"] in ("rejected", "unavailable")
    assert "resultHandle" not in result

    # 断言：没有完成事件（没有结果卡就没有"分析完成"）
    assert [item for item in emitted_events if "pythonResult" in item] == []

    # 断言：代码块不能停在"正在运行"——执行没成，状态要落到失败。
    # 只断言"没有事件"会把一个卡在运行态的界面判成通过。
    stages = [
        item["pythonCode"]["stage"]
        for item in emitted_events
        if "pythonCode" in item
    ]
    # 这里一次**执行器都没启动**（没有可信数据登记 ⇒ prepare 阶段就停下），
    # 所以既没有 running，也没有 executedVersion —— 旧实现在这条路径上会发出
    # `validating → running(executedVersion=1) → execution_failed`，等于宣称
    # 一段从未运行的代码"跑过了"。
    assert stages == ["generated", "not_started"]
    assert "running" not in stages
    assert not any(
        "executedVersion" in item["pythonCode"]
        for item in emitted_events
        if "pythonCode" in item
    )
    assert not any(
        item["status"] == "completed"
        for item in emitted_events
        if "pythonCode" in item
    )


def test_python_event_includes_only_result_handle_not_values():
    """Python 完成事件只含 resultHandle，不含任何数据值。"""
    class FakeRequest:
        turnId = "test-turn-789"

    # 登记敏感数据（2x2 匹配 A1:B2）
    _, receipt = _bind(task_id="test-turn-789", values=[["SECRET_VALUE_123", "DATA1"], ["CONFIDENTIAL", "DATA2"]])

    emitted_events = []

    async def mock_on_event(event):
        emitted_events.append(event)

    arguments = {
        "code": "result = {'kind': 'matrix', 'values': [['ANALYZED_123']]}",
        "datasetHandles": [receipt.datasetHandle],
    }

    result = asyncio.run(execute_python_analysis_async(
        FakeRequest(), arguments, on_event=mock_on_event
    ))

    assert result["status"] == "completed"
    completion = [item for item in emitted_events if "pythonResult" in item]
    assert len(completion) == 1, emitted_events
    event = completion[0]
    assert len([item for item in emitted_events if item.get("status") == "completed"]) == 1
    assert event["pythonCode"]["stage"] == "analysis_completed"

    # 断言：**所有**事件里都没有源数据或分析结果的值（不只完成事件）
    event_str = str(emitted_events)
    assert "SECRET_VALUE" not in event_str
    assert "CONFIDENTIAL" not in event_str
    assert "ANALYZED_123" not in event_str
    # 只有不透明的句柄（resultRef），但结果卡本身是完整的
    assert "resultRef" in event["pythonResult"]
    assert event["pythonResult"]["title"] == "Python 分析结果"


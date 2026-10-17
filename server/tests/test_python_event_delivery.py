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

    card = dict(resultRef="res_" + "a" * 32, title="Python 分析结果", headers=["列1"],
                rowCount=2, columnCount=1, sourceSheets=["Sheet1"], complete=True)

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

    # 5. 断言：事件已发送
    assert len(emitted_events) == 1
    event = emitted_events[0]
    assert event["phase"] == "planning"
    assert event["stepId"] == "python.completed"
    assert event["status"] == "completed"
    assert event["title"] == "隔离 Python 分析完成"
    assert "pythonResult" in event
    # pythonResult 现在是完整的结果卡
    assert event["pythonResult"]["resultRef"] == result_handle
    assert event["pythonResult"]["title"] == "Python 分析结果"
    assert event["pythonResult"]["rowCount"] > 0
    assert event["pythonResult"]["columnCount"] > 0


def test_python_failure_does_not_emit_completion_event():
    """Python 分析失败时，不应发送完成事件。"""
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

    # 断言：没有事件被发送
    assert len(emitted_events) == 0


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
    event = emitted_events[0]

    # 断言：事件中没有任何源数据或分析结果的值
    event_str = str(event)
    assert "SECRET_VALUE" not in event_str
    assert "CONFIDENTIAL" not in event_str
    assert "ANALYZED_123" not in event_str
    # 只有不透明的句柄（resultRef），但结果卡本身是完整的
    assert "resultRef" in event["pythonResult"]
    assert event["pythonResult"]["title"] == "Python 分析结果"


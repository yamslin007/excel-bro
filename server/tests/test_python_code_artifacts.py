"""生成的 Python 代码：正文留在进程内，事件与账本只留身份。

边界（与仓库既有判据一致）：

* ``planning_task_events`` 里只能有 ``codeId`` / 阶段 / 版本 / 字符数；
* 取回正文只走本机回环端点，并且必须同时对上任务身份；
* 取不回来时要能如实报"没有这段代码"，而不是给一段猜的。
"""

from __future__ import annotations

import asyncio
import time

import pytest

from server.app.python_code_artifacts import PythonCodeRegistry, python_code_artifacts

CODE_ID = "code_py_0123456789ab"


@pytest.fixture(autouse=True)
def _clean_registry():
    python_code_artifacts.clear()
    yield
    python_code_artifacts.clear()


def _client(base_url: str = "http://localhost"):
    import httpx

    from server.app.main import app

    return httpx.AsyncClient(
        transport=httpx.ASGITransport(app=app), base_url=base_url
    )


# --- 登记表本身 -------------------------------------------------------------


def test_registry_returns_body_for_the_owning_task_only():
    registry = PythonCodeRegistry()
    registry.register(
        task_id="turn-a", code_id=CODE_ID, attempt_id="py_a", code="print(1)"
    )
    assert registry.get(CODE_ID, task_id="turn-a").code == "print(1)"
    # 别的任务拿同一个 codeId 问，必须塌成"没有这段代码"，
    # 否则这个入口就是跨任务读代码的工具。
    assert registry.get(CODE_ID, task_id="turn-b") is None
    assert registry.get(CODE_ID) is not None


def test_registry_drops_entries_after_ttl():
    registry = PythonCodeRegistry(ttl_seconds=0.01)
    registry.register(
        task_id="turn-a", code_id=CODE_ID, attempt_id="py_a", code="print(1)"
    )
    assert registry.get(CODE_ID) is not None
    time.sleep(0.05)
    assert registry.get(CODE_ID) is None


def test_registry_evicts_oldest_beyond_capacity():
    registry = PythonCodeRegistry(max_entries=2)
    for index in range(3):
        registry.register(
            task_id="turn-a",
            code_id=f"code_py_00000000000{index}",
            attempt_id=f"py_{index}",
            code=f"print({index})",
        )
        time.sleep(0.002)
    assert registry.get("code_py_000000000000") is None
    assert registry.get("code_py_000000000002") is not None


# --- 取回端点 ---------------------------------------------------------------


def test_endpoint_returns_body_and_forbids_caching():
    python_code_artifacts.register(
        task_id="turn-a", code_id=CODE_ID, attempt_id="py_a", code="print('hello')"
    )

    async def scenario():
        async with _client() as client:
            response = await client.post(
                "/api/sdk/python-code", json={"taskId": "turn-a", "codeId": CODE_ID}
            )
            return response

    response = asyncio.run(scenario())
    assert response.status_code == 200
    assert response.json() == {
        "codeId": CODE_ID,
        "attemptId": "py_a",
        "code": "print('hello')",
        "chars": len("print('hello')"),
    }
    assert response.headers.get("cache-control") == "no-store"


def test_endpoint_hides_code_from_another_task():
    python_code_artifacts.register(
        task_id="turn-a", code_id=CODE_ID, attempt_id="py_a", code="print(1)"
    )

    async def scenario():
        async with _client() as client:
            return await client.post(
                "/api/sdk/python-code", json={"taskId": "turn-b", "codeId": CODE_ID}
            )

    response = asyncio.run(scenario())
    assert response.status_code == 404
    assert response.json()["detail"]["code"] == "PYTHON_CODE_UNAVAILABLE"


def test_endpoint_reports_unavailable_instead_of_guessing():
    """没有登记过就是 404：绝不回一段猜的代码当"这就是跑的那段"。"""

    async def scenario():
        async with _client() as client:
            return await client.post(
                "/api/sdk/python-code", json={"taskId": "turn-a", "codeId": CODE_ID}
            )

    response = asyncio.run(scenario())
    assert response.status_code == 404
    assert response.json()["detail"]["code"] == "PYTHON_CODE_UNAVAILABLE"


def test_endpoint_rejects_malformed_code_id():
    """形状不对的 codeId 直接 422（不做前缀/模糊匹配，避免变成探测工具）。"""

    async def scenario():
        async with _client() as client:
            return await client.post(
                "/api/sdk/python-code", json={"taskId": "turn-a", "codeId": "code_py_x"}
            )

    assert asyncio.run(scenario()).status_code == 422


def test_endpoint_is_loopback_only():
    python_code_artifacts.register(
        task_id="turn-a", code_id=CODE_ID, attempt_id="py_a", code="print(1)"
    )

    async def scenario():
        async with _client("http://evil.example") as client:
            return await client.post(
                "/api/sdk/python-code", json={"taskId": "turn-a", "codeId": CODE_ID}
            )

    response = asyncio.run(scenario())
    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "BRIDGE_FORBIDDEN"

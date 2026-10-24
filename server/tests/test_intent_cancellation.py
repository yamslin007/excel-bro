"""Real ASGI disconnects must cancel model calls before a planning ID exists."""
import asyncio
import json

import pytest

from server.app import main
from server.app.models import IntentConversationResponse, MessageConversationResponse
from server.app.turn_state import turn_registry
from server.tests.test_main import _intent_payload


@pytest.mark.parametrize("path", ["/api/turn", "/api/intent", "/api/intent/message"])
@pytest.mark.parametrize("late_result", [False, True])
def test_disconnect_cancels_early_model_and_discards_late_decision(monkeypatch, path, late_result):
    async def scenario():
        started = asyncio.Event()
        cancelled = asyncio.Event()
        delivered = []
        decisions = []

        async def model(*_args, **_kwargs):
            started.set()
            try:
                await asyncio.Future()
            except asyncio.CancelledError:
                cancelled.set()
                if not late_result:
                    raise
                # Even a transport that returns success while being cancelled
                # cannot record/publish a decision for a disconnected request.
                response = MessageConversationResponse if path.endswith("message") else IntentConversationResponse
                return response(kind="conversation", provider="model", message="late result")

        monkeypatch.setattr(main, "check_intent", model)
        monkeypatch.setattr(main, "check_message_understanding", model)
        monkeypatch.setattr(main, "resolve_model_config", lambda *_: None)
        monkeypatch.setattr(main, "selected_model_config", lambda *_: None)
        monkeypatch.setattr(turn_registry, "record_decision", lambda *args: decisions.append(args))
        turn_registry.reset()
        payload = {"prompt": "分析当前表"} if path.endswith("message") else _intent_payload("分析当前表")
        queue = asyncio.Queue()
        await queue.put({"type": "http.request", "body": json.dumps(payload).encode(), "more_body": False})

        async def send(message):
            delivered.append(message)

        task = asyncio.create_task(main.app({
            "type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1",
            "method": "POST", "scheme": "http", "path": path, "raw_path": path.encode(),
            "query_string": b"", "root_path": "", "headers": [(b"content-type", b"application/json")],
            "client": ("127.0.0.1", 1234), "server": ("127.0.0.1", 8765),
        }, queue.get, send))
        try:
            await asyncio.wait_for(started.wait(), 2)
            await queue.put({"type": "http.disconnect"})
            await asyncio.wait_for(cancelled.wait(), 2)
            await asyncio.wait_for(task, 2)
            assert decisions == []
            assert not any(item.get("status") == 200 for item in delivered)
            assert b"late result" not in b"".join(item.get("body", b"") for item in delivered)
        finally:
            if not task.done():
                task.cancel()
            await asyncio.gather(task, return_exceptions=True)
            turn_registry.reset()

    asyncio.run(scenario())

"""Protocol tests plus a real installed SDK -> local gateway -> fake provider chain."""
import asyncio
import json
import re
import threading
from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import httpx
import pytest

from server.app import claude_sdk_agent
from server.app.agent_runtime import sdk_agent_connection
from server.app.llm.config import ModelConnection
from server.app.llm.errors import LLMHTTPStatusError, LLMResponseError, LLMUsageLimitError
from server.app.openai_sdk_protocol import CompletionStream, ProtocolError, request_payload
from server.app.openai_sdk_transport import completion_url, openai_sdk_transport


TOOL = {"name": "lookup", "description": "Synthetic lookup", "input_schema": {"type": "object", "properties": {"text": {"type": "string"}}}}
REQUEST = {"model": "ignored-sdk-alias", "stream": True, "max_tokens": 100,
           "system": [{"type": "text", "text": "System", "cache_control": {"type": "ephemeral"}}],
           "messages": [{"role": "user", "content": "Hello"}], "tools": [TOOL]}


def translate(body):
    return request_payload(body, "selected-model", max_tokens=200, supports_vision=True)


def chunk(delta=None, finish=None, **values):
    return {"choices": [{"index": 0, "delta": delta or {}, "finish_reason": finish}], **values}


def wire(chunks):
    return ("".join("data: " + json.dumps(item, ensure_ascii=False) + "\n\n" for item in chunks) + "data: [DONE]\n\n").encode()


@contextmanager
def provider(reply):
    calls = []
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass
        def do_POST(self):
            payload = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
            calls.append({"path": self.path, "authorization": self.headers.get("Authorization"), "payload": payload})
            status, body, mime = reply(payload, len(calls))
            self.send_response(status)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            try:
                self.wfile.write(body)
            except (BrokenPipeError, ConnectionResetError):
                pass
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/v1", calls
    finally:
        server.shutdown()
        server.server_close()
        thread.join(2)


def text_reply(payload, count):
    return 200, wire([chunk({"content": "你好"}), chunk(finish="stop", usage={"prompt_tokens": 10, "completion_tokens": 2})]), "text/event-stream"


def test_roundtrip_request_preserves_tool_ids_results_reasoning_images_and_schema():
    request = {**REQUEST, "messages": [
        {"role": "user", "content": [{"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": "YWJj"}}]},
        {"role": "assistant", "content": [{"type": "thinking", "thinking": "think", "signature": "local"},
            {"type": "tool_use", "id": "call1", "name": "lookup", "input": {"text": "值"}},
            {"type": "tool_use", "id": "call2", "name": "lookup", "input": {"text": "二"}}]},
        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": "call1", "content": [{"type": "text", "text": "OK"}]},
            {"type": "tool_result", "tool_use_id": "call2", "content": "失败", "is_error": True}, {"type": "text", "text": "继续"}]},
    ], "tool_choice": {"type": "tool", "name": "lookup"}}
    result = translate(request)
    assert result["model"] == "selected-model" and result["max_tokens"] == 100
    assert result["messages"][1]["content"][0]["image_url"]["url"] == "data:image/png;base64,YWJj"
    assert result["messages"][2]["reasoning_content"] == "think"
    assert result["messages"][3] == {"role": "tool", "tool_call_id": "call1", "content": "OK"}
    assert result["messages"][4]["content"] == "Tool error: 失败"
    assert result["tools"][0]["function"]["parameters"] == TOOL["input_schema"]
    assert result["tool_choice"]["function"]["name"] == "lookup"
    assert "cache_control" not in json.dumps(result)


@pytest.mark.parametrize("content", [
    [{"type": "tool_result", "tool_use_id": "unknown", "content": "OK"}],
    [{"type": "document", "source": {"type": "file", "file_id": "x"}}],
])
def test_unknown_content_or_unmatched_receipt_is_rejected(content):
    with pytest.raises(ProtocolError):
        translate({**REQUEST, "messages": [{"role": "user", "content": content}]})


def test_runtime_system_reminder_keeps_its_role_and_order():
    result = translate({**REQUEST, "messages": [*REQUEST["messages"], {"role": "system", "content": "Runtime reminder"}]})
    assert result["messages"][-1] == {"role": "system", "content": "Runtime reminder"}


def test_parallel_interleaved_tool_arguments_are_reassembled_before_any_tool_is_emitted():
    stream = CompletionStream("chosen", translate(REQUEST)["tools"])
    first = stream.feed(chunk({"tool_calls": [
        {"index": 0, "id": "a", "function": {"name": "lookup", "arguments": '{"text":"'}},
        {"index": 1, "id": "b", "function": {"name": "lookup", "arguments": '{"text":'}},
    ]}))
    assert first == []
    assert stream.feed(chunk({"tool_calls": [
        {"index": 1, "function": {"arguments": '"B"}'}}, {"index": 0, "function": {"arguments": 'A"}'}},
    ]}, "tool_calls")) == []
    events = stream.end()
    assert [e["content_block"]["id"] for e in events if e["type"] == "content_block_start"] == ["a", "b"]
    assert events[-2]["delta"]["stop_reason"] == "tool_use"
    assert events[-1]["type"] == "message_stop"


@pytest.mark.parametrize("finish,name,args,expected", [
    (None, "lookup", "{}", ProtocolError), ("length", "lookup", '{"x":', LLMUsageLimitError),
    ("tool_calls", "Bash", "{}", ProtocolError), ("tool_calls", "lookup", "{", ProtocolError),
    ("tool_calls", "lookup", "[]", ProtocolError), ("stop", "lookup", "{}", ProtocolError),
])
def test_incomplete_or_invalid_tool_stream_never_finishes(finish, name, args, expected):
    stream = CompletionStream("chosen", translate(REQUEST)["tools"])
    assert stream.feed(chunk({"tool_calls": [{"index": 0, "id": "x", "function": {"name": name, "arguments": args}}]}, finish)) == []
    with pytest.raises(expected):
        stream.end()


def test_bridge_binds_model_endpoint_credentials_and_closes_port():
    with provider(text_reply) as (url, calls):
        async def run():
            async with openai_sdk_transport(ModelConnection(url, "chosen", "upstream-secret", False)) as bridge:
                async with httpx.AsyncClient(trust_env=False) as client:
                    rejected = await client.post(bridge.base_url + "/v1/messages", json=REQUEST)
                    assert rejected.status_code == 401 and calls == []
                    assert (await client.post(bridge.base_url + "/arbitrary", headers={"x-api-key": bridge.token}, json=REQUEST)).status_code == 404
                    response = await client.post(bridge.base_url + "/v1/messages", headers={"x-api-key": bridge.token}, json=REQUEST)
                    assert response.status_code == 200 and "你好" in response.text and "message_stop" in response.text
                    assert "upstream-secret" not in response.text and bridge.token != "upstream-secret"
                    assert '"input_tokens": 10' in response.text
                local_url = bridge.base_url
            async with httpx.AsyncClient(trust_env=False) as client:
                with pytest.raises(httpx.ConnectError):
                    await client.post(local_url + "/v1/messages", json=REQUEST)
        asyncio.run(run())
        assert len(calls) == 1 and calls[0]["path"] == "/v1/chat/completions"
        assert calls[0]["payload"]["model"] == "chosen"
        assert calls[0]["authorization"] == "Bearer upstream-secret"


@pytest.mark.parametrize("status,code,error_type", [(401, "invalid_api_key", LLMHTTPStatusError), (429, "insufficient_quota", LLMUsageLimitError)])
def test_provider_failures_are_classified_without_echoing_keys_or_prompt(status, code, error_type):
    def reply(payload, count):
        return status, json.dumps({"error": {"code": code, "message": "secret-key private-prompt"}}).encode(), "application/json"
    with provider(reply) as (url, calls):
        async def run():
            async with openai_sdk_transport(ModelConnection(url, "chosen", "secret-key", False)) as bridge:
                async with httpx.AsyncClient(trust_env=False) as client:
                    response = await client.post(bridge.base_url + "/v1/messages", headers={"x-api-key": bridge.token}, json=REQUEST)
                assert response.status_code == status and isinstance(bridge.error, error_type)
                assert "secret-key" not in response.text and "private-prompt" not in response.text
        asyncio.run(run())


def test_explicit_parameter_rejection_supports_max_completion_tokens_without_changing_model():
    def reply(payload, count):
        if "max_tokens" in payload:
            return 400, b'{"error":{"code":"unsupported_parameter","param":"max_tokens"}}', "application/json"
        if "stream_options" in payload:
            return 400, b'{"error":{"code":"unknown_parameter","param":"stream_options"}}', "application/json"
        return text_reply(payload, count)
    with provider(reply) as (url, calls):
        async def run():
            async with openai_sdk_transport(ModelConnection(url, "chosen", "key", False)) as bridge:
                async with httpx.AsyncClient(trust_env=False) as client:
                    response = await client.post(bridge.base_url + "/v1/messages", headers={"x-api-key": bridge.token}, json=REQUEST)
                    assert "message_stop" in response.text
        asyncio.run(run())
        assert len(calls) == 3 and all(call["payload"]["model"] == "chosen" for call in calls)
        assert calls[-1]["payload"]["max_completion_tokens"] == 100


def test_parallel_sessions_cannot_cross_models_keys_or_bridge_tokens():
    with provider(text_reply) as (url, calls):
        async def run():
            async with openai_sdk_transport(ModelConnection(url, "model-a", "key-a", False)) as first:
                async with openai_sdk_transport(ModelConnection(url, "model-b", "key-b", False)) as second:
                    async with httpx.AsyncClient(trust_env=False) as client:
                        denied = await client.post(second.base_url + "/v1/messages", headers={"x-api-key": first.token}, json=REQUEST)
                        assert denied.status_code == 401 and calls == []
                        responses = await asyncio.gather(*[
                            client.post(binding.base_url + "/v1/messages", headers={"x-api-key": binding.token}, json={**REQUEST, "model": "attacker-alias", "base_url": "https://unrelated.invalid"})
                            for binding in (first, second)])
                        assert all("message_stop" in response.text for response in responses)
        asyncio.run(run())
        assert {(call["payload"]["model"], call["authorization"]) for call in calls} == {("model-a", "Bearer key-a"), ("model-b", "Bearer key-b")}


def test_context_cancellation_closes_listener_and_upstream_request():
    started, release = threading.Event(), threading.Event()
    def slow_reply(payload, count):
        started.set()
        release.wait(5)
        return text_reply(payload, count)
    with provider(slow_reply) as (url, calls):
        async def run():
            async with httpx.AsyncClient(trust_env=False) as client:
                async with openai_sdk_transport(ModelConnection(url, "m", "k", False)) as bridge:
                    local = bridge.base_url
                    pending = asyncio.create_task(client.post(local + "/v1/messages", headers={"x-api-key": bridge.token}, json=REQUEST))
                    for _ in range(100):
                        if started.is_set():
                            break
                        await asyncio.sleep(0.01)
                    assert started.is_set()
                    pending.cancel()
                    await asyncio.gather(pending, return_exceptions=True)
                with pytest.raises(httpx.ConnectError):
                    await client.post(local + "/v1/messages", json=REQUEST)
        try:
            asyncio.run(asyncio.wait_for(run(), 4))
        finally:
            release.set()


def test_request_limit_fails_before_contacting_provider(monkeypatch):
    from server.app import openai_sdk_transport as module
    original = module.capability_int
    monkeypatch.setattr(module, "capability_int", lambda section, key: 64 if key == "maxRequestBytes" else original(section, key))
    with provider(text_reply) as (url, calls):
        async def run():
            async with openai_sdk_transport(ModelConnection(url, "m", "k", False)) as bridge:
                async with httpx.AsyncClient(trust_env=False) as client:
                    response = await client.post(bridge.base_url + "/v1/messages", headers={"x-api-key": bridge.token}, json=REQUEST)
                    assert response.status_code == 413 and calls == []
        asyncio.run(run())


def test_real_sdk_completes_mcp_tool_roundtrip_through_openai_protocol(monkeypatch, tmp_path):
    """Uses the installed CLI, no network except the two loopback servers."""
    from server.app import openai_sdk_transport as transport_module
    observed = []
    original = transport_module.request_payload
    def inspect_request(body, *args, **kwargs):
        observed.append({"keys": list(body), "roles": [message.get("role") for message in body.get("messages", [])],
            "tools": [list(tool) for tool in body.get("tools", [])]})
        return original(body, *args, **kwargs)
    monkeypatch.setattr(transport_module, "request_payload", inspect_request)
    def reply(payload, count):
        results = [message for message in payload["messages"] if message["role"] == "tool"]
        if results:
            assert "OK" in results[-1]["content"]
            return 200, wire([chunk({"content": "OK"}, "stop")]), "text/event-stream"
        token = re.search(r"Verify connection using token: ([a-f0-9]{32})", json.dumps(payload["messages"]))
        assert token, "real SDK must send the synthetic token"
        name = next(tool["function"]["name"] for tool in payload["tools"] if "connection_probe" in tool["function"]["name"])
        return 200, wire([chunk({"tool_calls": [{"index": 0, "id": "probe1", "type": "function", "function": {
            "name": name, "arguments": json.dumps({"token": token[1]}),
        }}]}, "tool_calls")]), "text/event-stream"
    monkeypatch.setattr(claude_sdk_agent, "_state_dir", lambda: tmp_path)
    with provider(reply) as (url, calls):
        try:
            asyncio.run(claude_sdk_agent.check_sdk_connection(ModelConnection(url, "generic-test-model", "synthetic-key", False)))
        except Exception as error:
            raise AssertionError(f"SDK request shapes: {observed}") from error
    assert len(calls) >= 2
    assert all(call["payload"]["model"] == "generic-test-model" for call in calls)
    assert any(any(message["role"] == "tool" for message in call["payload"]["messages"]) for call in calls)


@pytest.mark.parametrize("url", ["https://gateway.invalid/v1", "http://127.0.0.1:1234/v1", "https://gateway.invalid/custom/chat/completions"])
def test_generic_connection_routes_without_altering_credentials_or_address(url):
    connection = ModelConnection(url, "model", "key", False)
    assert sdk_agent_connection(connection) is connection
    assert completion_url(url).endswith("/chat/completions")

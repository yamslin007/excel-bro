"""Per-query loopback transport for the SDK's Anthropic HTTP client.

The upstream address/model/key are bound by trusted connection configuration.
The SDK receives only an ephemeral local credential. No persistent gateway,
global connection selection, prompt logging, or additional tools are created.
"""
from __future__ import annotations

import asyncio
import codecs
import hmac
import json
import secrets
import socket
from contextlib import asynccontextmanager, nullcontext
from dataclasses import dataclass
from typing import AsyncIterator
from urllib.parse import urlsplit

import httpx
import uvicorn
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse, StreamingResponse
from starlette.routing import Route

from .capabilities import capability_float, capability_int
from .llm.config import ModelConnection
from .llm.errors import LLMError, LLMHTTPStatusError, LLMResponseError, LLMTimeoutError, LLMTransportError, LLMUsageLimitError, model_usage_error_code
from .openai_sdk_protocol import CompletionStream, ProtocolError, event, request_payload, sse


def completion_url(base_url: str) -> str:
    parsed = urlsplit(base_url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname or parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise ProtocolError("模型服务地址无效，请填写 OpenAI 兼容接口的基础地址。")
    base = base_url.rstrip("/")
    return base if parsed.path.rstrip("/").endswith("/chat/completions") else base + "/chat/completions"


async def completion_chunks(response: httpx.Response, max_bytes: int) -> AsyncIterator[dict]:
    decoder = codecs.getincrementaldecoder("utf-8")("strict")
    buffer, data = "", []
    total = 0
    async for raw in response.aiter_bytes():
        total += len(raw)
        if total > max_bytes:
            raise ProtocolError("模型响应超过适配通道大小上限，已停止接收。")
        buffer += decoder.decode(raw)
        while "\n" in buffer:
            line, buffer = buffer.split("\n", 1)
            line = line.rstrip("\r")
            if line.startswith("data:"):
                data.append(line[5:].lstrip(" "))
            elif not line and data:
                value = "\n".join(data)
                data = []
                if value == "[DONE]":
                    return
                try:
                    parsed = json.loads(value)
                except ValueError as error:
                    raise ProtocolError("模型流返回了无法解析的数据。") from error
                if isinstance(parsed, dict) and parsed.get("error"):
                    code = model_usage_error_code(value)
                    if code:
                        raise LLMUsageLimitError(code)
                    raise ProtocolError("模型服务在流式响应中报告错误。")
                yield parsed
    buffer += decoder.decode(b"", final=True)
    if buffer.strip() or data:
        raise ProtocolError("模型流中断在未完成的数据块，未接受本轮结果。")


async def _bounded_error(response: httpx.Response) -> str:
    body = bytearray()
    async for chunk in response.aiter_bytes():
        body.extend(chunk[:max(0, 16384 - len(body))])
        if len(body) >= 16384:
            break
    return body.decode("utf-8", errors="replace")


def _parameter_fallback(payload: dict, body: str) -> bool:
    """Only retry explicit parameter rejection, never a successful model turn."""
    try:
        error = json.loads(body).get("error", {})
        param, code, message = error.get("param"), error.get("code"), error.get("message", "")
    except (ValueError, AttributeError):
        return False
    if code not in {"unsupported_parameter", "unknown_parameter"} and "not supported" not in str(message).lower():
        return False
    if param == "max_tokens" and "max_tokens" in payload:
        payload["max_completion_tokens"] = payload.pop("max_tokens")
        return True
    if param == "stream_options" and "stream_options" in payload:
        payload.pop("stream_options")
        return True
    return False


@dataclass
class TransportBinding:
    base_url: str
    token: str
    error: LLMError | None = None


class _LocalServer(uvicorn.Server):
    # An embedded server must never replace the parent backend's signal handlers.
    def capture_signals(self):
        return nullcontext()


@asynccontextmanager
async def openai_sdk_transport(connection: ModelConnection):
    endpoint = completion_url(connection.base_url)
    binding = TransportBinding("", secrets.token_urlsafe(32))
    max_request = capability_int("openaiSdkTransport", "maxRequestBytes")
    max_response = capability_int("openaiSdkTransport", "maxResponseBytes")
    max_requests = capability_int("openaiSdkTransport", "maxRequests")
    count = 0
    active: set[asyncio.Task] = set()
    responses: set[httpx.Response] = set()
    client = httpx.AsyncClient(timeout=httpx.Timeout(capability_float("openaiSdkTransport", "requestTimeoutSeconds")),
        follow_redirects=False, trust_env=urlsplit(endpoint).hostname not in {"127.0.0.1", "localhost", "::1"})

    def failure(error: LLMError, status: int = 400):
        binding.error = error
        return JSONResponse({"type": "error", "error": {"type": "invalid_request_error", "message": str(error)}}, status_code=status)

    async def messages(request: Request):
        nonlocal count
        supplied = request.headers.get("x-api-key") or request.headers.get("authorization", "").removeprefix("Bearer ")
        if not hmac.compare_digest(supplied, binding.token):
            return JSONResponse({"error": {"type": "authentication_error", "message": "Invalid local credential"}}, status_code=401)
        if count >= max_requests:
            return failure(LLMUsageLimitError("MODEL_TURN_LIMIT"), 429)
        count += 1
        binding.error = None
        current = asyncio.current_task()
        active.add(current)
        upstream = None
        try:
            body = bytearray()
            async for chunk in request.stream():
                body.extend(chunk)
                if len(body) > max_request:
                    return failure(ProtocolError("模型请求超过适配通道大小上限。"), 413)
            incoming = json.loads(body)
            if not isinstance(incoming, dict):
                raise ProtocolError("模型请求必须是对象。")
            if incoming.get("stream") is not True:
                return failure(ProtocolError("SDK 适配通道要求流式请求。"))
            payload = request_payload(incoming, connection.model,
                max_tokens=capability_int("openaiSdkTransport", "maxOutputTokens"), supports_vision=connection.supports_vision)
            headers = {"Content-Type": "application/json", "Accept": "text/event-stream"}
            if connection.api_key:
                headers["Authorization"] = "Bearer " + connection.api_key
            for attempt in range(3):
                upstream = await client.send(client.build_request("POST", endpoint, json=payload, headers=headers), stream=True)
                if upstream.is_success:
                    break
                raw_error = await _bounded_error(upstream)
                status = upstream.status_code
                await upstream.aclose()
                if status == 400 and attempt < 2 and _parameter_fallback(payload, raw_error):
                    continue
                usage_code = model_usage_error_code(raw_error)
                if usage_code:
                    return failure(LLMUsageLimitError(usage_code), status if status >= 400 else 502)
                # Never echo provider body: it can include keys or prompt text.
                error = LLMHTTPStatusError(status, "")
                return failure(error, status if status >= 400 else 502)
            binding.error = None
            responses.add(upstream)

            async def translated():
                task = asyncio.current_task()
                active.add(task)
                converter = CompletionStream(connection.model, payload.get("tools", []))
                try:
                    yield sse(converter.start())
                    if "text/event-stream" not in upstream.headers.get("content-type", ""):
                        raise ProtocolError("模型接口未按请求返回 SSE 流；请检查流式响应支持。")
                    async for chunk in completion_chunks(upstream, max_response):
                        for item in converter.feed(chunk):
                            yield sse(item)
                    for item in converter.end():
                        yield sse(item)
                except (LLMError, httpx.HTTPError, ValueError, TypeError, KeyError, AttributeError) as error:
                    safe = _safe_error(error)
                    binding.error = safe
                    yield sse(event("error", error={"type": "api_error", "message": str(safe)}))
                finally:
                    await upstream.aclose()
                    responses.discard(upstream)
                    active.discard(task)

            return StreamingResponse(translated(), media_type="text/event-stream", headers={"Cache-Control": "no-store"})
        except asyncio.CancelledError:
            if upstream is not None:
                await upstream.aclose()
            # Parent cancellation closes this task-local bridge. Return no
            # successful result and avoid logging a routine ASGI traceback.
            return JSONResponse({"type": "error", "error": {"type": "api_error", "message": "Model request cancelled"}}, status_code=499)
        except (LLMError, httpx.HTTPError, ValueError, TypeError, KeyError, AttributeError) as error:
            if upstream is not None:
                await upstream.aclose()
            return failure(_safe_error(error), 502)
        finally:
            active.discard(current)

    app = Starlette(routes=[Route("/v1/messages", messages, methods=["POST"])])
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_task = None
    server = None
    try:
        sock.bind(("127.0.0.1", 0))
        sock.listen(8)
        binding.base_url = f"http://127.0.0.1:{sock.getsockname()[1]}"
        server = _LocalServer(uvicorn.Config(app, host="127.0.0.1", port=0, lifespan="off", access_log=False,
            log_config=None, log_level=None, timeout_graceful_shutdown=1, ws="none"))
        server_task = asyncio.create_task(server.serve(sockets=[sock]))
        async with asyncio.timeout(5):
            while not server.started:
                if server_task.done():
                    await server_task
                    raise ProtocolError("本地模型适配通道未能启动。")
                await asyncio.sleep(0.01)
        yield binding
    finally:
        for task in list(active):
            task.cancel()
        if active:
            await asyncio.gather(*list(active), return_exceptions=True)
        for response in list(responses):
            await response.aclose()
        await client.aclose()
        if server is not None:
            server.should_exit = True
        if server_task is not None:
            try:
                await asyncio.wait_for(server_task, timeout=2)
            except (asyncio.TimeoutError, asyncio.CancelledError):
                server_task.cancel()
                await asyncio.gather(server_task, return_exceptions=True)
        sock.close()


def _safe_error(error: Exception) -> LLMError:
    if isinstance(error, LLMError):
        return error
    if isinstance(error, httpx.TimeoutException):
        return LLMTimeoutError("OpenAI 兼容模型响应超时。")
    if isinstance(error, httpx.HTTPError):
        return LLMTransportError("无法连接所选 OpenAI 兼容接口，请检查服务地址和网络。")
    return LLMResponseError("OpenAI 兼容接口返回了不支持或无效的响应结构。")

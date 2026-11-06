"""Messages wire adapter for the shared text/formula client (no host access).

The caller-facing shape stays Chat Completions so formula parsing is shared.
Reference: https://platform.claude.com/docs/en/api/messages
"""
from __future__ import annotations

import json
from typing import Any
from urllib.parse import urlsplit

from .errors import LLMResponseError, LLMUsageLimitError


def messages_url(base_url: str) -> str:
    base = base_url.rstrip("/")
    parsed = urlsplit(base)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname or parsed.query or parsed.fragment or parsed.username or parsed.password:
        raise LLMResponseError("模型服务地址无效，请填写不含凭据、查询参数的接口地址。")
    if parsed.path.endswith("/messages"):
        return base
    return base + ("/messages" if parsed.path.endswith("/v1") else "/v1/messages")


def _content(value: Any, supports_vision: bool) -> list[dict[str, Any]]:
    if value is None:
        return []
    if isinstance(value, str):
        return [{"type": "text", "text": value}] if value else []
    if not isinstance(value, list):
        raise LLMResponseError("模型消息内容格式无效。")
    blocks = []
    for block in value:
        if not isinstance(block, dict):
            raise LLMResponseError("模型消息内容格式无效。")
        if block.get("type") == "text" and isinstance(block.get("text"), str):
            blocks.append({"type": "text", "text": block["text"]})
        elif block.get("type") == "image_url" and supports_vision:
            url = block.get("image_url", {}).get("url", "")
            if url.startswith("data:image/") and ";base64," in url:
                media_type, data = url[5:].split(";base64,", 1)
                if media_type not in {"image/png", "image/jpeg", "image/gif", "image/webp"}:
                    raise LLMResponseError("模型不支持该图片格式。")
                source = {"type": "base64", "media_type": media_type, "data": data}
            elif url.startswith("https://"):
                source = {"type": "url", "url": url}
            else:
                raise LLMResponseError("模型图片地址无效。")
            blocks.append({"type": "image", "source": source})
        else:
            raise LLMResponseError("所选模型不支持此消息内容。")
    return blocks


def request_payload(payload: dict[str, Any], *, supports_vision: bool, default_max_tokens: int) -> dict[str, Any]:
    messages, system = [], []
    for message in payload["messages"]:
        role = message["role"]
        blocks = _content(message.get("content"), supports_vision)
        if role in {"system", "developer"}:
            if any(block["type"] != "text" for block in blocks):
                raise LLMResponseError("系统消息只能包含文本。")
            system.extend(blocks)
            continue
        if role == "tool":
            blocks = [{"type": "tool_result", "tool_use_id": message["tool_call_id"], "content": blocks}]
            role = "user"
        elif role == "assistant":
            for call in message.get("tool_calls", []):
                try:
                    arguments = json.loads(call["function"]["arguments"])
                except (ValueError, KeyError, TypeError) as error:
                    raise LLMResponseError("工具参数不是有效 JSON。") from error
                if not isinstance(arguments, dict):
                    raise LLMResponseError("工具参数必须是对象。")
                blocks.append({"type": "tool_use", "id": call["id"], "name": call["function"]["name"], "input": arguments})
        elif role != "user":
            raise LLMResponseError("模型消息角色无效。")
        messages.append({"role": role, "content": blocks})
    result = {"model": payload["model"], "messages": messages, "max_tokens": payload.get("max_tokens", default_max_tokens)}
    if system:
        result["system"] = system
    if "temperature" in payload:
        result["temperature"] = payload["temperature"]
    if payload.get("tools"):
        result["tools"] = [{"name": tool["function"]["name"],
                            "description": tool["function"].get("description", ""),
                            "input_schema": tool["function"]["parameters"]} for tool in payload["tools"]]
        choice = payload.get("tool_choice")
        if choice in {"auto", "none", "required"}:
            result["tool_choice"] = {"type": "any" if choice == "required" else choice}
    return result


def response_payload(body: dict[str, Any]) -> dict[str, Any]:
    if body.get("type") != "message" or body.get("role") != "assistant" or not isinstance(body.get("content"), list):
        raise LLMResponseError("模型未返回有效的 Messages 响应。")
    stop = body.get("stop_reason")
    if stop == "max_tokens":
        raise LLMUsageLimitError("MODEL_OUTPUT_LIMIT")
    if stop not in {"end_turn", "stop_sequence", "tool_use"}:
        raise LLMResponseError("模型未返回完整结果。")
    text, calls = [], []
    for block in body["content"]:
        if not isinstance(block, dict):
            raise LLMResponseError("模型返回的内容格式无效。")
        if block.get("type") == "text" and isinstance(block.get("text"), str):
            text.append(block["text"])
        elif block.get("type") == "tool_use":
            if not isinstance(block.get("input"), dict) or not block.get("id") or not block.get("name"):
                raise LLMResponseError("模型工具调用格式无效。")
            calls.append({"id": block["id"], "type": "function", "function": {
                "name": block["name"], "arguments": json.dumps(block["input"], ensure_ascii=False)}})
        elif block.get("type") not in {"thinking", "redacted_thinking"}:
            raise LLMResponseError("模型返回了不支持的内容类型。")
    if not any(part.strip() for part in text) and not calls:
        raise LLMResponseError("模型未返回有效文本。")
    if bool(calls) != (stop == "tool_use"):
        raise LLMResponseError("模型工具调用结束状态不一致。")
    message: dict[str, Any] = {"role": "assistant", "content": "".join(text)}
    if calls:
        message["tool_calls"] = calls
    raw = body.get("usage") or {}
    def count(key: str) -> int:
        value = raw.get(key, 0) if isinstance(raw, dict) else 0
        return value if isinstance(value, int) and not isinstance(value, bool) and value >= 0 else 0
    prompt = count("input_tokens") + count("cache_creation_input_tokens") + count("cache_read_input_tokens")
    output = count("output_tokens")
    return {"choices": [{"message": message, "finish_reason": "tool_calls" if calls else "stop"}],
            "usage": {"prompt_tokens": prompt, "completion_tokens": output, "total_tokens": prompt + output}}

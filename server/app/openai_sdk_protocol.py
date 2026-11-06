"""Pure Anthropic Messages <-> OpenAI Chat Completions translation.

No workbook, credential, network or execution access. Only the SDK's existing
messages and tool definitions cross this boundary.
"""
from __future__ import annotations

import json
import uuid
from typing import Any

from .llm.errors import LLMResponseError, LLMUsageLimitError


class ProtocolError(LLMResponseError):
    pass


def _text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if not isinstance(value, list):
        raise ProtocolError("模型接口返回了不支持的文本结构。")
    if any(not isinstance(block, dict) or block.get("type") != "text" for block in value):
        raise ProtocolError("当前 OpenAI 兼容接口通道不支持此内容类型。")
    return "\n".join(str(block.get("text", "")) for block in value)


def request_payload(request: dict, model: str, *, max_tokens: int, supports_vision: bool) -> dict:
    messages: list[dict] = []
    if request.get("system"):
        messages.append({"role": "system", "content": _text(request["system"])})
    history = request.get("messages")
    if not isinstance(history, list) or not history:
        raise ProtocolError("模型请求缺少消息历史。")
    pending: set[str] = set()
    for message in history:
        role = message.get("role")
        # Current SDK versions append runtime system reminders to messages as
        # well as using the top-level system field. Preserve their role/order.
        if role not in {"user", "assistant", "system", "developer"}:
            raise ProtocolError("模型请求包含不支持的消息角色。")
        blocks = message.get("content", [])
        if isinstance(blocks, str):
            blocks = [{"type": "text", "text": blocks}]
        if not isinstance(blocks, list):
            raise ProtocolError("模型消息格式无效。")
        content, calls, results, thinking = [], [], [], []
        for block in blocks:
            if not isinstance(block, dict):
                raise ProtocolError("模型消息内容格式无效。")
            kind = block.get("type")
            if kind == "text":
                content.append({"type": "text", "text": str(block.get("text", ""))})
            elif kind == "image" and role == "user":
                if not supports_vision:
                    raise ProtocolError("所选连接未声明支持图片。")
                source = block.get("source", {})
                if source.get("type") == "base64" and source.get("media_type") in {"image/png", "image/jpeg", "image/gif", "image/webp"}:
                    url = f"data:{source['media_type']};base64,{source.get('data', '')}"
                elif source.get("type") == "url" and str(source.get("url", "")).startswith("https://"):
                    url = source["url"]
                else:
                    raise ProtocolError("模型请求中的图片格式不受支持。")
                content.append({"type": "image_url", "image_url": {"url": url}})
            elif kind == "thinking" and role == "assistant":
                thinking.append(str(block.get("thinking", "")))
            elif kind == "tool_use" and role == "assistant":
                call_id, name, arguments = block.get("id"), block.get("name"), block.get("input")
                if not isinstance(call_id, str) or not call_id or call_id in pending or not isinstance(name, str) or not isinstance(arguments, dict):
                    raise ProtocolError("模型工具调用标识或参数无效。")
                pending.add(call_id)
                calls.append({"id": call_id, "type": "function", "function": {
                    "name": name, "arguments": json.dumps(arguments, ensure_ascii=False, allow_nan=False),
                }})
            elif kind == "tool_result" and role == "user":
                call_id = block.get("tool_use_id")
                if call_id not in pending:
                    raise ProtocolError("工具结果与原调用不匹配。")
                pending.remove(call_id)
                result = _text(block.get("content", ""))
                if block.get("is_error"):
                    result = "Tool error: " + result
                results.append({"role": "tool", "tool_call_id": call_id, "content": result})
            else:
                raise ProtocolError("当前 OpenAI 兼容通道不支持此 SDK 消息类型。")
        messages.extend(results)
        if role == "user" and pending:
            raise ProtocolError("工具调用缺少对应结果，无法继续。")
        if content or calls or thinking:
            plain = all(part["type"] == "text" for part in content)
            entry: dict = {"role": role, "content": "\n".join(part["text"] for part in content) if plain else content}
            if calls:
                entry["tool_calls"] = calls
            if thinking:
                entry["reasoning_content"] = "".join(thinking)
            messages.append(entry)
    if pending:
        raise ProtocolError("模型请求仍有未完成的工具调用。")
    requested_limit = request.get("max_tokens", max_tokens)
    if isinstance(requested_limit, bool) or not isinstance(requested_limit, int) or requested_limit <= 0:
        raise ProtocolError("模型输出限额无效。")
    payload: dict = {"model": model, "messages": messages, "stream": True,
                     "max_tokens": min(requested_limit, max_tokens), "stream_options": {"include_usage": True}}
    tools = request.get("tools", [])
    if tools:
        payload["tools"] = []
        for tool in tools:
            if not isinstance(tool.get("name"), str) or not isinstance(tool.get("input_schema"), dict) or tool.get("type", "custom") != "custom":
                raise ProtocolError("此通道仅支持已登记的客户端工具。")
            payload["tools"].append({"type": "function", "function": {
                "name": tool["name"], "description": tool.get("description", ""), "parameters": tool["input_schema"],
            }})
        choice = request.get("tool_choice", {"type": "auto"})
        mode = choice.get("type")
        if mode == "tool":
            if choice.get("name") not in {tool["name"] for tool in tools}:
                raise ProtocolError("指定的工具不在本轮工具清单中。")
            payload["tool_choice"] = {"type": "function", "function": {"name": choice["name"]}}
        elif mode in {"auto", "any", "none"}:
            payload["tool_choice"] = {"auto": "auto", "any": "required", "none": "none"}[mode]
        else:
            raise ProtocolError("不支持此工具选择方式。")
        if choice.get("disable_parallel_tool_use"):
            payload["parallel_tool_calls"] = False
    for key in ("temperature", "top_p"):
        if key in request:
            payload[key] = request[key]
    if request.get("stop_sequences"):
        payload["stop"] = request["stop_sequences"]
    output = request.get("output_config", {}).get("format") or request.get("output_format")
    if output:
        if output.get("type") != "json_schema" or not isinstance(output.get("schema"), dict):
            raise ProtocolError("不支持此结构化输出格式。")
        payload["response_format"] = {"type": "json_schema", "json_schema": {
            "name": "response", "schema": output["schema"], "strict": True,
        }}
    return payload


def event(kind: str, **values: Any) -> dict:
    return {"type": kind, **values}


def sse(value: dict) -> bytes:
    return (f"event: {value['type']}\ndata: " + json.dumps(value, ensure_ascii=False, allow_nan=False) + "\n\n").encode()


class CompletionStream:
    """Text streams immediately; tool calls become visible only after validation."""
    def __init__(self, model: str, tools: list[dict]):
        self.model = model
        self.allowed = {tool["function"]["name"] for tool in tools}
        self.calls: dict[int, dict] = {}
        self.blocks: dict[str, int] = {}
        self.index = 0
        self.finish: str | None = None
        self.usage = {"input_tokens": 0, "output_tokens": 0}
        self.has_content = False

    def start(self) -> dict:
        return event("message_start", message={"id": "msg_" + uuid.uuid4().hex,
            "type": "message", "role": "assistant", "model": self.model, "content": [],
            "stop_reason": None, "stop_sequence": None, "usage": self.usage.copy()})

    def feed(self, chunk: dict) -> list[dict]:
        if not isinstance(chunk, dict) or chunk.get("error"):
            raise ProtocolError("模型流返回了错误，未完成本轮响应。")
        usage = chunk.get("usage")
        if isinstance(usage, dict):
            for upstream, local in (("prompt_tokens", "input_tokens"), ("completion_tokens", "output_tokens")):
                count = usage.get(upstream)
                if isinstance(count, int) and not isinstance(count, bool) and count >= 0:
                    self.usage[local] = count
        choices = chunk.get("choices", [])
        if not choices:
            return []
        if len(choices) != 1 or choices[0].get("index", 0) != 0:
            raise ProtocolError("模型接口返回了不支持的多候选响应。")
        choice = choices[0]
        delta = choice.get("delta", {})
        if self.finish is not None and any(delta.get(key) for key in ("content", "tool_calls", "reasoning_content", "refusal")):
            raise ProtocolError("模型在结束标记后继续输出，响应无效。")
        events = []
        for field, kind, delta_kind in (("reasoning_content", "thinking", "thinking_delta"), ("content", "text", "text_delta"), ("refusal", "text", "text_delta")):
            value = delta.get(field)
            if value is None or value == "":
                continue
            if not isinstance(value, str):
                raise ProtocolError("模型流文本格式无效。")
            self.has_content = True
            if kind not in self.blocks:
                self.blocks[kind] = self.index
                self.index += 1
                events.append(event("content_block_start", index=self.blocks[kind], content_block={"type": kind, kind: ""}))
            events.append(event("content_block_delta", index=self.blocks[kind], delta={"type": delta_kind, kind: value}))
        for call in delta.get("tool_calls") or []:
            index = call.get("index")
            if not isinstance(index, int) or isinstance(index, bool) or index < 0:
                raise ProtocolError("工具流缺少有效序号。")
            target = self.calls.setdefault(index, {"id": "", "name": "", "arguments": ""})
            if call.get("id"):
                if target["id"] and target["id"] != call["id"]:
                    raise ProtocolError("工具流调用标识发生变化。")
                target["id"] = call["id"]
            function = call.get("function", {})
            for field in ("name", "arguments"):
                value = function.get(field)
                if value is not None:
                    if not isinstance(value, str):
                        raise ProtocolError("工具流参数格式无效。")
                    target[field] += value
        if choice.get("finish_reason") is not None:
            self.finish = choice["finish_reason"]
        return events

    def end(self) -> list[dict]:
        if self.finish == "length":
            raise LLMUsageLimitError("MODEL_OUTPUT_LIMIT")
        if self.finish not in {"stop", "tool_calls"}:
            raise ProtocolError("模型响应未完整结束，未接受工具调用或结果。")
        if (self.finish == "tool_calls") != bool(self.calls):
            raise ProtocolError("模型结束原因与工具调用不一致。")
        if not self.has_content and not self.calls:
            raise ProtocolError("模型接口返回了空响应。")
        # Validate every call before publishing any call, including parallel calls.
        ids: set[str] = set()
        for call in self.calls.values():
            if not isinstance(call["id"], str) or not call["id"] or call["id"] in ids or call["name"] not in self.allowed:
                raise ProtocolError("模型请求了未知工具或重复的工具调用。")
            ids.add(call["id"])
            try:
                arguments = json.loads(call["arguments"], parse_constant=lambda _: (_ for _ in ()).throw(ValueError()))
            except (ValueError, TypeError) as error:
                raise ProtocolError("模型工具参数不是完整 JSON，未调用工具。") from error
            if not isinstance(arguments, dict):
                raise ProtocolError("模型工具参数必须是对象。")
        events = []
        for kind, index in self.blocks.items():
            if kind == "thinking":
                events.append(event("content_block_delta", index=index, delta={"type": "signature_delta", "signature": "excelbro-openai-reasoning"}))
            events.append(event("content_block_stop", index=index))
        for call in self.calls.values():
            index = self.index
            self.index += 1
            events.extend([
                event("content_block_start", index=index, content_block={"type": "tool_use", "id": call["id"], "name": call["name"], "input": {}}),
                event("content_block_delta", index=index, delta={"type": "input_json_delta", "partial_json": call["arguments"]}),
                event("content_block_stop", index=index),
            ])
        events.extend([event("message_delta", delta={"stop_reason": "tool_use" if self.calls else "end_turn", "stop_sequence": None}, usage=self.usage), event("message_stop")])
        return events

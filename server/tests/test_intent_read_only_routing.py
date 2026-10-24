from __future__ import annotations

import asyncio
import json
from typing import Any

import pytest

from server.app import intent as intent_module
from server.app.intent import check_intent, is_explicit_read_only_query
from server.app.models import IntentCheckRequest
from server.app.llm import LLMResponseError
from server.tests.sdk_text_harness import install_text_transport
from server.tests.test_claude_sdk_agent import _connection


def _request(prompt: str) -> IntentCheckRequest:
    return IntentCheckRequest.model_validate(
        {
            "prompt": prompt,
            "scope": {
                "workbookName": "测试.xlsx",
                "sourceMode": "workbook",
                "selectionMode": "auto",
                "activeWorksheet": "明细",
                "totalWorksheetCount": 1,
                "worksheetNames": ["明细"],
                "sheets": [
                    {
                        "name": "明细",
                        "rowCount": 100,
                        "columnCount": 3,
                        "headers": ["影院名称", "影院编码", "影片名称"],
                    }
                ],
            },
        }
    )


class _FakeClient:
    def __init__(self, responses: list[dict[str, Any]]) -> None:
        self.responses = responses
        self.calls: list[list[dict[str, Any]]] = []

    async def __aenter__(self) -> "_FakeClient":
        return self

    async def __aexit__(self, exc_type, exc, tb) -> bool:
        return False

    async def chat_completions(
        self, *, messages: list[dict[str, Any]], **_kwargs: Any
    ) -> tuple[dict[str, Any], dict[str, int]]:
        self.calls.append(messages)
        response = self.responses[len(self.calls) - 1]
        return (
            {"choices": [{"message": {"content": json.dumps(response, ensure_ascii=False)}}]},
            {},
        )


def _install_client(
    monkeypatch: pytest.MonkeyPatch, client: _FakeClient
) -> None:
    install_text_transport(monkeypatch, client)


def test_explicit_count_query_rejects_proceed_and_repairs_to_tool_request(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    prompt = "怀化横店电影城有多少条记录"
    client = _FakeClient(
        [
            {
                "kind": "proceed",
                "summary": "统计该影院记录数",
                "confirmedPrompt": prompt,
            },
            {
                "kind": "tool_request",
                "summary": "统计该影院记录数",
                "confirmedPrompt": prompt,
                "request": {
                    "id": "count-cinema",
                    "tool": "query_table",
                    "arguments": {
                        "mode": "aggregate",
                        "scope": "active",
                        "filters": [
                            {
                                "field": "影院名称",
                                "operator": "equals",
                                "value": "怀化横店电影城",
                            }
                        ],
                        "metrics": [
                            {
                                "operation": "countRows",
                                "outputName": "记录数",
                            }
                        ],
                        "limit": 20,
                    },
                },
            },
        ]
    )
    _install_client(monkeypatch, client)

    response = asyncio.run(check_intent(_request(prompt), config=_connection()))

    assert response.kind == "tool_request"
    assert len(client.calls) == 2
    assert json.loads(client.calls[0][1]["content"])["readOnlyRequired"] is True
    assert "上次输出未通过结构校验" in client.calls[1][1]["content"]


def test_repeated_invalid_read_only_routing_fails_closed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    prompt = "怀化横店电影城有多少条记录"
    proceed = {
        "kind": "proceed",
        "summary": "统计该影院记录数",
        "confirmedPrompt": prompt,
    }
    client = _FakeClient([proceed, proceed])
    _install_client(monkeypatch, client)

    response = asyncio.run(check_intent(_request(prompt), config=_connection()))

    assert response.kind == "clarification"
    assert response.provider == "local"
    assert "不会写入 Excel" in response.clarification.summary


def test_intent_rejects_parseable_but_truncated_response_without_format_retry(monkeypatch):
    class TruncatedClient(_FakeClient):
        async def chat_completions(self, *, messages, **kwargs):
            self.calls.append(messages)
            assert kwargs["max_tokens"] is None
            return {"choices": [{"finish_reason": "length", "message": {"content": json.dumps({
                "kind": "conversation", "message": "不得接受截断回复"})}}]}, {}
    client = TruncatedClient([])
    _install_client(monkeypatch, client)
    with pytest.raises(LLMResponseError):
        asyncio.run(check_intent(_request("统计记录数量"), config=_connection()))
    assert len(client.calls) == 1


def test_explicit_write_request_can_still_proceed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    prompt = "统计怀化横店电影城的记录数并写入汇总表"
    client = _FakeClient(
        [
            {
                "kind": "proceed",
                "summary": "统计后写入汇总表",
                "confirmedPrompt": prompt,
            }
        ]
    )
    _install_client(monkeypatch, client)

    response = asyncio.run(check_intent(_request(prompt), config=_connection()))

    assert response.kind == "proceed"
    assert len(client.calls) == 1
    assert is_explicit_read_only_query(prompt) is False


@pytest.mark.parametrize("prompt", ["你好", "你能做什么？"])
def test_non_task_conversation_is_answered_without_planning(
    monkeypatch: pytest.MonkeyPatch, prompt: str
) -> None:
    client = _FakeClient([{"kind": "conversation", "message": "你好！我可以帮你查询和分析已授权的工作表。"}])
    _install_client(monkeypatch, client)

    response = asyncio.run(check_intent(_request(prompt), config=_connection()))

    assert response.kind == "conversation"
    assert response.provider == "model"
    assert "你好" in response.message
    assert len(client.calls) == 1
    assert "scope" in json.loads(client.calls[0][1]["content"])
    assert "影院名称" in client.calls[0][1]["content"]
    assert "甲" not in client.calls[0][1]["content"]


@pytest.mark.parametrize(
    "prompt",
    [
        "怀化横店电影城有多少条记录",
        "查询怀化横店电影城的记录数",
        "看一下各影院的记录数量",
        "统计各影片的平均场次",
    ],
)
def test_generic_information_requests_are_read_only(prompt: str) -> None:
    assert is_explicit_read_only_query(prompt) is True

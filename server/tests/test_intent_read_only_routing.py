from __future__ import annotations

import asyncio
import json
from typing import Any

import pytest

from server.app import intent as intent_module
from server.app.intent import check_intent, is_explicit_read_only_query
from server.app.models import IntentCheckRequest


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
    monkeypatch.setattr(
        intent_module,
        "OpenAICompatibleClient",
        lambda *_args, **_kwargs: client,
    )


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

    response = asyncio.run(check_intent(_request(prompt), config=object()))

    assert response.kind == "tool_request"
    assert len(client.calls) == 2
    assert "明确只读查询，禁止返回 proceed" in client.calls[0][1]["content"]
    assert "必须返回 tool_request" in client.calls[1][1]["content"]


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

    response = asyncio.run(check_intent(_request(prompt), config=object()))

    assert response.kind == "clarification"
    assert response.provider == "local"
    assert "不会写入 Excel" in response.clarification.summary


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

    response = asyncio.run(check_intent(_request(prompt), config=object()))

    assert response.kind == "proceed"
    assert len(client.calls) == 1
    assert is_explicit_read_only_query(prompt) is False


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

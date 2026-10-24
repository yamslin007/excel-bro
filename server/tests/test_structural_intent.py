from __future__ import annotations

import pytest

from server.app.intent import check_intent
import asyncio
from server.tests.test_claude_sdk_agent import _FakeSdk, _install_fake, _connection
from server.app.models import IntentCheckRequest
from server.app.structural_intent import parse_bulk_worksheet_deletion


@pytest.mark.parametrize(
    "prompt",
    [
        "清除sheet-dc1外的所有子表",
        "删除除 sheet-dc1 外的所有工作表",
        "只保留 sheet-dc1，删除其他工作表",
        "保留 sheet-dc1，删除其他所有工作表",
        "僅保留『sheet-dc1』，刪除其餘工作表",
    ],
)
def test_parse_bulk_worksheet_deletion(prompt: str) -> None:
    parsed = parse_bulk_worksheet_deletion(prompt)

    assert parsed is not None
    assert parsed.keeper_name == "sheet-dc1"


@pytest.mark.parametrize(
    "prompt",
    [
        "清空 sheet-dc1 的数据",
        "删除 A2:A20",
        "按影片名称拆分工作表",
        "把 sheet-dc1 外的数据清空",
    ],
)
def test_parser_does_not_turn_other_requests_into_sheet_deletion(prompt: str) -> None:
    assert parse_bulk_worksheet_deletion(prompt) is None


def test_bulk_delete_is_resolved_by_sdk_not_old_prior_intent(monkeypatch, tmp_path) -> None:
    request = IntentCheckRequest.model_validate(
        {
            "prompt": "清除sheet-dc1外的所有子表",
            "scope": {
                "workbookName": "影院放映.xlsx",
                "sourceMode": "workbook",
                "selectionMode": "manual",
                "activeWorksheet": "sheet-dc1",
                "totalWorksheetCount": 3,
                "worksheetNames": ["sheet-dc1", "恐怖游轮", "痴迷"],
                "sheets": [
                    {
                        "name": "sheet-dc1",
                        "rowCount": 811,
                        "columnCount": 23,
                        "headers": ["影院名称", "影院编码", "影片名称"],
                    }
                ],
            },
            "conversation": [
                {"role": "user", "text": "按影片名称拆分工作表"},
                {"role": "assistant", "text": "准备生成拆分统计计划"},
            ],
            "priorIntent": {
                "confirmedPrompt": "按影片名称拆分工作表并按影院汇总"
            },
        }
    )

    fake = _FakeSdk(response='{"kind":"proceed","summary":"删除其他授权表","confirmedPrompt":"只保留 sheet-dc1"}')
    _install_fake(monkeypatch, tmp_path, fake)
    response = asyncio.run(check_intent(request, config=_connection()))
    assert response.kind == "proceed"
    assert response.provider == "model"
    assert response.confirmedPrompt == "只保留 sheet-dc1"
    assert request.prompt in fake.calls[0][0]

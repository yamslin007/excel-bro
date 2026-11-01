"""开发来源白名单：单一配置来源、开发页可用、未授权来源仍被拒。

背景（2026-10-28）：`wpsjs debug -p 3889` 的 portfinder 会把 Windows 保留端口
（EACCES）当成"不可用"静默跳过，实际把 WPS 开发页漂到 3987/3988，而后端白名单只有
3000/3889/3890 —— 于是任务窗格（POST）被 `FORBIDDEN_ORIGIN` 打回、`GET` 也读不到
响应，界面只剩 `Failed to fetch`。

这里钉住修复后的口径：
- 白名单派生自 `config/capabilities.json` 的 `devServers`，服务端不再自带第二份端口清单；
- 候选里的每个 WPS 开发来源都能读健康端点与只读发现端点；
- 候选外的来源仍然被拒（不做通配、不因"同机"放行）；
- Excel 桌面窗口守卫（`_review_layout_origin`）依旧只认 Excel 的四个来源。
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
from server.app import capabilities as capabilities_module
from server.app.main import ALLOWED_ORIGINS, EXCEL_ORIGINS, WPS_DEV_PORT_CANDIDATES, app
from server.app.models import ExcelReviewLayoutPrepareRequest
from server.app.excel_window_layout import path_hash

REPO_ROOT = Path(__file__).resolve().parents[2]
CAPABILITIES = json.loads(
    (REPO_ROOT / "config" / "capabilities.json").read_text(encoding="utf-8")
)

WPS_ORIGINS = [f"http://127.0.0.1:{port}" for port in WPS_DEV_PORT_CANDIDATES]
# 明确"不在候选里"的回环来源：用来证明拒绝不是"没测到"。
UNAUTHORIZED_ORIGIN = "http://127.0.0.1:65533"
SAVED_WPS_IDENTITY = "wps-path:c:/work/report.xlsx"


@pytest.fixture()
def client() -> TestClient:
    return TestClient(app)


def test_wps_origins_come_from_the_single_config_source() -> None:
    dev = CAPABILITIES["devServers"]
    assert WPS_DEV_PORT_CANDIDATES == dev["wpsPortCandidates"]
    assert dev["excelPort"] not in dev["wpsPortCandidates"]
    for port in dev["wpsPortCandidates"]:
        assert f"http://127.0.0.1:{port}" in ALLOWED_ORIGINS
        assert f"http://localhost:{port}" in ALLOWED_ORIGINS
        # 端口一旦与 Excel 撞车，两端会互相抢端口，启动器与配置都要拒绝
        assert f"http://127.0.0.1:{port}" not in EXCEL_ORIGINS


def test_excel_origins_keep_their_contract_position() -> None:
    """`_review_layout_origin` 只认 Excel 来源；顺序是契约，别让 WPS 端口混进去。"""
    assert ALLOWED_ORIGINS[: len(EXCEL_ORIGINS)] == EXCEL_ORIGINS
    assert EXCEL_ORIGINS == [
        "https://localhost:3000",
        "https://127.0.0.1:3000",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ]


@pytest.mark.parametrize("origin", WPS_ORIGINS)
def test_wps_dev_origin_can_read_health(client: TestClient, origin: str) -> None:
    response = client.get("/health", headers={"Origin": origin})
    assert response.status_code == 200
    # 缺了 ACAO，浏览器会把 200 也报成 Failed to fetch —— 这正是当初的现象
    assert response.headers["access-control-allow-origin"] == origin


@pytest.mark.parametrize("origin", WPS_ORIGINS)
def test_wps_dev_origin_can_reach_readonly_discovery(
    client: TestClient, origin: str
) -> None:
    response = client.get(
        "/api/wps/executions/pending",
        params={"workbookIdentity": SAVED_WPS_IDENTITY},
        headers={"Origin": origin},
    )
    assert response.status_code == 200
    payload = response.json()
    assert payload["workbookIdentity"] == SAVED_WPS_IDENTITY
    # 只读发现：没有未完成执行就是空列表，不创建任务、不写账本
    assert payload["candidates"] == []


def test_unknown_origin_is_still_rejected_for_state_changing_requests(
    client: TestClient,
) -> None:
    forbidden = client.post(
        "/api/__dev_origin_probe", headers={"Origin": UNAUTHORIZED_ORIGIN}
    )
    assert forbidden.status_code == 403
    assert forbidden.json()["detail"]["code"] == "FORBIDDEN_ORIGIN"

    # 同一个不存在的路由，来源被允许时走完守卫进入路由匹配（404），
    # 说明 403 的确来自来源判据，而不是"路由不存在"的副作用。
    allowed = client.post("/api/__dev_origin_probe", headers={"Origin": WPS_ORIGINS[-1]})
    assert allowed.status_code == 404


def test_unknown_origin_cannot_read_responses(client: TestClient) -> None:
    response = client.get("/health", headers={"Origin": UNAUTHORIZED_ORIGIN})
    assert response.status_code == 200
    assert "access-control-allow-origin" not in response.headers


def test_excel_layout_guard_still_refuses_wps_dev_origin(client: TestClient) -> None:
    payload = ExcelReviewLayoutPrepareRequest(
        workbookPathHash=path_hash("E:/Original.xlsx"),
        worksheetNameHash="f" * 64,
        copySheetName="EB_" + "b" * 20,
    ).model_dump()
    response = client.post(
        "/api/office/review-layout/prepare",
        json=payload,
        headers={"Origin": WPS_ORIGINS[-1], "Content-Type": "application/json"},
    )
    assert response.status_code == 403
    assert response.json()["detail"]["code"] == "LAYOUT_FORBIDDEN"


def test_pending_discovery_uses_the_isolated_test_ledger() -> None:
    """只读发现不能碰用户真实账本：conftest 把配置目录指到临时目录。"""
    root = Path(os.environ["EXCEL_BRO_CONFIG_DIR"]).resolve()
    ledger_path = main_module._folder_operation_ledger._path()
    assert ledger_path.is_relative_to(root), f"账本落在真实目录：{ledger_path}"


def test_capability_int_list_rejects_duplicates_and_non_ports(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    for bad in ([3987, 3987], [0], [True], []):
        monkeypatch.setattr(
            capabilities_module,
            "capabilities",
            lambda bad=bad: {"devServers": {"wpsPortCandidates": bad}},
        )
        with pytest.raises(RuntimeError):
            capabilities_module.capability_int_list("devServers", "wpsPortCandidates")

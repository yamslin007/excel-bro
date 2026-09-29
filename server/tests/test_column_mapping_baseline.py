"""阶段 G 基线：参数化列映射的跨宿主贯通用例（文件夹 / openpyxl 通道）。

对应 docs/DUAL_HOST_REFACTOR_PROPOSAL_2026-09-26.md 第 3 节。
文件夹模式是三条通道中唯一可全自动的，因此是回归网主力。

本文件最初只建立基线：已具备的底层（writeValues 落盘、回读、撤销、
外部改动拒绝）正常通过，尚未实现的参数化映射计算层用
xfail(strict=True) 锁定失败原因。阶段 4 实现 server/app/column_mapping.py
后那三条转为 XPASS，strict 如期报错，标记已移除（方案 §3.4）。

D→E 只是测试实例。映射表、源/目标列、工作表名全部来自参数，
运行时代码不得出现 D/E/Sheet1 或示例类别常量。
"""

from __future__ import annotations

import asyncio
import importlib
import json
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook, load_workbook

from server.app import intent as intent_module
from server.app.folder_workbooks import (
    FolderExecuteRequest,
    FolderSelection,
    FolderSnapshotRequest,
    create_folder_snapshot,
    execute_folder_plan,
    scan_folder,
    undo_folder_execution,
)
from server.app.intent import check_intent
from server.app.models import AnalysisPlan, FolderUndoRequest, IntentCheckRequest

# ---------------------------------------------------------------- 参数化夹具

SHEET_NAME = "订单"

# 源数据。第 5 行来源为空白单元格，用于验证"空白→未分类"映射策略；
# 它不是"跨宿主空串原样持久化"断言——openpyxl 保存后会把 "" 读回 None，
# 那个问题由 apps/wps-addin/src/wpsEmptyStringBaseline.test.ts 单独固定。
SOURCE_ROWS: list[tuple[str, str, int, str | None]] = [
    ("X001", "2026-01-05", 1200, "办公用品"),
    ("X002", "2026-01-06", 8600, "电子设备"),
    ("X003", "2026-01-07", 300, "办公用品"),
    ("X004", "2026-01-08", 450, None),
    ("X005", "2026-01-09", 2000, "差旅"),
]

MAPPING: dict[str, str] = {
    "办公用品": "行政",
    "电子设备": "资产",
    "差旅": "费用",
}
BLANK_LABEL = "未分类"

EXPECTED_TARGET = ["行政", "资产", "行政", "未分类", "费用"]


def _expected_for(mapping: dict[str, str], blank_label: str) -> list[str]:
    """从参数推导期望值，不硬编码结果。"""
    return [
        blank_label if category is None else mapping[category]
        for *_rest, category in SOURCE_ROWS
    ]


def _create_source(
    path: Path,
    *,
    source_sheet: str,
    target_sheet: str,
    source_column: str,
    target_column: str,
) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = source_sheet
    output = sheet if target_sheet == source_sheet else workbook.create_sheet(target_sheet)
    headers = {"A": "订单号", "B": "日期", "C": "金额", source_column: "类别"}
    for column, title in headers.items():
        sheet[f"{column}1"] = title
    output[f"{target_column}1"] = "类别分组"
    for offset, (order, date, amount, category) in enumerate(SOURCE_ROWS):
        row = offset + 2
        sheet[f"A{row}"] = order
        sheet[f"B{row}"] = date
        sheet[f"C{row}"] = amount
        if category is not None:
            sheet[f"{source_column}{row}"] = category
    workbook.save(path)
    workbook.close()


def _prepare(
    tmp_path: Path,
    *,
    source_sheet: str = SHEET_NAME,
    target_sheet: str = SHEET_NAME,
    source_column: str,
    target_column: str,
):
    source = tmp_path / "orders.xlsx"
    _create_source(
        source,
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        source_column=source_column,
        target_column=target_column,
    )
    catalog = scan_folder(tmp_path)
    selected_sheets = [source_sheet]
    if target_sheet != source_sheet:
        selected_sheets.append(target_sheet)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=selected_sheets)
            ],
        )
    )
    return source, catalog, snapshot


def _write_values_plan(
    snapshot, *, target_sheet: str = SHEET_NAME, target_column: str, values: list[str]
):
    first_row = 2
    last_row = first_row + len(values) - 1
    qualified_target = f"orders.xlsx › {target_sheet}"
    assert qualified_target in {sheet.name for sheet in snapshot.worksheets}
    plan = AnalysisPlan.model_validate(
        {
            "id": "plan-column-map",
            "title": "按映射写入分组列",
            "summary": f"把类别按映射表写入 {target_column} 列",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": qualified_target,
                    "range": f"{target_column}{first_row}:{target_column}{last_row}",
                    "values": [[value] for value in values],
                }
            ],
        }
    )
    return plan.model_copy(
        update={
            "sourceFingerprint": snapshot.sourceFingerprint,
            "sourceFingerprintSheets": snapshot.sourceFingerprintSheets,
        }
    )


def _read_column(
    path: Path, sheet_name: str, column: str, count: int
) -> list[object]:
    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook[sheet_name]
        return [sheet[f"{column}{row + 2}"].value for row in range(count)]
    finally:
        workbook.close()


# ------------------------------------------------- 已具备能力：应当正常通过

@pytest.mark.parametrize(
    ("source_sheet", "target_sheet", "source_column", "target_column"),
    [(SHEET_NAME, SHEET_NAME, "D", "E"), ("输入", "结果", "B", "H")],
    ids=["baseline-D-to-E", "cross-sheet-B-to-H"],
)
def test_a7_a9_write_and_undo_round_trip(
    tmp_path: Path,
    source_sheet: str,
    target_sheet: str,
    source_column: str,
    target_column: str,
) -> None:
    """A7 + A9：写入后持久化重读一致，撤销后目标回到执行前、源未被触碰。

    换列/换表变形只改参数，不增加分支实现——第二组真的从一张已授权表
    读到另一张已授权表，而非仅在同一表里换列。
    """
    expected = _expected_for(MAPPING, BLANK_LABEL)
    assert expected == EXPECTED_TARGET

    source, catalog, snapshot = _prepare(
        tmp_path,
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        source_column=source_column,
        target_column=target_column,
    )
    assert {sheet.name for sheet in snapshot.worksheets} == {
        f"orders.xlsx › {name}" for name in {source_sheet, target_sheet}
    }
    before_source = _read_column(source, source_sheet, source_column, len(SOURCE_ROWS))
    before_target = _read_column(source, target_sheet, target_column, len(SOURCE_ROWS))

    plan = _write_values_plan(
        snapshot, target_sheet=target_sheet, target_column=target_column, values=expected
    )
    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )
    assert result.undo is not None

    # A7：持久化重读后与计划一致。
    assert _read_column(source, target_sheet, target_column, len(expected)) == expected
    assert _read_column(source, source_sheet, source_column, len(expected)) == before_source

    # A9：撤销后目标回到执行前，源区域从未被触碰。
    undo_folder_execution(
        FolderUndoRequest(sessionId=result.undo.sessionId, undoId=result.undo.undoId)
    )
    assert _read_column(source, target_sheet, target_column, len(SOURCE_ROWS)) == before_target
    assert _read_column(source, source_sheet, source_column, len(SOURCE_ROWS)) == before_source


def test_a4_authorization_does_not_expand(tmp_path: Path) -> None:
    """A4：快照授权集合只含本轮明确选择的工作表。"""
    source = tmp_path / "orders.xlsx"
    _create_source(
        source,
        source_sheet=SHEET_NAME,
        target_sheet=SHEET_NAME,
        source_column="D",
        target_column="E",
    )
    workbook = load_workbook(source)
    workbook.create_sheet("未选中的表")
    workbook.save(source)
    workbook.close()

    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=[SHEET_NAME])
            ],
        )
    )
    # 文件夹通道用「文件 › 工作表」限定名，避免同名表跨文件歧义
    # （见 docs/ARCHITECTURE.md 的工作表授权边界一节）。
    authorized = {worksheet.name for worksheet in snapshot.worksheets}
    assert authorized == {f"orders.xlsx › {SHEET_NAME}"}
    assert not any("未选中的表" in name for name in authorized)


@pytest.mark.parametrize(
    ("changed_cell", "changed_value"),
    [("D2", "预览后来源变化"), ("E2", "预览后目标变化")],
    ids=["source-changed", "target-changed"],
)
def test_a12_source_or_target_changed_between_preview_and_execute(
    tmp_path: Path, changed_cell: str, changed_value: str
) -> None:
    """A12：预览与执行之间来源或目标被改动，必须拒绝旧计划。

    这是 A10（撤销前改动）之外的另一个空窗。当前由源指纹校验覆盖。
    """
    expected = _expected_for(MAPPING, BLANK_LABEL)
    source, catalog, snapshot = _prepare(tmp_path, source_column="D", target_column="E")
    plan = _write_values_plan(snapshot, target_column="E", values=expected)

    # 预览已生成，执行前第三方改动来源或目标区域。
    workbook = load_workbook(source)
    workbook[SHEET_NAME][changed_cell] = changed_value
    workbook.save(source)
    workbook.close()
    before_target = _read_column(source, SHEET_NAME, "E", len(SOURCE_ROWS))

    with pytest.raises(Exception) as excinfo:
        execute_folder_plan(
            FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
        )
    # 锁定拒绝原因与"预览后文件发生变化"相关，而不是任意报错。
    assert "发生变化" in str(excinfo.value)
    assert "重新" in str(excinfo.value)

    # 并且确实没有写入：保留第三方改动，不以旧计划覆盖目标。
    assert _read_column(source, SHEET_NAME, "E", len(SOURCE_ROWS)) == before_target


# ------------------------------------ 参数化映射计算层（阶段 4 已实现）
#
# 原先这三条用 xfail(strict=True, raises=_MissingColumnMappingCapability) 锁定
# "只允许模块缺席这一种失败"。server/app/column_mapping.py 落地后它们转为
# XPASS，strict 如期报错，标记随即移除（方案 §3.4）。
# 现在它们是普通用例：再失败就是真实回归。


def _column_mapping_module():
    return importlib.import_module("server.app.column_mapping")


@pytest.mark.parametrize(
    ("source_sheet", "target_sheet", "source_range", "target_range"),
    [
        (SHEET_NAME, SHEET_NAME, "D2:D6", "E2:E6"),
        ("输入", "结果", "B2:B6", "H2:H6"),
    ],
    ids=["same-sheet-D-to-E", "cross-sheet-B-to-H"],
)
def test_a3_column_mapping_plan_is_generated_from_parameters(
    source_sheet: str,
    target_sheet: str,
    source_range: str,
    target_range: str,
) -> None:
    """A3：本地确定性工具应能从映射参数生成单个 writeValues 动作。

    期望的共用入口：给定源区域值矩阵 + 映射表 + 空白/未匹配策略，
    返回目标矩阵与白名单动作。三条通道共用同一实现。
    """
    actions = _column_mapping_module().build_column_mapping_actions(
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        source_range=source_range,
        target_range=target_range,
        source_values=[[row[3]] for row in SOURCE_ROWS],
        mapping=MAPPING,
        blank_label=BLANK_LABEL,
    )
    assert len(actions) == 1
    assert actions[0]["type"] == "writeValues"
    assert actions[0]["sheet"] == target_sheet
    assert actions[0]["range"] == target_range
    assert actions[0]["values"] == [[value] for value in EXPECTED_TARGET]


def test_unmatched_source_value_does_not_get_guessed() -> None:
    """映射规则已给，但来源值未命中：拒绝猜测，不等同于未给映射规则。"""
    module = _column_mapping_module()
    with pytest.raises(module.UnmappedCategoryError):
        module.build_column_mapping_actions(
            source_sheet=SHEET_NAME,
            target_sheet=SHEET_NAME,
            source_range="D2:D6",
            target_range="E2:E6",
            source_values=[
                ["未在映射表中的类别"],
                *[[row[3]] for row in SOURCE_ROWS[1:]],
            ],
            mapping=MAPPING,
            blank_label=BLANK_LABEL,
        )


class _ClarificationClient:
    def __init__(self) -> None:
        self.calls: list[list[dict[str, Any]]] = []

    async def __aenter__(self) -> "_ClarificationClient":
        return self

    async def __aexit__(self, exc_type, exc, tb) -> bool:
        return False

    async def chat_completions(
        self, *, messages: list[dict[str, Any]], **_kwargs: Any
    ) -> tuple[dict[str, Any], dict[str, int]]:
        self.calls.append(messages)
        clarification = {
            "kind": "clarification",
            "clarification": {
                "id": "missing-map",
                "summary": "需要类别与分组之间的映射规则",
                "question": "请补充类别映射，以及空白和未匹配类别的处理规则？",
                "reason": "这些规则决定写入的具体值，不能猜测。",
                "options": [
                    {
                        "id": "supply-rules",
                        "label": "补充规则",
                        "description": "先说明映射，再生成写入预览。",
                        "resolution": "先不写入，待我补充映射、空白和未匹配规则。",
                        "action": "resolve",
                    },
                    {
                        "id": "cancel-request",
                        "label": "取消写入",
                        "description": "这次不修改工作簿。",
                        "resolution": "取消本次写入，不生成计划。",
                        "resolutionMode": "replace",
                        "action": "resolve",
                    },
                ],
            },
        }
        return (
            {"choices": [{"message": {"content": json.dumps(clarification, ensure_ascii=False)}}]},
            {},
        )


def test_b1_b3_missing_mapping_rule_routes_to_clarification_without_data_rows(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """规则本身缺席时走真实意图入口；桩只固定模型的语义判断。"""
    secret_row = "原始行不应进入意图模型-仅作泄漏哨兵"
    request = IntentCheckRequest.model_validate(
        {
            "prompt": "按 D 列类别映射，把分组结果写到 E 列",
            "scope": {
                "workbookName": "orders.xlsx",
                "sourceMode": "folder",
                "selectionMode": "folder",
                "activeWorksheet": SHEET_NAME,
                "totalWorksheetCount": 1,
                "worksheetNames": [SHEET_NAME],
                "sheets": [
                    {
                        "name": SHEET_NAME,
                        "usedRange": "A1:E6",
                        "rowCount": 6,
                        "columnCount": 5,
                        "headers": ["订单号", "日期", "金额", "类别", "类别分组"],
                        "dataRows": [[secret_row]],
                    }
                ],
            },
        }
    )
    client = _ClarificationClient()
    monkeypatch.setattr(
        intent_module, "OpenAICompatibleClient", lambda *_args, **_kwargs: client
    )

    response = asyncio.run(check_intent(request, config=object()))

    assert response.kind == "clarification"
    assert "映射" in response.clarification.question
    assert len(client.calls) == 1
    outbound = client.calls[0][1]["content"]
    assert "数据范围结构摘要（不含数据行）" in outbound
    assert "类别分组" in outbound
    assert secret_row not in outbound
    assert "dataRows" not in outbound

"""阶段 G 基线：参数化列映射的跨宿主贯通用例（文件夹 / openpyxl 通道）。

对应 docs/DUAL_HOST_REFACTOR_PROPOSAL_2026-09-26.md 第 3 节。
文件夹模式是三条通道中唯一可全自动的，因此是回归网主力。

本文件只建立基线，不实现能力、不改生产代码：
- 已具备的底层（writeValues 落盘、回读、撤销、外部改动拒绝）→ 正常通过
- 尚未实现的参数化映射计算层 → xfail(strict=True) 锁定失败原因

strict=True 的作用是：一旦能力实现导致 XPASS，pytest 会报错而不是静默转绿，
迫使重新审查并移除标记。对应方案 §3.4。

D→E 只是测试实例。映射表、源/目标列、工作表名全部来自参数，
运行时代码不得出现 D/E/Sheet1 或示例类别常量。
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from server.app.folder_workbooks import (
    FolderExecuteRequest,
    FolderSelection,
    FolderSnapshotRequest,
    create_folder_snapshot,
    execute_folder_plan,
    scan_folder,
    undo_folder_execution,
)
from server.app.models import AnalysisPlan, FolderUndoRequest

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


def _create_source(path: Path, *, source_column: str, target_column: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME
    headers = {"A": "订单号", "B": "日期", "C": "金额", source_column: "类别"}
    for column, title in headers.items():
        sheet[f"{column}1"] = title
    sheet[f"{target_column}1"] = "类别分组"
    for offset, (order, date, amount, category) in enumerate(SOURCE_ROWS):
        row = offset + 2
        sheet[f"A{row}"] = order
        sheet[f"B{row}"] = date
        sheet[f"C{row}"] = amount
        if category is not None:
            sheet[f"{source_column}{row}"] = category
    workbook.save(path)
    workbook.close()


def _prepare(tmp_path: Path, *, source_column: str, target_column: str):
    source = tmp_path / "orders.xlsx"
    _create_source(source, source_column=source_column, target_column=target_column)
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=[SHEET_NAME])
            ],
        )
    )
    return source, catalog, snapshot


def _write_values_plan(snapshot, *, target_column: str, values: list[str]):
    first_row = 2
    last_row = first_row + len(values) - 1
    plan = AnalysisPlan.model_validate(
        {
            "id": "plan-column-map",
            "title": "按映射写入分组列",
            "summary": f"把类别按映射表写入 {target_column} 列",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
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


def _read_column(path: Path, column: str, count: int) -> list[object]:
    workbook = load_workbook(path, data_only=True)
    try:
        sheet = workbook[SHEET_NAME]
        return [sheet[f"{column}{row + 2}"].value for row in range(count)]
    finally:
        workbook.close()


# ------------------------------------------------- 已具备能力：应当正常通过

@pytest.mark.parametrize(
    ("source_column", "target_column"),
    [("D", "E"), ("B", "H")],
    ids=["baseline-D-to-E", "variant-B-to-H"],
)
def test_a7_a9_write_and_undo_round_trip(
    tmp_path: Path, source_column: str, target_column: str
) -> None:
    """A7 + A9：写入后持久化重读一致，撤销后目标回到执行前、源未被触碰。

    换列/换表变形只改参数，不增加分支实现——这能防止"测试通过，
    产品却只会 D→E"。
    """
    expected = _expected_for(MAPPING, BLANK_LABEL)
    assert expected == EXPECTED_TARGET

    source, catalog, snapshot = _prepare(
        tmp_path, source_column=source_column, target_column=target_column
    )
    before_source = _read_column(source, source_column, len(SOURCE_ROWS))
    before_target = _read_column(source, target_column, len(SOURCE_ROWS))

    plan = _write_values_plan(snapshot, target_column=target_column, values=expected)
    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )
    assert result.undo is not None

    # A7：持久化重读后与计划一致。
    assert _read_column(source, target_column, len(expected)) == expected

    # A9：撤销后目标回到执行前，源区域从未被触碰。
    undo_folder_execution(
        FolderUndoRequest(sessionId=result.undo.sessionId, undoId=result.undo.undoId)
    )
    assert _read_column(source, target_column, len(SOURCE_ROWS)) == before_target
    assert _read_column(source, source_column, len(SOURCE_ROWS)) == before_source


def test_a4_authorization_does_not_expand(tmp_path: Path) -> None:
    """A4：快照授权集合只含本轮明确选择的工作表。"""
    source = tmp_path / "orders.xlsx"
    _create_source(source, source_column="D", target_column="E")
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


def test_a12_target_changed_between_preview_and_execute(tmp_path: Path) -> None:
    """A12：预览与执行之间源数据被改动，执行必须拒绝而不是照旧写入。

    这是 A10（撤销前改动）之外的另一个空窗。当前由源指纹校验覆盖。
    """
    expected = _expected_for(MAPPING, BLANK_LABEL)
    source, catalog, snapshot = _prepare(tmp_path, source_column="D", target_column="E")
    plan = _write_values_plan(snapshot, target_column="E", values=expected)

    # 预览已生成，执行前第三方改动源区域。
    workbook = load_workbook(source)
    workbook[SHEET_NAME]["D2"] = "预览后被改成别的类别"
    workbook.save(source)
    workbook.close()

    with pytest.raises(Exception) as excinfo:
        execute_folder_plan(
            FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
        )
    # 锁定拒绝原因与"预览后源发生变化"相关，而不是任意报错。
    assert "发生变化" in str(excinfo.value)
    assert "重新" in str(excinfo.value)

    # 并且确实没有写入：目标列仍为空。
    assert _read_column(source, "E", len(SOURCE_ROWS)) == [None] * len(SOURCE_ROWS)


# ------------------------------------ 尚未实现：xfail(strict) 锁定失败原因

@pytest.mark.xfail(
    strict=True,
    raises=ImportError,
    reason=(
        "阶段 G 基线：参数化列映射的确定性计算层尚未实现。"
        "阶段 4 实现 build_column_mapping_actions 后本条会 XPASS，"
        "strict=True 会使其报错，届时移除此标记。"
    ),
)
def test_a3_column_mapping_plan_is_generated_from_parameters() -> None:
    """A3：本地确定性工具应能从映射参数生成单个 writeValues 动作。

    期望的共用入口：给定源区域值矩阵 + 映射表 + 空白/未匹配策略，
    返回目标矩阵与白名单动作。三条通道共用同一实现。
    """
    from server.app.column_mapping import (  # noqa: PLC0415
        build_column_mapping_actions,
    )

    actions = build_column_mapping_actions(
        sheet=SHEET_NAME,
        source_range="D2:D6",
        target_range="E2:E6",
        source_values=[[row[3]] for row in SOURCE_ROWS],
        mapping=MAPPING,
        blank_label=BLANK_LABEL,
    )
    assert len(actions) == 1
    assert actions[0]["type"] == "writeValues"
    assert actions[0]["range"] == "E2:E6"
    assert actions[0]["values"] == [[value] for value in EXPECTED_TARGET]


@pytest.mark.xfail(
    strict=True,
    raises=ImportError,
    reason=(
        "阶段 G 基线：未匹配类别应触发澄清而不是猜测，判定入口尚未实现。"
        "对应方案 §3.2.1 的 B1-B3。"
    ),
)
def test_b1_unmatched_category_requires_clarification() -> None:
    """B1-B3：映射表未覆盖的类别不得猜测，必须要求澄清且不生成写入动作。"""
    from server.app.column_mapping import (  # noqa: PLC0415
        build_column_mapping_actions,
        UnmappedCategoryError,
    )

    with pytest.raises(UnmappedCategoryError):
        build_column_mapping_actions(
            sheet=SHEET_NAME,
            source_range="D2:D6",
            target_range="E2:E6",
            source_values=[["未在映射表中的类别"]],
            mapping=MAPPING,
            blank_label=BLANK_LABEL,
        )

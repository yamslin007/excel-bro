from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
import pytest

from server.app.folder_workbooks import (
    FolderExecuteRequest,
    FolderSelection,
    FolderSnapshotRequest,
    create_folder_snapshot,
    execute_folder_plan,
    scan_folder,
)
from server.app.models import AnalysisPlan


def _create_source(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "得分"
    sheet.append(["人员", "得分"])
    sheet.append(["嘟嘟嘟", 33])
    sheet.append(["阿里", 44])
    workbook.create_sheet("说明")
    workbook.save(path)
    workbook.close()


def test_folder_scan_and_selected_snapshot(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)

    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])
            ],
        )
    )

    assert len(catalog.files) == 1
    assert snapshot.name == f"文件夹：{tmp_path.name}"
    assert snapshot.worksheets[0].sourceFile == "scores.xlsx"
    assert snapshot.worksheets[0].sourceSheet == "得分"
    assert snapshot.worksheets[0].dataRows[0] == ["嘟嘟嘟", 33]


def test_folder_plan_writes_output_workbook(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    plan = AnalysisPlan.model_validate(
        {
            "id": "folder-output",
            "title": "生成汇总",
            "summary": "写入独立结果文件",
            "actions": [
                {"type": "createWorksheet", "sheet": "汇总"},
                {
                    "type": "writeTable",
                    "sheet": "汇总",
                    "startCell": "A1",
                    "headers": ["人员", "得分"],
                    "rows": [["嘟嘟嘟", 33]],
                },
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    output = load_workbook(tmp_path / "Excel Bro 结果.xlsx", data_only=True)
    assert result.filesModified == ["Excel Bro 结果.xlsx"]
    assert result.backups == []
    assert len(result.actionResults) == 2
    assert result.verification.passed is True
    assert all(check.passed for check in result.verification.checks)
    assert output["汇总"]["B2"].value == 33
    output.close()


def test_existing_file_is_backed_up_before_write(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])
            ],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "folder-edit",
            "title": "修正得分",
            "summary": "修改已有文件",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": "scores.xlsx › 得分",
                    "range": "B2",
                    "values": [[35]],
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    updated = load_workbook(source, data_only=True)
    assert updated["得分"]["B2"].value == 35
    updated.close()
    assert result.filesModified == ["scores.xlsx"]
    assert len(result.backups) == 1
    assert (tmp_path / result.backups[0]).exists()


def test_plan_cannot_modify_an_unselected_sheet(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])
            ],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "unselected-edit",
            "title": "越权修改",
            "summary": "不应执行",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": "scores.xlsx › 说明",
                    "range": "A1",
                    "values": [["不允许"]],
                }
            ],
        }
    )

    with pytest.raises(ValueError, match="未选择"):
        execute_folder_plan(
            FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
        )


def test_folder_verification_reports_an_expected_value_mismatch(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    plan = AnalysisPlan.model_validate(
        {
            "id": "failed-verification",
            "title": "验证失败示例",
            "summary": "执行值与验收值不同",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": "结果",
                    "range": "A1",
                    "values": [[1]],
                }
            ],
            "acceptanceCriteria": [
                {
                    "type": "rangeEquals",
                    "sheet": "结果",
                    "range": "A1",
                    "expected": [[2]],
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    assert result.verification.passed is False
    assert result.verification.checks[0].actual == [[1]]
    assert "与预期不一致" in result.verification.checks[0].message


def test_folder_mode_writes_formulas_formats_and_clears_ranges(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])
            ],
        )
    )
    target = snapshot.worksheets[0].name
    plan = AnalysisPlan.model_validate(
        {
            "id": "advanced-edit",
            "title": "公式与格式",
            "summary": "写入公式、设置格式并清空旧值",
            "actions": [
                {
                    "type": "writeFormulas",
                    "sheet": target,
                    "range": "C2",
                    "formulas": [["=SUM(B2:B4)"]],
                },
                {
                    "type": "setNumberFormat",
                    "sheet": target,
                    "range": "C2",
                    "formatCode": "0.00",
                },
                {
                    "type": "clearRange",
                    "sheet": target,
                    "range": "B2",
                    "applyTo": "contents",
                },
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    updated = load_workbook(source, data_only=False)
    assert updated["得分"]["C2"].value == "=SUM(B2:B4)"
    assert updated["得分"]["C2"].number_format == "0.00"
    assert updated["得分"]["B2"].value is None
    updated.close()
    assert result.verification.passed is True


def test_folder_mode_deletes_a_selected_sheet_and_verifies_it(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=["说明"])
            ],
        )
    )
    target = snapshot.worksheets[0].name
    plan = AnalysisPlan.model_validate(
        {
            "id": "delete-sheet",
            "title": "删除说明",
            "summary": "删除已选说明工作表",
            "actions": [{"type": "deleteWorksheet", "sheet": target}],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    updated = load_workbook(source)
    assert "说明" not in updated.sheetnames
    updated.close()
    assert result.verification.passed is True


def test_folder_mode_splits_full_table_and_calculates_group_ratios(
    tmp_path: Path,
) -> None:
    source = tmp_path / "cinema.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "sheet-dc1"
    sheet.append(["hunan(制表日期:2026-07-25)", None, None, None])
    sheet.append([None, None, None, None])
    sheet.append(["影院名称", "影院编码", "影厅", "影片名称"])
    sheet.append(["怀化横店电影城", "43122301", "1号厅", "恐怖游轮"])
    sheet.append(["怀化横店电影城", "43122301", "2号厅", "恐怖游轮"])
    sheet.append(["怀化横店电影城", "43122301", "3号厅", "功夫女足"])
    sheet.append(["广州影院", "44010001", "1号厅", "恐怖游轮"])
    sheet.append(["长沙影院", "43010001", "1号厅", "三国第一部：争洛阳"])
    workbook.save(source)
    workbook.close()

    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(
                    fileId=catalog.files[0].id, sheets=["sheet-dc1"]
                )
            ],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "split-cinema",
            "title": "按影片拆分",
            "summary": "按影院统计影片记录数和占比",
            "actions": [
                {
                    "type": "splitGroupAggregate",
                    "sheet": snapshot.worksheets[0].name,
                    "splitBy": "影片名称",
                    "groupBy": ["影院名称", "影院编码"],
                    "metrics": [
                        {
                            "operation": "countRows",
                            "outputName": "计数项:影厅",
                            "ratioOutputName": "占比",
                        }
                    ],
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    updated = load_workbook(source, data_only=True)
    assert "恐怖游轮" in updated.sheetnames
    assert "功夫女足" in updated.sheetnames
    assert "三国第一部 争洛阳" in updated.sheetnames
    horror = updated["恐怖游轮"]
    assert [cell.value for cell in horror[1]] == [
        "影院名称",
        "影院编码",
        "影片名称",
        "计数项:影厅",
        "占比",
    ]
    assert [cell.value for cell in horror[2]] == [
        "怀化横店电影城",
        "43122301",
        "恐怖游轮",
        2,
        pytest.approx(2 / 3),
    ]
    assert horror["E2"].number_format == "0.00%"
    updated.close()
    assert result.verification.passed is True

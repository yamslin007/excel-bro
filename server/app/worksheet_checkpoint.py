"""Fixed native worksheet deletion adapter; no generated host scripts.

Restore copies deleted sheets back into the open original; open_copy is a separate fallback.
Neither action replaces or closes the original workbook.
A complete archive requires selected-sheet authorization or explicit backup-only consent.
"""
from __future__ import annotations

import hashlib
import json
import ntpath
import os
import re
import secrets
import shutil
import threading
import time
import zipfile
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote, urlparse
from xml.etree import ElementTree
from xml.parsers import expat

from .app_data import application_data_root
from .folder_mutex import FolderWriteBusyError, folder_write_guard
from .capabilities import capability_int, capabilities
from .models import (WorksheetCheckpointListRequest, WorksheetCheckpointListResult,
                     WorksheetCheckpointSummary, WorksheetDeleteCheckpointRequest,
                     WorksheetDeleteCheckpointResult, WorksheetRestoreResult)
from .worksheet_restore import RestoreConflict, restore_in_place

_LOCK = threading.Lock()


class CheckpointPersistenceError(OSError):
    """Persistence failed; the caller must not replay a native operation."""


def canonical_path(value: str) -> str:
    if value.startswith("url:"):
        value = value[4:]
    if value.startswith("file:"):
        parsed = urlparse(value)
        if parsed.netloc not in ("", "localhost"):
            raise ValueError("只支持本机已保存的 XLSX 工作簿。")
        value = unquote(parsed.path).lstrip("/")
    if not re.match(r"^[a-zA-Z]:[\\/]", value) or not value.lower().endswith(".xlsx"):
        raise ValueError("请先将原表保存为本机 XLSX 文件，再执行工作表删除。")
    return ntpath.normcase(ntpath.normpath(value))


def _folder(operation_id: str) -> Path:
    if not re.fullmatch(r"[a-f0-9]{32}", operation_id):
        raise ValueError("检查点标识无效。")
    root = application_data_root().resolve() / "worksheet-checkpoints"
    root.mkdir(parents=True, exist_ok=True)
    folder = root / operation_id
    if root.is_symlink() or folder.is_symlink() or folder.resolve().parent != root.resolve():
        raise ValueError("检查点目录不可信。")
    return folder


def _write(folder: Path, record: dict) -> None:
    temp = folder / "state.tmp"
    with temp.open("w", encoding="utf-8") as stream:
        json.dump(record, stream, ensure_ascii=False)
        stream.flush()
        os.fsync(stream.fileno())
    # Retry the atomic replacement only, never the surrounding delete/copy.
    attempts = capability_int("worksheetCheckpoint", "stateWriteAttempts")
    delay = capability_int("worksheetCheckpoint", "stateWriteRetryDelayMs") / 1000
    for attempt in range(attempts):
        try:
            temp.replace(folder / "state.json")
            return
        except PermissionError as exc:
            if attempt + 1 == attempts:
                raise CheckpointPersistenceError(
                    "检查点记录暂时无法保存，操作结果尚未确认。请查询状态并核对工作表，"
                    "不要重复执行删除；如已建立执行前副本，可用它核对。"
                ) from exc
            time.sleep(delay)


def _read(operation_id: str, origin: str) -> tuple[Path, dict]:
    folder = _folder(operation_id)
    path = folder / "state.json"
    if path.is_symlink():
        raise ValueError("检查点记录不可信。")
    record = json.loads(path.read_text(encoding="utf-8"))
    if record["origin"] != origin:
        raise ValueError("检查点不属于当前加载项。")
    return folder, record


def _result(record: dict) -> WorksheetDeleteCheckpointResult:
    result = WorksheetDeleteCheckpointResult.model_validate(record["result"])
    if record.get("restorePendingSheet"):
        result.restore = _uncertain_restore(record)
    return result


def _uncertain_restore(record: dict) -> WorksheetRestoreResult:
    pending = record["restorePendingSheet"]
    deleted = record["result"]["deletedSheets"]
    if not isinstance(pending, str) or pending not in deleted:
        raise ValueError("恢复在途记录无效，请核对执行前副本。")
    stored = _stored_restore(record)
    restored = list(stored.restoredSheets) if stored else []
    return WorksheetRestoreResult(
        operationId=record["result"]["operationId"], status="partial",
        restoredSheets=restored,
        failedSheets=[name for name in deleted if name not in restored and name != pending],
        unverifiedSheets=[pending],
        message=f"「{pending}」的恢复结果未留下完整回执，请核对当前工作表。"
                "不会自动覆盖或重复复制同名表；可从备份与恢复打开执行前副本核对。"
    )


def _stored_restore(record: dict) -> WorksheetRestoreResult | None:
    """记录里最近一次原地撤销的结果；旧记录没有这一项，就是"还没撤销过"。"""
    raw = record["result"].get("restore")
    return WorksheetRestoreResult.model_validate(raw) if isinstance(raw, dict) else None


def _remember_restore(record: dict, payload: dict) -> None:
    record["result"]["restore"] = payload


def _checkpoints_root() -> Path:
    return application_data_root().resolve() / "worksheet-checkpoints"


def _normalise_crashed(record: dict) -> bool:
    """Correct a receipt left behind by a stopped process; never replays a write."""
    status = record["result"]["status"]
    if status == "prepared":
        record["result"]["status"] = "failed"
        record["result"]["message"] = "检查点准备阶段中断，未开始删除。请重新发起任务。"
        return True
    if status == "executing":
        record["result"]["status"] = "partial_or_unknown"
        record["result"]["message"] = "上次删除未留下完整完成回执，请打开恢复副本核对；不会自动续删。"
        return True
    return False


@contextmanager
def _optional_write_guard(root: Path):
    """Listing is read-only; a busy writer must not turn a lookup into a failure."""
    try:
        with folder_write_guard(root):
            yield True
    except FolderWriteBusyError:
        yield False


def _validate_archive(path: Path, names: list[str]) -> str:
    if path.is_symlink() or not path.is_file() or path.stat().st_size > capability_int("worksheetCheckpoint", "maxBytes"):
        raise ValueError("完整备份未生成或超过容量限制，未开始删除。")
    try:
        archive = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as exc:
        # 损坏的归档是"这份备份不能用"，不是"未知宿主错误"：要给出可读原因，
        # 否则调用方只会看到笼统的"尚未确认"，然后去重试一个永远失败的操作。
        raise ValueError("检查点文件校验失败。") from exc
    with archive:
        entries = archive.infolist()
        if (len(entries) > capability_int("worksheetCheckpoint", "maxZipEntries")
            or sum(item.file_size for item in entries) > capability_int("worksheetCheckpoint", "maxExpandedBytes")
            or len({item.filename for item in entries}) != len(entries)):
            raise ValueError("检查点文件结构或容量不受支持。")
        if any("vbaproject" in item.filename.lower() or "macrosheets/" in item.filename.lower() for item in entries):
            raise ValueError("本检查点通道不打开含宏的工作簿。")
        if archive.testzip() is not None:
            raise ValueError("检查点文件校验失败。")
        xml = archive.read("xl/workbook.xml")
        parser = expat.ParserCreate()
        def reject_dtd(*_):
            raise ValueError("检查点 XML 不接受 DTD 或实体声明。")
        parser.StartDoctypeDeclHandler = reject_dtd
        parser.Parse(xml, True)
        tree = ElementTree.fromstring(xml)
        actual = [node.attrib["name"] for node in tree.findall("{*}sheets/{*}sheet")]
        if actual != names:
            raise ValueError("备份中的工作表与执行现场不一致，未开始删除。")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _sheets(book):
    sheets = [book.Worksheets.Item(i) for i in range(1, int(book.Worksheets.Count) + 1)]
    if int(book.Sheets.Count) != len(sheets):
        raise ValueError("工作簿包含图表表或其他非工作表对象，本批量删除通道暂不支持。")
    return sheets


def _identity(book, request, expected_names):
    if (canonical_path(str(book.FullName)) != canonical_path(request.workbookPath)
        or not str(book.Path) or bool(book.ReadOnly) or bool(book.ProtectStructure)
        or int(book.FileFormat) != 51):
        raise ValueError("工作簿身份、格式或保护状态变化，未继续删除。")
    names = [str(sheet.Name) for sheet in _sheets(book)]
    if names != expected_names:
        raise ValueError("工作表列表发生变化，未继续删除。")


class NativeWorksheetCheckpoint:
    @contextmanager
    def workbook(self, request):
        # Attach only: never launch an Application or reopen the original file.
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            target = canonical_path(request.workbookPath)
            if request.host == "excel":
                from .excel_window_layout import WindowsExcelWindows
                matches = [self._excel_workbook(WindowsExcelWindows(), target)]
            else:
                app = win32com.client.GetActiveObject("ket.Application")
                if not request.hostWindowId or int(app.Hwnd) != request.hostWindowId:
                    raise ValueError("无法确认当前 WPS 实例，请在原表中重新发起。")
                matches = [book for book in app.Workbooks if str(book.Path)
                           and ntpath.normcase(ntpath.normpath(str(book.FullName))) == target]
            if len(matches) != 1:
                raise ValueError("未找到唯一的原工作簿，未执行删除。")
            yield matches[0]
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _excel_workbook(native, target):
        def identify(window):
            # Read identity from, and retain, the SAME Workbook COM object.
            # Do not describe a window then call AccessibleObjectFromWindow again.
            book = window.ActiveSheet.Parent
            path = ntpath.normcase(ntpath.normpath(str(book.FullName))) if str(book.Path) else ""
            return path, book

        attempts = capability_int("worksheetCheckpoint", "nativeReadAttempts")
        for attempt in range(attempts):
            candidates = []
            unavailable = False
            for key in native.keys():
                if not native.gui.IsWindowVisible(key.hwnd):
                    continue
                if native._is_excel(key.hwnd) != key:
                    raise ValueError("原工作簿窗口身份发生变化。")
                children = []
                native.gui.EnumChildWindows(key.hwnd, lambda h, _: children.append(h)
                    if native.gui.GetClassName(h) == "EXCEL7" else None, None)
                observations = [native._document_window(child, identify) for child in children]
                if any(item is None for item in observations):
                    unavailable = True
                    continue
                if not observations:
                    continue
                if len({item[0] for item in observations}) != 1:
                    raise ValueError("无法唯一定位原工作簿窗口，未执行操作。")
                if observations[0][0] == target:
                    candidates.append((key, observations[0][1]))
            if len(candidates) > 1:
                raise ValueError("无法唯一定位原工作簿窗口，未执行操作。")
            if unavailable:
                # Give the host time to leave its transient busy state. Retrying
                # back-to-back can repeat the same E_FAIL. No mutation is retried.
                if attempt + 1 < attempts:
                    time.sleep(capability_int("worksheetCheckpoint", "nativeReadRetryDelayMs") / 1000)
                continue
            if not candidates:
                raise ValueError("未找到唯一的原工作簿，未执行操作。")
            key, retained = candidates[0]
            if native._is_excel(key.hwnd) != key:
                raise ValueError("原工作簿窗口身份发生变化。")
            if canonical_path(str(retained.FullName)) != target:
                raise ValueError("原工作簿身份发生变化。")
            return retained
        raise ValueError("原工作簿对象暂不可读取，请稍后重试；本次未执行操作。")

    def open_copy(self, request, path):
        with self.workbook(request) as original:
            app = original.Application
            # Open the archive derivative with links/events/macros disabled.
            # Restore application settings even if Open throws. Original stays open.
            security, events = app.AutomationSecurity, app.EnableEvents
            try:
                app.AutomationSecurity = 3
                app.EnableEvents = False
                opened = app.Workbooks.Open(str(path), 0, False)
                if canonical_path(str(opened.FullName)) != canonical_path(str(path)):
                    raise ValueError("恢复副本打开结果尚未确认。")
            finally:
                try:
                    app.EnableEvents = events
                finally:
                    app.AutomationSecurity = security

    @contextmanager
    def restore_session(self, request, archive):
        """挂到原工作簿，**只读**打开归档，供原地撤销使用。

        与 `open_copy` 的区别：那边把副本作为新窗口交还给用户；这里归档只作为
        "内容来源"，用完即关且不保存。原工作簿全程不关闭、不保存、不换对象，
        用户后续编辑因此不会被"整簿覆盖"抹掉。
        """
        import pythoncom
        pythoncom.CoInitialize()
        opened = None
        try:
            with self.workbook(request) as original:
                app = original.Application
                security, events, alerts = app.AutomationSecurity, app.EnableEvents, app.DisplayAlerts
                try:
                    app.AutomationSecurity = 3
                    app.EnableEvents = False
                    app.DisplayAlerts = False
                    # UpdateLinks=0, ReadOnly=True：归档是证据，不是工作簿。
                    opened = app.Workbooks.Open(str(archive), 0, True)
                    if canonical_path(str(opened.FullName)) != canonical_path(str(archive)):
                        raise ValueError("无法只读打开本次的执行前归档，未执行撤销。")
                    yield original, opened
                finally:
                    try:
                        if opened is not None:
                            opened.Close(SaveChanges=False)
                    finally:
                        try:
                            app.DisplayAlerts = alerts
                        finally:
                            try:
                                app.EnableEvents = events
                            finally:
                                app.AutomationSecurity = security
        finally:
            pythoncom.CoUninitialize()


class WorksheetCheckpointService:
    def __init__(self, native=None):
        self.native = native or NativeWorksheetCheckpoint()

    def status(self, operation_id, origin):
        with _LOCK, folder_write_guard(application_data_root() / "worksheet-checkpoints"):
            folder, record = _read(operation_id, origin)
            # Live operations hold the same lock. Seeing executing here means
            # the previous process stopped before a durable completion receipt.
            if _normalise_crashed(record):
                _write(folder, record)
            return _result(record)

    def list_for_workbook(self, request: WorksheetCheckpointListRequest, origin: str):
        """Recovery points for ONE workbook identity; attribution comes from saved metadata.

        Legacy records only hold the request metadata (workbook path, host). That is
        enough to show them under the right workbook without guessing a conversation.
        """
        target = canonical_path(request.workbookPath)
        root = _checkpoints_root()
        rows = []
        if root.is_symlink() or not root.is_dir():
            return WorksheetCheckpointListResult(checkpoints=[])
        with _LOCK, _optional_write_guard(root) as durable:
            for folder in sorted(root.iterdir()):
                if not re.fullmatch(r"[a-f0-9]{32}", folder.name):
                    continue
                path = folder / "state.json"
                if folder.is_symlink() or path.is_symlink() or not path.is_file():
                    continue
                try:
                    record = json.loads(path.read_text(encoding="utf-8"))
                    saved = record.get("request") or {}
                    if record.get("origin") != origin or saved.get("host") != request.host:
                        continue
                    if canonical_path(str(saved.get("workbookPath", ""))) != target:
                        continue
                    if durable and _normalise_crashed(record):
                        _write(folder, record)
                    result = _result(record)
                except (OSError, ValueError, KeyError, TypeError):
                    # A damaged record must not hide the healthy ones next to it.
                    continue
                rows.append(WorksheetCheckpointSummary(
                    operationId=result.operationId,
                    status=result.status,
                    deletedSheets=result.deletedSheets,
                    remainingSheets=result.remainingSheets,
                    recoveryAvailable=result.recoveryAvailable,
                    message=result.message,
                    createdAt=record.get("createdAt") if isinstance(record.get("createdAt"), str) else None,
                    planTitle=(saved.get("plan") or {}).get("title") if isinstance(saved.get("plan"), dict) else None,
                    restoreStatus=result.restore.status if result.restore else None,
                ))
        rows.sort(key=lambda item: (item.createdAt or "", item.operationId), reverse=True)
        return WorksheetCheckpointListResult(checkpoints=rows)

    def execute(self, request: WorksheetDeleteCheckpointRequest, origin: str):
        with _LOCK, folder_write_guard(application_data_root() / "worksheet-checkpoints"):
            return self._execute_locked(request, origin)

    def confirm_backup(self, consent, origin):
        with _LOCK, folder_write_guard(application_data_root() / "worksheet-checkpoints"):
            _, record = _read(consent.operationId, origin)
            result = _result(record)
            if result.status != "needs_backup_consent":
                return result  # Repeated/lost responses never replay a native operation.
            if (len(consent.backupSheets) != len(set(consent.backupSheets))
                or consent.backupSheets != result.backupSheets):
                raise ValueError("备份许可与待确认的工作表不一致，未执行。")
            return self._execute_locked(WorksheetDeleteCheckpointRequest.model_validate(record["request"]), origin, consent.backupSheets)

    def _execute_locked(self, request, origin, backup_consent=None):
        folder = _folder(request.operationId)
        payload = request.model_dump(mode="json")
        if folder.exists():
            _, prior = _read(request.operationId, origin)
            if prior["request"] != payload:
                raise ValueError("检查点标识已用于另一份计划，未重复执行。")
            if backup_consent is None or prior["result"]["status"] != "needs_backup_consent":
                return _result(prior)  # Includes crashed/executing: never replay writes.
        actions = request.plan.actions
        if "deleteWorksheet" not in capabilities()["worksheetCheckpoint"]["allowedActions"]:
            raise ValueError("工作表检查点删除当前未开放。")
        if request.host == "wps-et" and not capabilities()["wpsWrite"]["enabled"]:
            raise ValueError("WPS 写入能力当前关闭，未删除。")
        if not actions or any(a.type != "deleteWorksheet" for a in actions):
            raise ValueError("结构检查点只接受纯删除工作表的白名单计划，请拆分混合任务。")
        targets = [a.sheet for a in actions]
        if len(targets) != len(set(targets)) or not set(targets).issubset(request.authorizedSheets):
            raise ValueError("删除目标重复或超出授权范围，未执行。")
        canonical_path(request.workbookPath)
        folder.mkdir(exist_ok=True)
        result = WorksheetDeleteCheckpointResult(operationId=request.operationId, status="prepared", message="正在创建完整恢复副本，尚未删除。")
        record = {"origin": origin, "createdAt": datetime.now(timezone.utc).isoformat(),
                  "request": payload, "result": result.model_dump(), "sha256": None}
        if backup_consent is not None:
            record["expectedSheetNames"] = prior["sheetNames"]
            record["backupConsent"] = list(backup_consent)
        _write(folder, record)
        started = False
        try:
            with self.native.workbook(request) as book:
                sheets = _sheets(book)
                names = [str(sheet.Name) for sheet in sheets]
                # Metadata may explain a missing permission; cell contents are not read yet.
                if (len(request.authorizedSheets) != len(set(request.authorizedSheets))
                    or not set(request.authorizedSheets).issubset(names)):
                    raise ValueError("已授权工作表发生变化，未执行删除。")
                remaining = [name for name in names if name not in targets]
                if not remaining or not any(int(sheet.Visible) == -1 for sheet in sheets if str(sheet.Name) in remaining):
                    raise ValueError("至少需要保留一张可见工作表，未执行。")
                _identity(book, request, names)
                missing = [name for name in names if name not in request.authorizedSheets]
                if backup_consent is not None and (names != record["expectedSheetNames"] or missing != backup_consent):
                    raise ValueError("等待备份许可期间工作表列表发生变化，请重新发起；未执行删除。")
                if missing and backup_consent is None:
                    result.status = "needs_backup_consent"
                    result.backupSheets = missing
                    result.message = "完整恢复备份还需包含以下未勾选工作表。是否仅允许它们进入本地备份？这些表不会因此进入分析或删除范围，也不会发送给模型。尚未备份或删除。"
                    record["sheetNames"] = names
                    record["result"] = result.model_dump()
                    _write(folder, record)
                    return result
                retained_targets = [sheet for sheet in sheets if str(sheet.Name) in targets]
                archive = folder / "before.xlsx"
                book.SaveCopyAs(str(archive))
                record["sha256"] = _validate_archive(archive, names)
                result.recoveryAvailable = True
                record["sheetNames"] = names
                _identity(book, request, names)
                result.status = "executing"
                result.message = "完整检查点已落盘，正在删除并逐表核对。"
                record["result"] = result.model_dump()
                _write(folder, record)  # Durable BEFORE first delete, even if setter throws.
                app = book.Application
                alerts, events = app.DisplayAlerts, app.EnableEvents
                try:
                    app.EnableEvents = False
                    app.DisplayAlerts = False
                    for sheet in retained_targets:
                        current = [name for name in names if name not in result.deletedSheets]
                        _identity(book, request, current)
                        name = str(sheet.Name)
                        if name not in targets or canonical_path(str(sheet.Parent.FullName)) != canonical_path(request.workbookPath):
                            raise ValueError("待删除工作表对象身份变化。")
                        started = True
                        sheet.Delete()  # Same retained, verified object; never ActiveSheet.
                        expected = [item for item in current if item != name]
                        _identity(book, request, expected)
                        result.deletedSheets.append(name)
                        result.remainingSheets = expected
                        record["result"] = result.model_dump()
                        _write(folder, record)
                finally:
                    try:
                        app.DisplayAlerts = alerts
                    finally:
                        app.EnableEvents = events
                _identity(book, request, remaining)
                result.status = "completed"
                result.remainingSheets = remaining
                result.message = f"已删除 {len(targets)} 张工作表，并核对剩余工作表。恢复会打开执行前副本，不覆盖当前原表。"
        except Exception as exc:
            result.status = "partial_or_unknown" if started else "failed"
            # Native exceptions can contain workbook values/paths. Do not echo them.
            result.message = ("删除未全部完成或结果未确认；不会自动重试。请打开执行前恢复副本核对。" if started
                else str(exc) if isinstance(exc, ValueError) else "未能建立完整检查点或确认宿主接口，未开始删除。")
        record["result"] = result.model_dump()
        _write(folder, record)
        return result

    def open(self, operation_id, origin):
        with _LOCK, folder_write_guard(application_data_root() / "worksheet-checkpoints"):
            folder, record = _read(operation_id, origin)
            if not record.get("sha256"):
                raise ValueError("本次没有可用的完整恢复副本。")
            archive = folder / "before.xlsx"
            if _validate_archive(archive, record["sheetNames"]) != record["sha256"]:
                raise ValueError("恢复副本校验失败，未打开。")
            path = folder / f"recovered-{secrets.token_hex(8)}.xlsx"
            shutil.copyfile(archive, path)
            self.native.open_copy(WorksheetDeleteCheckpointRequest.model_validate(record["request"]), path)
            return {"opened": True, "recoveryPath": str(path)}

    def restore(self, operation_id, origin):
        """**原地**撤销本次删除：把归档里的表放回当前打开的原工作簿。

        与 `open` 是两条路，不能互相冒充：
        - 这里所有写入都发生在原工作簿对象上，文件身份不变，后续编辑保留；
        - `open` 只打开一份执行前副本，不改原表，仍是"备份与恢复"里的备用入口。
        """
        with _LOCK, folder_write_guard(application_data_root() / "worksheet-checkpoints"):
            folder, record = _read(operation_id, origin)
            result = _result(record)
            stored = _stored_restore(record)
            if stored is not None and stored.status == "completed" and not record.get("restorePendingSheet"):
                # 已经撤销过：不再碰宿主，重复点击不会造出第二份表。
                return stored
            request = WorksheetDeleteCheckpointRequest.model_validate(record["request"])
            if "restoreWorksheet" not in capabilities()["worksheetCheckpoint"]["allowedActions"]:
                raise ValueError("工作表删除的原地撤销当前未开放；可从「备份与恢复」打开执行前副本。")
            if result.status != "completed" or not result.recoveryAvailable or not record.get("sha256"):
                raise ValueError("本次删除没有完整完成，不能原地撤销；可从「备份与恢复」打开执行前副本核对。")
            archive = folder / "before.xlsx"
            if _validate_archive(archive, record["sheetNames"]) != record["sha256"]:
                raise ValueError("执行前归档校验失败，未执行撤销。")
            newer = self._newer_undeleted(folder, record, request, origin)
            if newer:
                raise ValueError(
                    "更晚的一次删除还没有撤销（" + str(newer)[:19] + "）。"
                    "请先撤销那一次；本次仍可从「备份与恢复」打开执行前副本核对。")
            already = list(stored.restoredSheets) if stored is not None else []
            try:
                with self.native.restore_session(request, archive) as (book, archive_book):
                    pending = record.get("restorePendingSheet")
                    if pending:
                        uncertain = _uncertain_restore(record)
                        names = [str(sheet.Name) for sheet in _sheets(book)]
                        _identity(book, request, names)
                        if pending in names:
                            # A copy may have succeeded before its receipt failed.
                            # A same-name user sheet is indistinguishable: never adopt it.
                            return uncertain
                        # No sheet to duplicate. Commit this observation before retrying.
                        record.pop("restorePendingSheet")
                        _write(folder, record)
                    payload = restore_in_place(
                        book=book, archive=archive_book,
                        deleted=result.deletedSheets, archive_order=record["sheetNames"],
                        already_restored=already,
                        max_cells=capability_int("worksheetCheckpoint", "maxRestoreCells"),
                        before_sheet=lambda name: self._record_restore_intent(folder, record, name),
                        on_sheet=lambda name: self._record_restore_progress(folder, record, name))
            except RestoreConflict as conflict:
                # 冲突是"要用户决定"，不是"已经改了"：不落盘成终态，也不改动原表。
                return WorksheetRestoreResult(
                    operationId=operation_id, status="conflict",
                    message=str(conflict) + " 本次未改动原工作簿；仍可从「备份与恢复」打开执行前副本。")
            except Exception as exc:
                return self._restore_failure(folder, record, result, already, exc)
            payload.update({
                "operationId": operation_id,
                "message": self._restore_message(payload, already),
            })
            _remember_restore(record, payload)
            _write(folder, record)
            return WorksheetRestoreResult.model_validate(payload)

    def _restore_failure(self, folder, record, result, already, exc):
        """中止时如实记账：已放回哪些、还差哪些；下次重试不会重复造表。"""
        current = record["result"].get("restore") if isinstance(record["result"].get("restore"), dict) else {}
        restored = [name for name in current.get("restoredSheets") or [] if name in set(result.deletedSheets)]
        restored = list(dict.fromkeys(list(already) + restored))
        missing = [name for name in result.deletedSheets if name not in set(restored)]
        payload = {
            "operationId": result.operationId,
            "status": "partial" if restored else "failed",
            "restoredSheets": restored,
            "failedSheets": missing,
            "unverifiedSheets": [],
            "repairedFormulas": 0,
            "repairedNames": 0,
            "limitations": [],
            "message": (str(exc) if isinstance(exc, ValueError) else
                        "撤销过程中宿主返回未确认的结果，请核对当前工作表列表。")
                       + (" 已放回：" + "、".join(restored) + "；未放回：" + "、".join(missing) + "。"
                          if restored else " 本次未放回任何工作表，原工作簿未改动。")
                       + " 可重试撤销；执行前副本仍保留。",
        }
        if record.get("restorePendingSheet"):
            # Copy may have taken effect before raising or before verification.
            # Never claim "unchanged" merely because no verified receipt exists.
            payload = _uncertain_restore(record).model_dump(mode="json")
        _remember_restore(record, payload)
        _write(folder, record)
        return WorksheetRestoreResult.model_validate(payload)

    def _record_restore_intent(self, folder, record, name: str) -> None:
        record["restorePendingSheet"] = name
        _write(folder, record)  # Durable before Copy, even when post-copy writes fail.

    def _record_restore_progress(self, folder, record, name: str) -> None:
        """每放回一张就落一次盘：进程被杀也知道哪些已经回去了。"""
        current = record["result"].get("restore")
        restored = list(current.get("restoredSheets") or []) if isinstance(current, dict) else []
        if name not in restored:
            restored.append(name)
        record["result"]["restore"] = {
            "operationId": record["result"].get("operationId"),
            "status": "partial",
            "restoredSheets": restored,
            "failedSheets": [],
            "unverifiedSheets": [],
            "repairedFormulas": 0,
            "repairedNames": 0,
            "limitations": [],
            "message": "正在把工作表放回原工作簿，尚未全部完成。",
        }
        record.pop("restorePendingSheet", None)
        _write(folder, record)

    @staticmethod
    def _restore_message(payload: dict, already: list[str]) -> str:
        names = "、".join(payload["restoredSheets"])
        text = f"已把 {len(payload['restoredSheets'])} 张工作表放回原工作簿：{names}。"
        if already:
            text = "本次重试补齐了剩余工作表：" + names + "。"
        if payload.get("unverifiedSheets"):
            text += "其中 " + "、".join(payload["unverifiedSheets"]) + " 未完成内容核对。"
        if payload.get("limitations"):
            text += " " + "；".join(payload["limitations"]) + "。"
        if payload.get("repairedFormulas"):
            text += f" 修复了 {payload['repairedFormulas']} 处被本次删除打断的跨表公式。"
        if payload.get("repairedNames"):
            text += f" 修复了 {payload['repairedNames']} 个工作簿级名称。"
        text += " 原工作簿保持打开，未保存。"
        return text

    @staticmethod
    def _newer_undeleted(folder, record, request, origin):
        """有没有"更晚、且还没撤销"的删除。

        默认按最近一次可撤销操作逐步撤销：拿旧检查点往新结构上套，最容易把
        用户后来的操作搅乱。这里只给原因，不替用户做决定，恢复档案一个都不删。
        """
        root = _checkpoints_root()
        if root.is_symlink() or not root.is_dir():
            return None
        created = str(record.get("createdAt") or "")
        for other in sorted(root.iterdir()):
            if other.name == folder.name or not re.fullmatch(r"[a-f0-9]{32}", other.name):
                continue
            path = other / "state.json"
            if other.is_symlink() or path.is_symlink() or not path.is_file():
                continue
            try:
                row = json.loads(path.read_text(encoding="utf-8"))
                saved = row.get("request") or {}
                if row.get("origin") != origin or saved.get("host") != request.host:
                    continue
                if canonical_path(str(saved.get("workbookPath", ""))) != canonical_path(request.workbookPath):
                    continue
                if str(row.get("createdAt") or "") <= created:
                    continue
                body = row.get("result") or {}
                if body.get("status") != "completed":
                    continue
                if (body.get("restore") or {}).get("status") == "completed":
                    continue
                return row.get("createdAt")
            except (OSError, ValueError, KeyError, TypeError):
                continue
        return None

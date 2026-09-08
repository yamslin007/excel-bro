from __future__ import annotations

import base64
import hashlib
import io
import json
import logging
import math
import os
import re
import shutil
import stat
import threading
import uuid
import time
from contextvars import ContextVar
from collections.abc import Callable, Iterator
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from functools import cmp_to_key
from pathlib import Path
from typing import Literal

from openpyxl import Workbook, load_workbook
from openpyxl.chart import AreaChart, BarChart, DoughnutChart, LineChart, PieChart, Reference, ScatterChart
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import FilterColumn, Filters
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import absolute_coordinate, get_column_letter, quote_sheetname, range_boundaries
from openpyxl.workbook.defined_name import DefinedName
from pydantic import BaseModel, Field

from .models import (
    ActionExecutionResult,
    AnalysisPlan,
    DataToolRequest,
    DataToolResult,
    FolderBackupCleanupResponse,
    FolderBackupDeleteResponse,
    FolderBackupEntry,
    FolderBackupExportResponse,
    FolderBackupFileInfo,
    FolderBackupOpenResponse,
    FolderBackupPinResponse,
    FolderBackupStorageSummary,
    FolderBackupsResponse,
    FolderExecuteResponse,
    FolderLegacyBackupInfo,
    FolderLegacyCollectResponse,
    FolderUndoReference,
    FolderUndoRequest,
    FolderUndoResponse,
    RangeEqualsCriterion,
    UnverifiedAction,
    VerificationCheck,
    VerificationCriterion,
    VerificationReport,
    WorkbookSnapshot,
    WorksheetExistsCriterion,
    WorksheetSnapshot,
)
from .capabilities import capability_int, capability_text
from .folder_mutex import folder_write_guard

ROW_LIMIT = capability_int("snapshot", "dataRows")
COLUMN_LIMIT = capability_int("snapshot", "dataColumns")
FILE_LIMIT = capability_int("folder", "maxFiles")
SOURCE_SEPARATOR = " › "
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
_OPEN_WRITE_WORKBOOKS: ContextVar[list[Workbook] | None] = ContextVar(
    "open_write_workbooks", default=None
)
OUTPUT_FILE_NAME = capability_text("folder", "outputFileName")
_FOLDER_CHART_CLASSES = {
    "ColumnClustered": BarChart,
    "BarClustered": BarChart,
    "Line": LineChart,
    "Pie": PieChart,
    "Doughnut": DoughnutChart,
    "XYScatter": ScatterChart,
    "Area": AreaChart,
}
_FOLDER_CHART_TYPES = frozenset(_FOLDER_CHART_CLASSES)


class FolderWorksheetInfo(BaseModel):
    name: str
    rowCount: int
    columnCount: int


class FolderFileInfo(BaseModel):
    id: str
    name: str
    relativePath: str
    worksheets: list[FolderWorksheetInfo] = Field(default_factory=list)
    error: str | None = None


class FolderCatalog(BaseModel):
    sessionId: str
    folderName: str
    folderPath: str
    files: list[FolderFileInfo]
    totalFiles: int
    truncated: bool
    expiresAt: str


class FolderSelection(BaseModel):
    fileId: str
    sheets: list[str] = Field(min_length=1)


class FolderSnapshotRequest(BaseModel):
    sessionId: str
    selections: list[FolderSelection] = Field(min_length=1)


class FolderScanRequest(BaseModel):
    folderPath: str


class FolderRefreshRequest(BaseModel):
    sessionId: str


class FolderExecuteRequest(BaseModel):
    sessionId: str
    plan: AnalysisPlan
    operationId: str | None = Field(
        default=None,
        min_length=8,
        max_length=100,
        pattern=r"^[A-Za-z0-9][A-Za-z0-9_.:-]*$",
    )
    afterSequence: int = Field(default=0, ge=0)


FolderProgressCallback = Callable[[dict[str, object]], None]


def _emit_folder_progress(
    callback: FolderProgressCallback | None,
    event: dict[str, object],
) -> None:
    """进度回调失败不能影响文件写入；事件只携带确定性元数据。"""
    if callback is None:
        return
    try:
        callback(event)
    except Exception:
        _logger.exception("Folder execution progress callback failed")


class FolderQueryRequest(BaseModel):
    sessionId: str
    request: DataToolRequest


@dataclass(frozen=True)
class _FolderUndoTarget:
    path: Path
    backup_path: Path | None
    prior_fingerprint: str | None
    post_fingerprint: str


@dataclass(frozen=True)
class _FolderUndoRecord:
    undo_id: str
    plan_id: str
    backup_directory: Path
    targets: tuple[_FolderUndoTarget, ...]


@dataclass(frozen=True)
class _RegisteredBackupDirectory:
    root: Path
    directory: Path
    expires_at: float


@dataclass(frozen=True)
class _BackupDeleteOutcome:
    committed: bool
    cleanup_pending: bool = False
    error: str | None = None


class _FolderSession:
    def __init__(self, root: Path, files: dict[str, Path]) -> None:
        self.root = root
        self.files = files
        self.selected_targets: dict[str, tuple[Path, str]] = {}
        self.selected_sources: dict[str, tuple[str, str, Path, str]] = {}
        self.source_fingerprint: str | None = None
        self.undo_stack: list[_FolderUndoRecord] = []
        self.last_access = time.monotonic()


_sessions: dict[str, _FolderSession] = {}
_active_backup_ids: set[str] = set()
_restoring_backup_ids: set[str] = set()
_deleting_backup_ids: set[str] = set()
_exporting_backup_ids: set[str] = set()
_registered_backup_directories: dict[str, _RegisteredBackupDirectory] = {}
_backup_cleanup_lock = threading.RLock()

BACKUP_DIRECTORY_NAME = ".excel-bro-backups"
BACKUP_MANIFEST_VERSION = 2
SUPPORTED_BACKUP_MANIFEST_VERSIONS = frozenset({1, BACKUP_MANIFEST_VERSION})
BACKUP_LIFECYCLE_STATES = frozenset(
    {
        "creating",  # v1 compatibility
        "preparing",
        "prepared",
        "applying",
        "ready",
        "recovering",
        "restored",
        "rolled_back",
        "needs_attention",
    }
)
BACKUP_AUTOMATIC_CLEANUP_STATES = frozenset({"ready", "restored", "rolled_back"})
_UNDO_ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")
_MANAGEMENT_OPERATION_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{7,99}$")
_LEGACY_BACKUP_MARKER = ".excel-bro-backup-"
_logger = logging.getLogger(__name__)


def _backup_root(root: Path) -> Path:
    return root / BACKUP_DIRECTORY_NAME


def _register_backup_directory(root: Path) -> str:
    now = time.monotonic()
    root_absolute = Path(os.path.abspath(root))
    directory = _backup_root(root_absolute)
    expires_at = now + capability_int("undo", "folderBackupOpenTokenTtlSeconds")
    with _backup_cleanup_lock:
        for token, registration in list(_registered_backup_directories.items()):
            if registration.expires_at <= now:
                del _registered_backup_directories[token]
                continue
            if registration.directory == directory:
                _registered_backup_directories[token] = _RegisteredBackupDirectory(
                    root=root_absolute,
                    directory=directory,
                    expires_at=expires_at,
                )
                return token
        max_entries = max(1, capability_int("folder", "maxSessions"))
        while len(_registered_backup_directories) >= max_entries:
            oldest_token = next(iter(_registered_backup_directories))
            del _registered_backup_directories[oldest_token]
        token = uuid.uuid4().hex
        _registered_backup_directories[token] = _RegisteredBackupDirectory(
            root=root_absolute,
            directory=directory,
            expires_at=expires_at,
        )
        return token


def _registered_backup_directory(token: str) -> _RegisteredBackupDirectory:
    now = time.monotonic()
    with _backup_cleanup_lock:
        for candidate_token, registration in list(
            _registered_backup_directories.items()
        ):
            if registration.expires_at <= now:
                del _registered_backup_directories[candidate_token]
        registration = _registered_backup_directories.get(token)
    if registration is None:
        raise ValueError("备份管理凭证已失效，请刷新备份管理后重试")
    return registration


def _backup_management_root(
    session_id: str | None,
    backup_management_token: str | None,
) -> Path:
    if session_id:
        return _session(session_id).root
    if backup_management_token:
        return _registered_backup_directory(backup_management_token).root
    raise ValueError("必须提供文件夹会话或备份管理凭证")


def _explorer_window_handles() -> list[tuple[int, str]]:
    if os.name != "nt":
        return []
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    handles: list[tuple[int, str]] = []
    callback_type = ctypes.WINFUNCTYPE(
        wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
    )
    user32.EnumWindows.argtypes = [callback_type, wintypes.LPARAM]
    user32.EnumWindows.restype = wintypes.BOOL
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.GetClassNameW.argtypes = [
        wintypes.HWND,
        wintypes.LPWSTR,
        ctypes.c_int,
    ]
    user32.GetClassNameW.restype = ctypes.c_int
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [
        wintypes.HWND,
        wintypes.LPWSTR,
        ctypes.c_int,
    ]
    user32.GetWindowTextW.restype = ctypes.c_int

    def collect(handle: int, _parameter: int) -> bool:
        if not user32.IsWindowVisible(handle):
            return True
        class_name = ctypes.create_unicode_buffer(256)
        if user32.GetClassNameW(handle, class_name, len(class_name)) == 0:
            return True
        if class_name.value not in {"CabinetWClass", "ExploreWClass"}:
            return True
        title_length = user32.GetWindowTextLengthW(handle)
        title = ctypes.create_unicode_buffer(max(1, title_length + 1))
        user32.GetWindowTextW(handle, title, len(title))
        handles.append((int(handle), title.value))
        return True

    user32.EnumWindows(callback_type(collect), 0)
    return handles


def _activate_explorer_window(handle: int) -> bool:
    if os.name != "nt":
        return False
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowThreadProcessId.argtypes = [
        wintypes.HWND,
        ctypes.POINTER(wintypes.DWORD),
    ]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.AttachThreadInput.argtypes = [
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.BOOL,
    ]
    user32.AttachThreadInput.restype = wintypes.BOOL
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.ShowWindow.restype = wintypes.BOOL
    user32.BringWindowToTop.argtypes = [wintypes.HWND]
    user32.BringWindowToTop.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.SetFocus.argtypes = [wintypes.HWND]
    user32.SetFocus.restype = wintypes.HWND
    user32.PeekMessageW.argtypes = [
        ctypes.POINTER(wintypes.MSG),
        wintypes.HWND,
        wintypes.UINT,
        wintypes.UINT,
        wintypes.UINT,
    ]
    user32.PeekMessageW.restype = wintypes.BOOL
    user32.SetWindowPos.argtypes = [
        wintypes.HWND,
        wintypes.HWND,
        ctypes.c_int,
        ctypes.c_int,
        ctypes.c_int,
        ctypes.c_int,
        wintypes.UINT,
    ]
    user32.SetWindowPos.restype = wintypes.BOOL
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD
    message = wintypes.MSG()
    user32.PeekMessageW(ctypes.byref(message), None, 0, 0, 0)
    foreground = int(user32.GetForegroundWindow() or 0)
    current_thread = int(kernel32.GetCurrentThreadId())
    target_thread = int(user32.GetWindowThreadProcessId(handle, None))
    foreground_thread = (
        int(user32.GetWindowThreadProcessId(foreground, None)) if foreground else 0
    )
    attached_threads: list[int] = []
    try:
        for thread_id in {target_thread, foreground_thread}:
            if (
                thread_id
                and thread_id != current_thread
                and user32.AttachThreadInput(current_thread, thread_id, True)
            ):
                attached_threads.append(thread_id)
        user32.ShowWindow(handle, 3)  # SW_MAXIMIZE
        user32.SetWindowPos(
            handle,
            None,
            0,
            0,
            0,
            0,
            0x0001 | 0x0002 | 0x0040,  # SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW
        )
        user32.BringWindowToTop(handle)
        user32.SetForegroundWindow(handle)
        user32.SetFocus(handle)
        return int(user32.GetForegroundWindow() or 0) == handle
    finally:
        for thread_id in reversed(attached_threads):
            user32.AttachThreadInput(current_thread, thread_id, False)


def _focus_opened_explorer_window(
    directory: Path, previous_handles: set[int]
) -> bool:
    if os.name != "nt":
        return False
    deadline = time.monotonic() + 1.5
    directory_title = directory.name.casefold()
    matching_handle: int | None = None
    while time.monotonic() < deadline:
        windows = _explorer_window_handles()
        new_windows = [
            handle for handle, _title in windows if handle not in previous_handles
        ]
        if new_windows:
            matching_handle = new_windows[0]
            break
        title_matches = [
            handle
            for handle, title in windows
            if directory_title and directory_title in title.casefold()
        ]
        if title_matches:
            matching_handle = title_matches[0]
        time.sleep(0.05)
    return (
        _activate_explorer_window(matching_handle)
        if matching_handle is not None
        else False
    )


def _executions_root(root: Path) -> Path:
    return _backup_root(root) / "executions"


def _exports_root(root: Path) -> Path:
    return _backup_root(root) / "exports"


def _legacy_root(root: Path) -> Path:
    return _backup_root(root) / "legacy"


def _operations_root(root: Path) -> Path:
    return _backup_root(root) / "operations"


def _ensure_safe_backup_storage(root: Path, child_name: str) -> Path:
    backup_root = _backup_root(root)
    child = backup_root / child_name
    if backup_root.exists():
        if not backup_root.is_dir():
            raise ValueError("统一备份位置已被同名文件占用")
        if _has_reparse_point(backup_root, backup_root):
            raise ValueError("统一备份位置不能是符号链接、junction 或 reparse point")
    if child.exists():
        if not child.is_dir():
            raise ValueError(f"备份子目录 {child_name} 已被同名文件占用")
        if _has_reparse_point(child, backup_root):
            raise ValueError(
                f"备份子目录 {child_name} 不能是符号链接、junction 或 reparse point"
            )
    return child


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def _is_managed_backup_path(root: Path, path: Path) -> bool:
    root_absolute = Path(os.path.abspath(root))
    path_absolute = Path(os.path.abspath(path))
    return _is_relative_to(path_absolute, root_absolute / BACKUP_DIRECTORY_NAME)


def _iter_non_managed_files(root: Path) -> Iterator[Path]:
    """遍历工作目录，但在目录层级直接剪枝所有受管备份仓库。"""
    for current, directory_names, file_names in os.walk(root, followlinks=False):
        directory_names[:] = [
            name for name in directory_names if name != BACKUP_DIRECTORY_NAME
        ]
        current_path = Path(current)
        for name in file_names:
            yield current_path / name


def _safe_manifest_relative_path(value: object) -> Path | None:
    if not isinstance(value, str) or not value.strip():
        return None
    normalized = Path(value.replace("/", os.sep))
    if normalized.is_absolute() or ".." in normalized.parts:
        return None
    return normalized


def _has_reparse_point(path: Path, boundary: Path) -> bool:
    """不跟随链接地检查候选路径至备份根目录的 Windows reparse point。"""
    marker = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    current = path
    boundary_absolute = Path(os.path.abspath(boundary))
    while True:
        current_absolute = Path(os.path.abspath(current))
        if not _is_relative_to(current_absolute, boundary_absolute.parent):
            return True
        try:
            info = os.lstat(current)
        except FileNotFoundError:
            pass
        except OSError:
            return True
        else:
            if current.is_symlink() or (getattr(info, "st_file_attributes", 0) & marker):
                return True
        if current_absolute == boundary_absolute:
            return False
        if current.parent == current:
            return True
        current = current.parent


def validate_managed_work_directory(root: Path) -> Path:
    """Validate a user-selected work root without granting access to aliases."""

    candidate = Path(os.path.abspath(root.expanduser()))
    if not candidate.exists():
        raise ValueError(f"文件夹不存在：{root}")
    if not candidate.is_dir():
        raise ValueError(f"路径不是文件夹：{root}")
    if _has_reparse_point(candidate, candidate):
        raise ValueError("工作目录不能是符号链接、junction 或 reparse point，请重新选择")
    return candidate.resolve()


def _tree_has_reparse_point(directory: Path) -> bool:
    marker = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
    for current, directory_names, file_names in os.walk(
        directory, topdown=True, followlinks=False
    ):
        for name in [*directory_names, *file_names]:
            candidate = Path(current) / name
            try:
                info = os.lstat(candidate)
            except OSError:
                return True
            if candidate.is_symlink() or (
                getattr(info, "st_file_attributes", 0) & marker
            ):
                return True
    return False


def _atomic_write_manifest(directory: Path, payload: dict[str, object]) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    target = directory / "manifest.json"
    temporary = directory / f".manifest-{uuid.uuid4().hex}.tmp"
    serialized = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    with temporary.open("x", encoding="utf-8", newline="\n") as output:
        output.write(serialized)
        output.flush()
        os.fsync(output.fileno())
    for attempt in range(3):
        try:
            os.replace(temporary, target)
            return
        except PermissionError:
            if attempt == 2:
                raise
            time.sleep(0.05 * (attempt + 1))


def _set_manifest_status(
    directory: Path,
    manifest: dict[str, object],
    status_value: str,
    *,
    completed: bool = False,
    reason: str | None = None,
) -> None:
    """原子推进恢复点状态；失败时调用方必须保守地保留现有资料。"""

    manifest["status"] = status_value
    manifest["updatedAt"] = datetime.now().astimezone().isoformat()
    manifest["revision"] = max(int(manifest.get("revision", 0) or 0), 0) + 1
    if completed:
        manifest["completedAt"] = manifest["updatedAt"]
    if reason:
        manifest["statusReason"] = reason[:2000]
    else:
        manifest.pop("statusReason", None)
    _atomic_write_manifest(directory, manifest)


def _read_manifest(directory: Path) -> dict[str, object]:
    manifest_path = directory / "manifest.json"
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ValueError("manifest.json 不存在或无法解析") from error
    if not isinstance(payload, dict):
        raise ValueError("manifest.json 顶层必须是对象")
    if payload.get("version") not in SUPPORTED_BACKUP_MANIFEST_VERSIONS:
        raise ValueError("manifest.json 版本不受支持")
    if payload.get("undoId") != directory.name:
        raise ValueError("manifest.json 的 undoId 与目录不一致")
    if payload.get("status") not in BACKUP_LIFECYCLE_STATES:
        raise ValueError("manifest.json 状态无效")
    files = payload.get("files")
    if not isinstance(files, list):
        raise ValueError("manifest.json 缺少文件列表")
    for item in files:
        if not isinstance(item, dict):
            raise ValueError("manifest.json 文件项无效")
        if _safe_manifest_relative_path(item.get("originalRelativePath")) is None:
            raise ValueError("manifest.json 包含无效原始相对路径")
        backup_relative = item.get("backupRelativePath")
        recovery_action = item.get("recoveryAction")
        if recovery_action is not None and recovery_action not in {"restore", "remove"}:
            raise ValueError("manifest.json 包含无效恢复动作")
        if backup_relative is not None:
            relative = _safe_manifest_relative_path(backup_relative)
            if relative is None:
                raise ValueError("manifest.json 包含无效备份相对路径")
            backup_path = Path(os.path.abspath(directory / relative))
            directory_absolute = Path(os.path.abspath(directory))
            if not _is_relative_to(backup_path, directory_absolute):
                raise ValueError("manifest.json 备份路径越界")
            if _has_reparse_point(backup_path, directory):
                raise ValueError("manifest.json 备份文件路径包含 reparse point")
    return payload


def _manifest_recovery_files_present(
    directory: Path, manifest: dict[str, object]
) -> tuple[bool, str | None]:
    """Cheap health check for listing and retention; undo/export still hash files."""

    raw_files = manifest.get("files")
    if not isinstance(raw_files, list):
        return False, "恢复文件清单不存在"
    for raw_item in raw_files:
        if not isinstance(raw_item, dict):
            return False, "恢复文件清单包含无效项目"
        recovery_action = raw_item.get("recoveryAction")
        backup_relative = _safe_manifest_relative_path(
            raw_item.get("backupRelativePath")
        )
        if recovery_action not in {"restore", "remove"}:
            recovery_action = "restore" if backup_relative is not None else "remove"
        if recovery_action == "remove":
            continue
        if backup_relative is None:
            return False, "恢复文件路径缺失"
        backup_path = directory / backup_relative
        try:
            if (
                not backup_path.is_file()
                or _has_reparse_point(backup_path, directory)
            ):
                return False, f"恢复文件不存在或路径不安全：{backup_relative}"
            expected_size = raw_item.get("sizeBytes")
            if (
                isinstance(expected_size, int)
                and expected_size >= 0
                and backup_path.stat().st_size != expected_size
            ):
                return False, f"恢复文件大小已变化：{backup_relative}"
        except OSError as error:
            return False, f"无法读取恢复文件 {backup_relative}：{error}"
    return True, None


def _validate_execution_directory(
    directory: Path, backup_root: Path
) -> tuple[dict[str, object] | None, str | None]:
    executions = backup_root / "executions"
    directory_absolute = Path(os.path.abspath(directory))
    executions_absolute = Path(os.path.abspath(executions))
    if directory_absolute.parent != executions_absolute:
        return None, "候选目录不是 executions 的直接子目录"
    if not _UNDO_ID_PATTERN.fullmatch(directory.name):
        return None, "候选目录名称不是合法撤销 ID"
    if _has_reparse_point(directory, backup_root):
        return None, "候选路径包含符号链接、junction 或 reparse point"
    if _tree_has_reparse_point(directory):
        return None, "候选目录内部包含符号链接、junction 或 reparse point"
    try:
        return _read_manifest(directory), None
    except ValueError as error:
        return None, str(error)


def _active_undo_ids_for_root(backup_root: Path) -> set[str]:
    active = {
        *_active_backup_ids,
        *_restoring_backup_ids,
        *_deleting_backup_ids,
        *_exporting_backup_ids,
    }
    for session in _sessions.values():
        if _backup_root(session.root) == backup_root:
            active.update(record.undo_id for record in session.undo_stack)
    return active


def _is_safe_to_delete(
    directory: Path,
    backup_root: Path,
    *,
    protected_ids: set[str] | None = None,
) -> tuple[bool, str | None]:
    manifest, error = _validate_execution_directory(directory, backup_root)
    if manifest is None:
        return False, error
    protected = protected_ids if protected_ids is not None else _active_undo_ids_for_root(backup_root)
    if directory.name in protected:
        return False, "备份仍属于可撤销或正在执行的记录"
    return True, None


def _delete_execution_backup_transaction(
    directory: Path,
    backup_root: Path,
    *,
    protected_ids: set[str] | None = None,
) -> _BackupDeleteOutcome:
    """Atomically remove a recovery point from the live index before recursion.

    Moving the whole execution directory under .trash prevents a failed recursive
    delete from leaving a half-readable recovery point in executions.  A staged
    directory is intentionally retained when Windows still has a file open.
    """
    safe, reason = _is_safe_to_delete(
        directory, backup_root, protected_ids=protected_ids
    )
    if not safe:
        return _BackupDeleteOutcome(committed=False, error=reason)
    trash_root = backup_root / ".trash"
    operation_directory: Path | None = None
    try:
        if trash_root.exists():
            if not trash_root.is_dir() or _has_reparse_point(trash_root, backup_root):
                return _BackupDeleteOutcome(
                    committed=False, error="删除暂存目录不安全"
                )
        trash_root.mkdir(parents=True, exist_ok=True)
        operation_directory = trash_root / uuid.uuid4().hex
        operation_directory.mkdir(parents=False, exist_ok=False)
        staged_directory = operation_directory / directory.name
        os.replace(directory, staged_directory)
    except OSError as error:
        if operation_directory is not None:
            try:
                operation_directory.rmdir()
            except OSError:
                pass
        return _BackupDeleteOutcome(committed=False, error=str(error))
    try:
        shutil.rmtree(staged_directory)
        operation_directory.rmdir()
        if not any(trash_root.iterdir()):
            trash_root.rmdir()
        return _BackupDeleteOutcome(committed=True)
    except OSError as error:
        return _BackupDeleteOutcome(
            committed=True,
            cleanup_pending=True,
            error=f"备份已移出恢复列表，但内部删除暂存仍待清理：{error}",
        )


def _delete_execution_backup_best_effort(
    directory: Path,
    backup_root: Path,
    *,
    protected_ids: set[str] | None = None,
) -> str | None:
    return _delete_execution_backup_transaction(
        directory, backup_root, protected_ids=protected_ids
    ).error


def _manifest_datetime(manifest: dict[str, object]) -> datetime:
    value = manifest.get("createdAt")
    if not isinstance(value, str):
        raise ValueError("manifest.json 缺少 createdAt")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as error:
        raise ValueError("manifest.json 的 createdAt 无效") from error
    return parsed if parsed.tzinfo is not None else parsed.astimezone()


def _manifest_retention_datetime(manifest: dict[str, object]) -> datetime:
    value = manifest.get("completedAt") or manifest.get("createdAt")
    if not isinstance(value, str):
        raise ValueError("manifest.json 缺少有效保留时间")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError as error:
        raise ValueError("manifest.json 的保留时间无效") from error
    return parsed if parsed.tzinfo is not None else parsed.astimezone()


def _legacy_backup_candidates(root: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in _iter_non_managed_files(root)
            if path.is_file()
            and _LEGACY_BACKUP_MARKER in path.name
        ),
        key=lambda path: str(path.relative_to(root)).casefold(),
    )


def _managed_tree_size(directory: Path, backup_root: Path) -> tuple[int, bool]:
    """统计受管目录大小；遇到链接、reparse point 或不可读项时跳过并标记不完整。"""
    if not directory.exists():
        return 0, True
    if not directory.is_dir() or _has_reparse_point(directory, backup_root):
        return 0, False
    total = 0
    complete = True
    for current, directory_names, file_names in os.walk(directory, followlinks=False):
        current_path = Path(current)
        safe_directories: list[str] = []
        for name in directory_names:
            candidate = current_path / name
            if _has_reparse_point(candidate, backup_root):
                complete = False
            else:
                safe_directories.append(name)
        directory_names[:] = safe_directories
        for name in file_names:
            candidate = current_path / name
            try:
                if _has_reparse_point(candidate, backup_root):
                    complete = False
                    continue
                info = os.lstat(candidate)
                if not stat.S_ISREG(info.st_mode):
                    complete = False
                    continue
                total += max(int(info.st_size), 0)
            except OSError:
                complete = False
    return total, complete


def _cleanup_pending_delete_trash(backup_root: Path) -> tuple[int, list[str]]:
    """Best-effort cleanup of previously committed delete transactions only."""

    trash_root = backup_root / ".trash"
    if not trash_root.exists():
        return 0, []
    if not trash_root.is_dir() or _has_reparse_point(trash_root, backup_root):
        return 0, ["内部删除暂存目录不安全，已保留等待人工核对"]
    bytes_freed = 0
    failures: list[str] = []
    for operation_directory in list(trash_root.iterdir()):
        if (
            not operation_directory.is_dir()
            or not _UNDO_ID_PATTERN.fullmatch(operation_directory.name)
            or _has_reparse_point(operation_directory, backup_root)
            or _tree_has_reparse_point(operation_directory)
        ):
            failures.append(
                f"{operation_directory.name}：删除暂存结构无法确认，已保留"
            )
            continue
        children = list(operation_directory.iterdir())
        if len(children) != 1 or not children[0].is_dir():
            failures.append(f"{operation_directory.name}：删除暂存清单不完整，已保留")
            continue
        staged_directory = children[0]
        try:
            _read_manifest(staged_directory)
        except ValueError as error:
            failures.append(f"{operation_directory.name}：{error}")
            continue
        size_bytes, complete = _managed_tree_size(operation_directory, backup_root)
        if not complete:
            failures.append(f"{operation_directory.name}：删除暂存大小无法完整核对，已保留")
            continue
        try:
            shutil.rmtree(operation_directory)
            bytes_freed += size_bytes
        except OSError as error:
            failures.append(f"{operation_directory.name}：内部删除暂存仍待清理（{error}）")
    try:
        if trash_root.exists() and not any(trash_root.iterdir()):
            trash_root.rmdir()
    except OSError as error:
        failures.append(f"内部删除暂存目录仍待清理（{error}）")
    return bytes_freed, failures


def _prune_sessions(*, reserve_slot: bool = False) -> None:
    """淘汰会话：仅按容量限制（maxSessions），不设时间过期——远程办公场景可能长时间挂起。"""
    capacity = capability_int("folder", "maxSessions")
    target_size = capacity - 1 if reserve_slot else capacity
    while len(_sessions) > target_size:
        oldest = min(_sessions, key=lambda key: _sessions[key].last_access)
        _sessions.pop(oldest, None)


def choose_folder() -> Path | None:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
    finally:
        root.destroy()
    return Path(selected) if selected else None


def _is_supported_file(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix.lower() in SUPPORTED_SUFFIXES
        and not path.name.startswith("~$")
        and ".excel-bro-backup-" not in path.name
    )


def _candidate_files(root: Path) -> tuple[list[Path], list[Path]]:
    """返回排序后的全部候选文件与受 FILE_LIMIT 限制的候选文件。"""
    all_candidates = sorted(
        (
            path for path in _iter_non_managed_files(root) if _is_supported_file(path)
        ),
        key=lambda path: str(path.relative_to(root)).casefold(),
    )
    return all_candidates, all_candidates[:FILE_LIMIT]


def _file_info(root: Path, file_id: str, path: Path) -> FolderFileInfo:
    """读取单个文件的工作表结构；读取失败时返回带错误信息的条目。"""
    relative_path = str(path.relative_to(root))
    try:
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True,
            keep_vba=path.suffix.lower() == ".xlsm",
        )
        worksheets = [
            FolderWorksheetInfo(
                name=sheet.title,
                rowCount=sheet.max_row,
                columnCount=sheet.max_column,
            )
            for sheet in workbook.worksheets
        ]
        workbook.close()
        return FolderFileInfo(
            id=file_id,
            name=path.name,
            relativePath=relative_path,
            worksheets=worksheets,
        )
    except Exception as error:
        return FolderFileInfo(
            id=file_id,
            name=path.name,
            relativePath=relative_path,
            error=f"无法读取：{error}",
        )


def scan_folder(root: Path) -> FolderCatalog:
    root = validate_managed_work_directory(root)
    all_candidates, candidates = _candidate_files(root)
    _prune_sessions(reserve_slot=True)
    session_id = uuid.uuid4().hex
    session_files: dict[str, Path] = {}
    files: list[FolderFileInfo] = []

    for path in candidates:
        file_id = uuid.uuid5(
            uuid.NAMESPACE_URL,
            str(path.relative_to(root)).replace("\\", "/").casefold(),
        ).hex
        session_files[file_id] = path
        files.append(_file_info(root, file_id, path))

    _sessions[session_id] = _FolderSession(root, session_files)
    try:
        cleanup_folder_backups(session_id)
    except Exception as error:
        # 扫描必须优先返回数据目录；备份清理失败由备份管理器继续显示和处理。
        _logger.warning("扫描工作目录后的备份清理失败：%s", error)
    return FolderCatalog(
        sessionId=session_id,
        folderName=root.name,
        folderPath=str(root),
        files=files,
        totalFiles=len(all_candidates),
        truncated=len(all_candidates) > FILE_LIMIT,
        expiresAt=datetime.fromtimestamp(
            datetime.now().timestamp() + 86400 * 30  # 30天后（容量淘汰优先，无固定过期）
        ).astimezone().isoformat(),
    )


def scan_folder_by_path(folder_path: str) -> FolderCatalog:
    """按路径扫描文件夹，不弹对话框。

    用于前端在会话丢失（后端重启导致 _sessions 清空）后静默重建：前端本地存了
    folderPath 和工作表勾选，拿路径重扫即可，用户无需重新走三步选择。

    这里不做路径白名单——恶意网页跨站调用已被 main.py 的 Origin 守卫拦住，而本机
    恶意进程本来就能直接读这些文件，该端点不增加其能力（见威胁模型）。
    """
    candidate = validate_managed_work_directory(Path(folder_path))
    return scan_folder(candidate)


def refresh_folder(session_id: str) -> FolderCatalog:
    """刷新文件夹会话：重新扫描文件列表，保持 sessionId 不变。"""
    session = _session(session_id)
    root = session.root
    all_candidates, candidates = _candidate_files(root)
    session_files: dict[str, Path] = {}
    files: list[FolderFileInfo] = []

    for path in candidates:
        file_id = uuid.uuid5(
            uuid.NAMESPACE_URL,
            str(path.relative_to(root)).replace("\\", "/").casefold(),
        ).hex
        session_files[file_id] = path
        files.append(_file_info(root, file_id, path))

    session.files = session_files
    session.last_access = time.monotonic()
    return FolderCatalog(
        sessionId=session_id,
        folderName=root.name,
        folderPath=str(root),
        files=files,
        totalFiles=len(all_candidates),
        truncated=len(all_candidates) > FILE_LIMIT,
        expiresAt=datetime.fromtimestamp(
            datetime.now().timestamp() + 86400 * 30  # 30天后（容量淘汰优先，无固定过期）
        ).astimezone().isoformat(),
    )


def select_and_scan_folder() -> FolderCatalog | None:
    selected = choose_folder()
    return scan_folder(selected) if selected else None


def _session(session_id: str) -> _FolderSession:
    _prune_sessions()
    try:
        session = _sessions[session_id]
        session.last_access = time.monotonic()
        return session
    except KeyError as error:
        raise ValueError("文件夹会话已失效，请重新选择文件夹") from error


def folder_session_root(session_id: str) -> Path:
    """Return the exact root previously authorized by a folder session."""

    return _session(session_id).root


def _backup_entry(
    backup_root: Path,
    directory: Path,
    manifest: dict[str, object],
) -> FolderBackupEntry:
    files: list[FolderBackupFileInfo] = []
    for raw_item in manifest.get("files", []):
        if not isinstance(raw_item, dict):
            continue
        original = str(raw_item.get("originalRelativePath", ""))
        backup_relative = raw_item.get("backupRelativePath")
        recovery_action = raw_item.get("recoveryAction")
        if recovery_action not in {"restore", "remove"}:
            recovery_action = "restore" if isinstance(backup_relative, str) else "remove"
        size = raw_item.get("sizeBytes", 0)
        files.append(
            FolderBackupFileInfo(
                originalRelativePath=original,
                backupRelativePath=(
                    str(backup_relative) if isinstance(backup_relative, str) else None
                ),
                recoveryAction=recovery_action,
                sizeBytes=int(size) if isinstance(size, int) and size >= 0 else 0,
            )
        )
    created_at = _manifest_datetime(manifest)
    retention_days = capability_int("undo", "folderBackupRetentionDays")
    stored_status = str(manifest.get("status", "ready"))
    supports_direct_undo = manifest.get("version") == BACKUP_MANIFEST_VERSION
    busy_ids = {
        *_active_backup_ids,
        *_restoring_backup_ids,
        *_deleting_backup_ids,
        *_exporting_backup_ids,
    }
    display_status = stored_status
    status_reason = (
        str(manifest.get("statusReason"))
        if isinstance(manifest.get("statusReason"), str)
        else None
    )
    if stored_status in {"creating", "preparing", "prepared", "applying", "recovering"}:
        if directory.name not in busy_ids:
            display_status = "needs_attention"
            status_reason = status_reason or "上次执行未留下可确认的完成状态，恢复资料已保留等待核对。"
    recovery_files_present, recovery_files_reason = _manifest_recovery_files_present(
        directory, manifest
    )
    if not recovery_files_present:
        display_status = "needs_attention"
        status_reason = recovery_files_reason
    undo_session_id = next(
        (
            candidate_session_id
            for candidate_session_id, candidate_session in _sessions.items()
            if _backup_root(candidate_session.root) == backup_root
            and any(
                record.undo_id == directory.name
                for record in candidate_session.undo_stack
            )
        ),
        None,
    )
    pinned = manifest.get("pinned") is True
    protected_reasons: list[str] = []
    if pinned:
        protected_reasons.append("用户已保留")
    if undo_session_id is not None:
        protected_reasons.append("当前可撤销链需要")
    if display_status == "needs_attention":
        protected_reasons.append("异常资料待核对")
    delete_invalidates: set[str] = set()
    for candidate_session in _sessions.values():
        if _backup_root(candidate_session.root) != backup_root:
            continue
        delete_invalidates.update(
            _undo_ids_invalidated_by_deletion(
                candidate_session.undo_stack, directory.name
            )
        )
    delete_invalidates.discard(directory.name)
    return FolderBackupEntry(
        revision=max(int(manifest.get("revision", 0) or 0), 0),
        undoId=directory.name,
        undoSessionId=undo_session_id,
        planId=str(manifest.get("planId", directory.name)),
        planTitle=str(manifest.get("planTitle", "文件夹执行")),
        createdAt=created_at.isoformat(),
        status=display_status,
        statusReason=status_reason,
        undoAvailability=(
            "eligible"
            if undo_session_id is not None
            and display_status == "ready"
            and supports_direct_undo
            else "busy"
            if directory.name in busy_ids
            else "view_only"
        ),
        undoUnavailableReason=(
            None
            if undo_session_id is not None
            and display_status == "ready"
            and supports_direct_undo
            else status_reason
            or (
                "旧版恢复点缺少可验证的直接撤销链，可取出副本后人工核对。"
                if not supports_direct_undo
                else "当前会话没有可验证的直接撤销链，可取出副本后人工核对。"
            )
        ),
        deleteInvalidatesUndoIds=sorted(delete_invalidates),
        pinned=pinned,
        protectedReasons=protected_reasons,
        canUndo=(
            undo_session_id is not None
            and display_status == "ready"
            and supports_direct_undo
        ),
        fileCount=len(files),
        totalBytes=sum(item.sizeBytes for item in files),
        expiresAt=(created_at + timedelta(days=retention_days)).isoformat(),
        files=files,
    )


def inspect_managed_folder_recovery(root: Path) -> dict[str, object]:
    """Inspect only managed recovery metadata; never read workbook data rows."""

    root = validate_managed_work_directory(root)
    backup_root = _backup_root(root)
    if not backup_root.exists():
        return {
            "recoveryCount": 0,
            "attentionCount": 0,
            "operationCount": 0,
            "warnings": [],
        }
    if not backup_root.is_dir() or _has_reparse_point(backup_root, backup_root):
        raise ValueError("统一备份位置不是安全目录，请重新选择工作目录后核对")

    recovery_count = 0
    attention_count = 0
    operation_count = 0
    warnings: list[str] = []
    executions = backup_root / "executions"
    if executions.exists():
        if not executions.is_dir() or _has_reparse_point(executions, backup_root):
            raise ValueError("恢复资料目录包含不安全路径，已停止扫描")
        for directory in executions.iterdir():
            if not directory.is_dir():
                continue
            manifest, error = _validate_execution_directory(directory, backup_root)
            if manifest is None:
                attention_count += 1
                warnings.append(f"{directory.name}：{error}")
                continue
            recovery_count += 1
            try:
                entry = _backup_entry(backup_root, directory, manifest)
            except (TypeError, ValueError) as entry_error:
                attention_count += 1
                warnings.append(f"{directory.name}：{entry_error}")
                continue
            if entry.status == "needs_attention":
                attention_count += 1

    operations = backup_root / "operations"
    terminal_operation_states = {"completed", "rolled_back", "aborted"}
    if operations.exists():
        if not operations.is_dir() or _has_reparse_point(operations, backup_root):
            raise ValueError("恢复事务目录包含不安全路径，已停止扫描")
        for directory in operations.iterdir():
            if not directory.is_dir():
                continue
            operation_count += 1
            if _has_reparse_point(directory, backup_root):
                attention_count += 1
                warnings.append(f"{directory.name}：恢复事务目录路径不安全")
                continue
            try:
                payload = json.loads(
                    (directory / "manifest.json").read_text(encoding="utf-8")
                )
            except (OSError, json.JSONDecodeError) as operation_error:
                attention_count += 1
                warnings.append(f"{directory.name}：恢复事务记录不可读（{operation_error}）")
                continue
            if (
                not isinstance(payload, dict)
                or payload.get("operationId") != directory.name
                or not isinstance(payload.get("status"), str)
            ):
                attention_count += 1
                warnings.append(f"{directory.name}：恢复事务记录无效")
                continue
            if payload["status"] not in terminal_operation_states:
                attention_count += 1

    return {
        "recoveryCount": recovery_count,
        "attentionCount": attention_count,
        "operationCount": operation_count,
        "warnings": warnings[:50],
    }


def list_folder_backups(
    session_id: str | None = None, folder_path: str | None = None
) -> FolderBackupsResponse:
    """列出文件夹的备份。

    优先用 session_id（若提供）；否则用 folder_path。至少需提供一个。
    用 folder_path 时不建立会话，只读备份目录，撤销能力字段会全为 False。
    """
    if session_id:
        root = _session(session_id).root
    elif folder_path:
        root = Path(folder_path)
        if not root.exists() or not root.is_dir():
            raise ValueError(f"文件夹路径无效：{folder_path}")
    else:
        raise ValueError("必须提供 sessionId 或 folderPath")

    backup_root = _backup_root(root)
    open_directory_token = (
        _register_backup_directory(root)
        if backup_root.exists()
        and backup_root.is_dir()
        and not _has_reparse_point(backup_root, backup_root)
        else None
    )
    entries: list[FolderBackupEntry] = []
    warnings: list[str] = []
    executions = backup_root / "executions"
    if executions.exists():
        for directory in executions.iterdir():
            if not directory.is_dir():
                continue
            manifest, error = _validate_execution_directory(directory, backup_root)
            if manifest is None:
                warnings.append(f"{directory.name}：{error}")
                continue
            try:
                entries.append(_backup_entry(backup_root, directory, manifest))
            except (TypeError, ValueError) as error_value:
                warnings.append(f"{directory.name}：{error_value}")
    entries.sort(key=lambda entry: entry.createdAt, reverse=True)
    legacy: list[FolderLegacyBackupInfo] = []
    for path in _legacy_backup_candidates(root):
        try:
            info = path.stat()
            legacy.append(
                FolderLegacyBackupInfo(
                    relativePath=str(path.relative_to(root)),
                    sizeBytes=max(int(info.st_size), 0),
                    modifiedAt=datetime.fromtimestamp(info.st_mtime)
                    .astimezone()
                    .isoformat(),
                )
            )
        except OSError as error:
            warnings.append(f"{path.name}：无法统计旧版备份大小（{error}）")

    automatic_bytes = 0
    pinned_bytes = 0
    abnormal_bytes = 0
    for entry in entries:
        if entry.status not in BACKUP_AUTOMATIC_CLEANUP_STATES:
            abnormal_bytes += entry.totalBytes
        elif entry.pinned:
            pinned_bytes += entry.totalBytes
        else:
            automatic_bytes += entry.totalBytes
    exported_bytes, exports_complete = _managed_tree_size(
        _exports_root(root), backup_root
    )
    collected_legacy_bytes, collected_legacy_complete = _managed_tree_size(
        _legacy_root(root), backup_root
    )
    trash_bytes, trash_complete = _managed_tree_size(
        backup_root / ".trash", backup_root
    )
    operations_bytes, operations_complete = _managed_tree_size(
        backup_root / "operations", backup_root
    )
    legacy_bytes = collected_legacy_bytes + sum(item.sizeBytes for item in legacy)
    pending_cleanup_bytes = trash_bytes + operations_bytes
    known_total = (
        automatic_bytes
        + pinned_bytes
        + abnormal_bytes
        + exported_bytes
        + legacy_bytes
        + pending_cleanup_bytes
    )
    storage_summary = FolderBackupStorageSummary(
        automaticBytes=automatic_bytes,
        pinnedBytes=pinned_bytes,
        abnormalBytes=abnormal_bytes,
        exportedBytes=exported_bytes,
        legacyBytes=legacy_bytes,
        pendingCleanupBytes=pending_cleanup_bytes,
        totalBytes=known_total,
        softLimitBytes=capability_int("undo", "folderBackupMaxBytes"),
        complete=(
            not warnings
            and exports_complete
            and collected_legacy_complete
            and trash_complete
            and operations_complete
        ),
        calculatedAt=datetime.now().astimezone().isoformat(),
    )
    return FolderBackupsResponse(
        folderPath=str(root),
        backupDirectory=str(backup_root),
        openDirectoryToken=open_directory_token,
        backupManagementToken=open_directory_token,
        backups=entries,
        legacyCandidates=legacy,
        storageSummary=storage_summary,
        warnings=warnings,
    )


def open_folder_backup_directory(directory_token: str) -> FolderBackupOpenResponse:
    """在系统文件管理器中打开后端已登记的统一备份目录。"""
    registration = _registered_backup_directory(directory_token)
    root = registration.root
    directory = registration.directory
    root_absolute = Path(os.path.abspath(root))
    directory_absolute = Path(os.path.abspath(directory))
    expected = Path(os.path.abspath(_backup_root(root_absolute)))
    if (
        directory_absolute != expected
        or directory_absolute.name != BACKUP_DIRECTORY_NAME
        or directory_absolute.parent != root_absolute
    ):
        raise ValueError("备份目录未通过路径边界检查")
    if not root_absolute.exists() or not root_absolute.is_dir():
        raise ValueError("原工作目录已不存在，请重新选择文件夹")
    if not directory_absolute.exists():
        raise ValueError("当前还没有备份目录，请先执行一次需要备份的写入操作")
    if not directory_absolute.is_dir():
        raise ValueError("统一备份位置已被同名文件占用")
    if _has_reparse_point(directory_absolute, directory_absolute):
        raise ValueError("统一备份位置不能是符号链接、junction 或 reparse point")

    opener = getattr(os, "startfile", None)
    if not callable(opener):
        raise OSError("当前系统暂不支持直接打开文件夹，请复制路径后手动打开")
    try:
        previous_handles = {
            handle for handle, _title in _explorer_window_handles()
        }
    except Exception:
        previous_handles = set()
        _logger.warning(
            "无法读取现有文件资源管理器窗口，将继续打开备份目录",
            exc_info=True,
        )
    try:
        opener(str(directory_absolute))
    except OSError as error:
        raise OSError(f"无法打开统一备份目录：{error}") from error
    try:
        _focus_opened_explorer_window(directory_absolute, previous_handles)
    except Exception:
        _logger.warning(
            "备份目录已打开，但无法将文件资源管理器切换到前台",
            exc_info=True,
        )
    return FolderBackupOpenResponse(opened=True, directory=str(directory_absolute))


def cleanup_folder_backups(
    session_id: str | None = None,
    *,
    backup_management_token: str | None = None,
) -> FolderBackupCleanupResponse:
    root = _backup_management_root(session_id, backup_management_token)
    backup_root = _backup_root(root)
    executions = backup_root / "executions"
    retention = timedelta(days=capability_int("undo", "folderBackupRetentionDays"))
    max_bytes = capability_int("undo", "folderBackupMaxBytes")
    now = datetime.now().astimezone()
    deleted_executions = 0
    deleted_files = 0
    bytes_freed = 0
    skipped_active = 0
    failures: list[str] = []

    with _backup_cleanup_lock:
        pending_bytes_freed, pending_failures = _cleanup_pending_delete_trash(backup_root)
        bytes_freed += pending_bytes_freed
        failures.extend(pending_failures)
        if not executions.exists():
            return FolderBackupCleanupResponse(
                deletedExecutions=0,
                deletedFiles=0,
                bytesFreed=bytes_freed,
                skippedActive=0,
                protectedExecutions=0,
                retainedBytes=0,
                overLimitBytes=0,
                failures=failures,
                legacyCandidateCount=len(_legacy_backup_candidates(root)),
            )
        protected = _active_undo_ids_for_root(backup_root)
        candidates: list[tuple[Path, dict[str, object], datetime, int, int]] = []
        for directory in executions.iterdir():
            if not directory.is_dir():
                continue
            manifest, error = _validate_execution_directory(directory, backup_root)
            if manifest is None:
                failures.append(f"{directory.name}：{error}")
                continue
            try:
                created_at = _manifest_retention_datetime(manifest)
            except ValueError as error_value:
                failures.append(f"{directory.name}：{error_value}")
                continue
            raw_files = manifest.get("files", [])
            file_count = (
                sum(
                    1
                    for item in raw_files
                    if isinstance(item, dict)
                    and item.get("backupRelativePath") is not None
                )
                if isinstance(raw_files, list)
                else 0
            )
            total_bytes = int(manifest.get("totalBytes", 0) or 0)
            candidates.append(
                (directory, manifest, created_at, max(total_bytes, 0), file_count)
            )

        remaining_bytes = sum(candidate[3] for candidate in candidates)
        deleted_ids: set[str] = set()
        protected_ids_seen: set[str] = set()
        latest_valid_by_file: dict[str, tuple[datetime, str]] = {}
        for directory, manifest, created_at, _size, _count in candidates:
            if manifest.get("pinned") is True:
                protected.add(directory.name)
            if manifest.get("status") not in BACKUP_AUTOMATIC_CLEANUP_STATES:
                continue
            recovery_files_present, _reason = _manifest_recovery_files_present(
                directory, manifest
            )
            if (
                manifest.get("version") != BACKUP_MANIFEST_VERSION
                or not recovery_files_present
            ):
                protected.add(directory.name)
                continue
            raw_files = manifest.get("files", [])
            if not isinstance(raw_files, list):
                continue
            for raw_item in raw_files:
                if not isinstance(raw_item, dict):
                    continue
                original = raw_item.get("originalRelativePath")
                if not isinstance(original, str) or not original:
                    continue
                previous = latest_valid_by_file.get(original.casefold())
                if previous is None or created_at > previous[0]:
                    latest_valid_by_file[original.casefold()] = (created_at, directory.name)
        protected_latest_ids = {value[1] for value in latest_valid_by_file.values()}

        def delete_candidate(candidate: tuple[Path, dict[str, object], datetime, int, int]) -> bool:
            nonlocal deleted_executions, deleted_files, bytes_freed, remaining_bytes
            directory, _manifest, _created, total_bytes, file_count = candidate
            outcome = _delete_execution_backup_transaction(
                directory, backup_root, protected_ids=protected
            )
            if not outcome.committed:
                failures.append(f"{directory.name}：{outcome.error or '删除未提交'}")
                return False
            if outcome.error:
                failures.append(f"{directory.name}：{outcome.error}")
            deleted_ids.add(directory.name)
            deleted_executions += 1
            deleted_files += file_count
            if not outcome.cleanup_pending:
                bytes_freed += total_bytes
                remaining_bytes = max(0, remaining_bytes - total_bytes)
            return True

        for candidate in sorted(candidates, key=lambda item: item[2]):
            directory, manifest, created_at, _size, _count = candidate
            if directory.name in protected:
                skipped_active += 1
                protected_ids_seen.add(directory.name)
                continue
            status_value = manifest.get("status")
            old_enough = now - created_at >= retention
            if status_value not in BACKUP_AUTOMATIC_CLEANUP_STATES:
                protected_ids_seen.add(directory.name)
                continue
            if directory.name in protected_latest_ids:
                skipped_active += 1
                protected_ids_seen.add(directory.name)
                continue
            if old_enough:
                delete_candidate(candidate)

        if remaining_bytes > max_bytes:
            for candidate in sorted(candidates, key=lambda item: item[2]):
                directory = candidate[0]
                if remaining_bytes <= max_bytes:
                    break
                if (
                    directory.name in protected
                    or directory.name in protected_latest_ids
                    or directory.name in deleted_ids
                ):
                    continue
                _directory, manifest, created_at, _size, _count = candidate
                if (
                    manifest.get("status") not in BACKUP_AUTOMATIC_CLEANUP_STATES
                    or now - created_at < retention
                ):
                    continue
                delete_candidate(candidate)

    return FolderBackupCleanupResponse(
        deletedExecutions=deleted_executions,
        deletedFiles=deleted_files,
        bytesFreed=bytes_freed,
        skippedActive=skipped_active,
        protectedExecutions=len(protected_ids_seen),
        retainedBytes=remaining_bytes,
        overLimitBytes=max(remaining_bytes - max_bytes, 0),
        failures=failures,
        legacyCandidateCount=len(_legacy_backup_candidates(root)),
    )


def _remove_empty_managed_backup_containers(root: Path) -> None:
    """Best-effort removal of empty managed containers only.

    The exact derived directories are checked again before rmdir.  Exported and
    legacy copies therefore keep the backup root alive, and an unexpected entry
    is never removed by this helper.
    """

    backup_root = _backup_root(root)
    executions = backup_root / "executions"
    for directory in (executions, backup_root):
        try:
            if not directory.is_dir():
                continue
            if _has_reparse_point(directory, backup_root):
                continue
            if any(directory.iterdir()):
                continue
            directory.rmdir()
        except OSError:
            _logger.warning("无法移除空备份目录 %s", directory, exc_info=True)


def _undo_ids_invalidated_by_deletion(
    stack: list[_FolderUndoRecord], undo_id: str
) -> set[str]:
    """返回删除一条记录后，因前驱缺口而不能再直接撤销的较早记录。"""

    deleted_index = next(
        (index for index, record in enumerate(stack) if record.undo_id == undo_id),
        -1,
    )
    if deleted_index < 0:
        return set()
    invalidated = {undo_id}
    affected_paths = {target.path for target in stack[deleted_index].targets}
    for record in reversed(stack[:deleted_index]):
        record_paths = {target.path for target in record.targets}
        if record_paths & affected_paths:
            invalidated.add(record.undo_id)
            affected_paths.update(record_paths)
    return invalidated


def delete_folder_backup(
    backup_management_token: str,
    undo_id: str,
    *,
    confirmed: bool,
    expected_revision: int | None = None,
    expected_invalidated_undo_ids: list[str] | None = None,
) -> FolderBackupDeleteResponse:
    if not confirmed:
        raise ValueError("删除备份前必须明确确认")
    if not _UNDO_ID_PATTERN.fullmatch(undo_id):
        raise ValueError("备份 ID 无效")

    registration = _registered_backup_directory(backup_management_token)
    root = registration.root
    backup_root = registration.directory
    root_absolute = Path(os.path.abspath(root))
    backup_root_absolute = Path(os.path.abspath(backup_root))
    expected = Path(os.path.abspath(_backup_root(root_absolute)))
    if (
        backup_root_absolute != expected
        or backup_root_absolute.name != BACKUP_DIRECTORY_NAME
        or backup_root_absolute.parent != root_absolute
    ):
        raise ValueError("备份目录未通过路径边界检查")
    if not root_absolute.exists() or not root_absolute.is_dir():
        raise ValueError("原工作目录已不存在，请重新打开备份管理")
    if not backup_root_absolute.exists() or not backup_root_absolute.is_dir():
        raise ValueError("当前备份目录已不存在，请刷新备份管理")
    if _has_reparse_point(backup_root_absolute, backup_root_absolute):
        raise ValueError("统一备份位置不能是符号链接、junction 或 reparse point")
    directory = backup_root / "executions" / undo_id

    with _backup_cleanup_lock:
        if undo_id in _active_backup_ids:
            raise ValueError("备份仍在创建中，暂时不能删除")
        if undo_id in _restoring_backup_ids:
            raise ValueError("备份正在用于撤销，暂时不能删除")
        if undo_id in _deleting_backup_ids:
            raise ValueError("备份正在删除，请稍后刷新")
        if undo_id in _exporting_backup_ids:
            raise ValueError("备份正在取出副本，暂时不能删除")
        _deleting_backup_ids.add(undo_id)

    try:
        manifest, error = _validate_execution_directory(directory, backup_root)
        if manifest is None:
            raise ValueError(f"找不到可删除的备份：{error}")
        if manifest.get("status") not in BACKUP_AUTOMATIC_CLEANUP_STATES:
            raise ValueError("这份恢复资料尚未处于安全完成状态，不能普通删除")
        actual_revision = max(int(manifest.get("revision", 0) or 0), 0)
        if expected_revision is not None and expected_revision != actual_revision:
            raise ValueError("备份状态已变化，请刷新列表并重新确认删除影响")
        current_invalidated_undo_ids: set[str] = set()
        for candidate_session in _sessions.values():
            if _backup_root(candidate_session.root) != backup_root:
                continue
            current_invalidated_undo_ids.update(
                _undo_ids_invalidated_by_deletion(
                    candidate_session.undo_stack, undo_id
                )
            )
        current_invalidated_undo_ids.discard(undo_id)
        if expected_invalidated_undo_ids is not None and (
            set(expected_invalidated_undo_ids) != current_invalidated_undo_ids
        ):
            raise ValueError("删除影响已变化，请刷新列表并重新确认")

        raw_files = manifest.get("files", [])
        files_deleted = (
            sum(
                1
                for item in raw_files
                if isinstance(item, dict)
                and item.get("backupRelativePath") is not None
            )
            if isinstance(raw_files, list)
            else 0
        )
        bytes_freed = max(int(manifest.get("totalBytes", 0) or 0), 0)

        # 手动删除和自动清理复用同一个事务：先原子移出 executions，再递归
        # 清理 .trash。递归清理失败时，恢复资格仍然与磁盘索引保持一致。
        delete_outcome = _delete_execution_backup_transaction(
            directory, backup_root, protected_ids=set()
        )
        if not delete_outcome.committed:
            raise ValueError(delete_outcome.error or "备份删除未提交")

        with _backup_cleanup_lock:
            discarded_undo = False
            invalidated_undo_ids: set[str] = set(current_invalidated_undo_ids)
            for candidate_session in _sessions.values():
                if _backup_root(candidate_session.root) != backup_root:
                    continue
                invalidated_undo_ids.update(
                    _undo_ids_invalidated_by_deletion(
                        candidate_session.undo_stack, undo_id
                    )
                )
                before = len(candidate_session.undo_stack)
                candidate_session.undo_stack = [
                    record
                    for record in candidate_session.undo_stack
                    if record.undo_id not in invalidated_undo_ids
                ]
                discarded_undo = discarded_undo or (
                    len(candidate_session.undo_stack) != before
                )
            remaining_undo_count = len(
                {
                    record.undo_id
                    for candidate_session in _sessions.values()
                    if _backup_root(candidate_session.root) == backup_root
                    for record in candidate_session.undo_stack
                }
            )
            remaining_undo_ids = sorted(
                {
                    record.undo_id
                    for candidate_session in _sessions.values()
                    if _backup_root(candidate_session.root) == backup_root
                    for record in candidate_session.undo_stack
                }
            )

        cleanup_pending = delete_outcome.cleanup_pending
        cleanup_warnings = (
            [delete_outcome.error]
            if delete_outcome.cleanup_pending and delete_outcome.error
            else []
        )

        _remove_empty_managed_backup_containers(root)
        return FolderBackupDeleteResponse(
            deletedUndoId=undo_id,
            discardedUndo=discarded_undo,
            filesDeleted=files_deleted,
            bytesFreed=0 if cleanup_pending else bytes_freed,
            remainingUndoCount=remaining_undo_count,
            invalidatedUndoIds=sorted(invalidated_undo_ids - {undo_id}),
            remainingUndoIds=remaining_undo_ids,
            cleanupPending=cleanup_pending,
            cleanupWarnings=cleanup_warnings,
        )
    finally:
        with _backup_cleanup_lock:
            _deleting_backup_ids.discard(undo_id)


def set_folder_backup_pinned(
    backup_management_token: str,
    undo_id: str,
    *,
    pinned: bool,
) -> FolderBackupPinResponse:
    if not _UNDO_ID_PATTERN.fullmatch(undo_id):
        raise ValueError("备份 ID 无效")
    registration = _registered_backup_directory(backup_management_token)
    backup_root = registration.directory
    directory = backup_root / "executions" / undo_id
    with _backup_cleanup_lock:
        if undo_id in {
            *_active_backup_ids,
            *_restoring_backup_ids,
            *_deleting_backup_ids,
            *_exporting_backup_ids,
        }:
            raise ValueError("恢复资料正在使用中，请稍后再修改保留状态")
        manifest, error = _validate_execution_directory(directory, backup_root)
        if manifest is None:
            raise ValueError(f"找不到可管理的备份：{error}")
        manifest["pinned"] = pinned
        manifest["updatedAt"] = datetime.now().astimezone().isoformat()
        manifest["revision"] = max(int(manifest.get("revision", 0) or 0), 0) + 1
        _atomic_write_manifest(directory, manifest)
    return FolderBackupPinResponse(undoId=undo_id, pinned=pinned)


def export_folder_backup(
    session_id: str | None,
    undo_id: str,
    *,
    backup_management_token: str | None = None,
    operation_id: str | None = None,
) -> FolderBackupExportResponse:
    root = _backup_management_root(session_id, backup_management_token)
    backup_root = _backup_root(root)
    directory = backup_root / "executions" / undo_id
    resolved_operation_id = operation_id or (
        f"export-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:12]}"
    )
    if not _MANAGEMENT_OPERATION_ID_PATTERN.fullmatch(resolved_operation_id):
        raise ValueError("取出副本的操作 ID 无效")
    with _backup_cleanup_lock:
        if undo_id in _deleting_backup_ids:
            raise ValueError("备份正在删除，不能同时取出副本")
        if undo_id in _restoring_backup_ids:
            raise ValueError("备份正在用于撤销，不能同时取出副本")
        if undo_id in _exporting_backup_ids:
            raise ValueError("这份备份正在取出副本，请稍后重试")
        _exporting_backup_ids.add(undo_id)
    try:
        manifest, error = _validate_execution_directory(directory, backup_root)
        if manifest is None:
            raise ValueError(f"找不到可导出的备份：{error}")
        if manifest.get("status") not in {
            "ready",
            "restored",
            "rolled_back",
            "needs_attention",
        }:
            raise ValueError("恢复资料仍在变化，暂时不能取出副本")
        exports_root = _ensure_safe_backup_storage(root, "exports")
        exports_root.mkdir(parents=True, exist_ok=True)
        export_directory = exports_root / resolved_operation_id
        if export_directory.exists():
            if (
                not export_directory.is_dir()
                or _has_reparse_point(export_directory, exports_root)
                or _tree_has_reparse_point(export_directory)
            ):
                raise ValueError("既有取出记录路径不安全")
            try:
                export_manifest = json.loads(
                    (export_directory / "manifest.json").read_text(encoding="utf-8")
                )
            except (OSError, json.JSONDecodeError) as replay_error:
                raise ValueError("上次取出副本未留下可验证记录，已保留等待核对") from replay_error
            if (
                not isinstance(export_manifest, dict)
                or export_manifest.get("version") != 1
                or export_manifest.get("operationType") != "export"
                or export_manifest.get("operationId") != resolved_operation_id
                or export_manifest.get("sourceUndoId") != undo_id
                or export_manifest.get("status") != "ready"
                or not isinstance(export_manifest.get("files"), list)
            ):
                raise ValueError("上次取出副本的状态无法确认，已保留等待核对")
            replayed: list[str] = []
            for raw_exported in export_manifest["files"]:
                if not isinstance(raw_exported, dict):
                    raise ValueError("上次取出副本的文件清单无效")
                relative = _safe_manifest_relative_path(raw_exported.get("relativePath"))
                if relative is None:
                    raise ValueError("上次取出副本包含无效相对路径")
                exported_path = export_directory / relative
                if (
                    not exported_path.is_file()
                    or _has_reparse_point(exported_path, export_directory)
                ):
                    raise ValueError(f"上次取出的文件已不存在或不安全：{relative}")
                expected = raw_exported.get("fingerprint")
                if (
                    isinstance(expected, str)
                    and _file_content_fingerprint(exported_path) != expected
                ):
                    raise ValueError(f"上次取出的文件已发生变化：{relative}")
                replayed.append(str(exported_path.relative_to(root)))
            return FolderBackupExportResponse(
                operationId=resolved_operation_id,
                exportDirectory=str(export_directory),
                files=replayed,
                complete=not bool(export_manifest.get("failures")),
                failures=[
                    str(item)
                    for item in export_manifest.get("failures", [])
                    if isinstance(item, str)
                ],
            )

        export_directory.mkdir(parents=False, exist_ok=False)
        export_manifest: dict[str, object] = {
            "version": 1,
            "operationType": "export",
            "operationId": resolved_operation_id,
            "sourceUndoId": undo_id,
            "status": "preparing",
            "createdAt": datetime.now().astimezone().isoformat(),
            "files": [],
        }
        _atomic_write_manifest(export_directory, export_manifest)
        exported: list[str] = []
        exported_manifest_files: list[dict[str, str]] = []
        export_failures: list[str] = []
        for raw_item in manifest.get("files", []):
            if not isinstance(raw_item, dict):
                continue
            backup_relative = _safe_manifest_relative_path(
                raw_item.get("backupRelativePath")
            )
            original_relative = _safe_manifest_relative_path(
                raw_item.get("originalRelativePath")
            )
            if backup_relative is None or original_relative is None:
                continue
            source = directory / backup_relative
            if not source.is_file():
                export_failures.append(f"{original_relative}：备份文件不存在")
                continue
            expected_fingerprint = raw_item.get("priorFingerprint")
            if (
                isinstance(expected_fingerprint, str)
                and _file_content_fingerprint(source) != expected_fingerprint
            ):
                export_failures.append(f"{original_relative}：备份文件完整性校验失败")
                continue
            target = export_directory / original_relative
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)
            except OSError as copy_error:
                export_failures.append(f"{original_relative}：{copy_error}")
                continue
            exported.append(str(target.relative_to(root)))
            exported_manifest_files.append(
                {
                    "relativePath": original_relative.as_posix(),
                    "fingerprint": _file_content_fingerprint(target),
                }
            )
            if not isinstance(expected_fingerprint, str):
                export_failures.append(
                    f"{original_relative}：旧版记录没有完整性指纹，已取出供人工核对"
                )
        if not exported:
            shutil.rmtree(export_directory)
            if export_failures:
                raise ValueError(
                    "没有可安全取出的执行前文件：" + "；".join(export_failures)
                )
            raise ValueError("这次执行没有可导出的执行前文件；它只创建了新文件")
        export_manifest["files"] = exported_manifest_files
        export_manifest["failures"] = export_failures
        export_manifest["status"] = "ready"
        export_manifest["completedAt"] = datetime.now().astimezone().isoformat()
        _atomic_write_manifest(export_directory, export_manifest)
        return FolderBackupExportResponse(
            operationId=resolved_operation_id,
            exportDirectory=str(export_directory),
            files=exported,
            complete=not export_failures,
            failures=export_failures,
        )
    except Exception:
        candidate = _exports_root(root) / resolved_operation_id
        if candidate.exists():
            try:
                payload = json.loads((candidate / "manifest.json").read_text("utf-8"))
            except (OSError, json.JSONDecodeError):
                payload = None
            if isinstance(payload, dict) and payload.get("status") == "preparing":
                if not _has_reparse_point(candidate, _exports_root(root)):
                    shutil.rmtree(candidate, ignore_errors=True)
        raise
    finally:
        with _backup_cleanup_lock:
            _exporting_backup_ids.discard(undo_id)


def collect_legacy_folder_backups(
    session_id: str | None,
    *,
    confirmed: bool,
    backup_management_token: str | None = None,
) -> FolderLegacyCollectResponse:
    if not confirmed:
        raise ValueError("归拢旧版备份前需要明确确认")
    root = _backup_management_root(session_id, backup_management_token)
    active_paths = {
        target.backup_path
        for other_session in _sessions.values()
        if other_session.root == root
        for record in other_session.undo_stack
        for target in record.targets
        if target.backup_path is not None
    }
    collected: list[str] = []
    failures: list[str] = []
    destination_root = _legacy_root(root)
    _ensure_safe_backup_storage(root, "legacy")
    for source in _legacy_backup_candidates(root):
        if source in active_paths:
            failures.append(f"{source.relative_to(root)}：仍属于可撤销记录")
            continue
        relative = source.relative_to(root)
        target = destination_root / relative
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                target = target.with_name(
                    f"{target.stem}-{uuid.uuid4().hex[:8]}{target.suffix}"
                )
            source.replace(target)
            collected.append(str(target.relative_to(root)))
        except OSError as error_value:
            failures.append(f"{relative}：{error_value}")
    return FolderLegacyCollectResponse(
        collectedFiles=collected, failures=failures
    )


def _selected_files_fingerprint(paths: set[Path]) -> str:
    digest = hashlib.sha256()
    for path in sorted(paths, key=lambda item: str(item).casefold()):
        digest.update(str(path).encode("utf-8"))
        with path.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
    return digest.hexdigest()


def _file_content_fingerprint(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _require_folder_write_space(root: Path, required_bytes: int, stage: str) -> None:
    """Stop before replacement when backup and temporary files cannot fit safely."""

    reserve = capability_int("undo", "folderBackupMinFreeBytes")
    try:
        free_bytes = max(int(shutil.disk_usage(root).free), 0)
    except OSError as error:
        raise ValueError(f"无法确认工作目录剩余空间，已停止{stage}：{error}") from error
    required_with_reserve = max(required_bytes, 0) + reserve
    if free_bytes < required_with_reserve:
        raise ValueError(
            f"工作目录剩余空间不足，已在写入原文件前停止{stage}。"
            f"至少需要 {required_with_reserve} 字节可用空间，当前约 {free_bytes} 字节。"
        )


def _normalized_openpyxl_value(cell):
    value = cell.value
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        if (
            isinstance(value, (int, float))
            and re.fullmatch(r"0+", cell.number_format or "")
        ):
            return f"{int(value):0{len(cell.number_format)}d}"
        return value
    return str(value)


def _display_openpyxl_value(cell):
    value = _normalized_openpyxl_value(cell)
    if (
        isinstance(cell.value, (int, float))
        and "%" in (cell.number_format or "")
    ):
        decimals = max(0, (cell.number_format.split("%", 1)[0].partition(".")[2].count("0")))
        return f"{cell.value * 100:.{decimals}f}%"
    return value


def create_folder_snapshot(request: FolderSnapshotRequest) -> WorkbookSnapshot:
    session = _session(request.sessionId)
    snapshots: list[WorksheetSnapshot] = []
    selected_targets: dict[str, tuple[Path, str]] = {}
    selected_sources: dict[str, tuple[str, str, Path, str]] = {}
    active_name = ""

    for selection in request.selections:
        try:
            path = session.files[selection.fileId]
        except KeyError as error:
            raise ValueError("选择中包含未知文件，请重新扫描文件夹") from error
        workbook = load_workbook(
            path,
            read_only=True,
            data_only=True,
            keep_vba=path.suffix.lower() == ".xlsm",
        )
        available = set(workbook.sheetnames)
        relative_path = str(path.relative_to(session.root))
        for sheet_name in selection.sheets:
            if sheet_name not in available:
                workbook.close()
                raise ValueError(f"{relative_path} 中不存在工作表「{sheet_name}」")
            sheet = workbook[sheet_name]
            row_count = sheet.max_row
            column_count = sheet.max_column
            cells = [
                list(row)
                for row in sheet.iter_rows(
                    min_row=1,
                    max_row=min(row_count, ROW_LIMIT + 1),
                    max_col=min(column_count, COLUMN_LIMIT),
                )
            ]
            values = [
                [_normalized_openpyxl_value(cell) for cell in row]
                for row in cells
            ]
            display_values = [
                [_display_openpyxl_value(cell) for cell in row]
                for row in cells
            ]
            display_name = f"{relative_path}{SOURCE_SEPARATOR}{sheet_name}"
            sheet_id = uuid.uuid5(
                uuid.NAMESPACE_URL,
                f"{selection.fileId}:{sheet_name}",
            ).hex
            selected_targets[display_name] = (path, sheet_name)
            selected_sources[display_name] = (
                selection.fileId,
                sheet_id,
                path,
                sheet_name,
            )
            snapshots.append(
                WorksheetSnapshot(
                    name=display_name,
                    sourceFileId=selection.fileId,
                    sourceSheetId=sheet_id,
                    sourceFile=relative_path,
                    sourceSheet=sheet_name,
                    usedRange=(
                        f"A1:{get_column_letter(column_count)}{row_count}"
                        if row_count and column_count
                        else None
                    ),
                    rowCount=row_count,
                    columnCount=column_count,
                    headers=list(values[0]) if values else [],
                    dataRows=[list(row) for row in values[1:]],
                    displayRows=[list(row) for row in display_values[1:]],
                    truncated=row_count > ROW_LIMIT + 1
                    or column_count > COLUMN_LIMIT,
                )
            )
            if not active_name:
                active_name = display_name
        workbook.close()

    if not snapshots:
        raise ValueError("请至少选择一个可读取的工作表")
    session.selected_targets = selected_targets
    session.selected_sources = selected_sources
    session.source_fingerprint = _selected_files_fingerprint(
        {path for path, _ in selected_targets.values()}
    )
    return WorkbookSnapshot(
        name=f"文件夹：{session.root.name}",
        capturedAt=datetime.now().astimezone().isoformat(),
        activeWorksheet=active_name,
        sourceFingerprint=session.source_fingerprint,
        sourceFingerprintSheets=list(selected_targets),
        worksheets=snapshots,
    )


def query_folder_data(request: FolderQueryRequest) -> DataToolResult:
    from .folder_data import AuthorizedSheet, execute_folder_query

    session = _session(request.sessionId)
    if not session.selected_sources:
        raise ValueError("文件夹数据范围尚未确认，请先生成预览")
    sources = [
        AuthorizedSheet(
            file_id=file_id,
            sheet_id=sheet_id,
            path=path,
            sheet_name=sheet_name,
            display_name=display_name,
        )
        for display_name, (
            file_id,
            sheet_id,
            path,
            sheet_name,
        ) in session.selected_sources.items()
    ]
    return execute_folder_query(request.request, sources)


def _target(
    session: _FolderSession, sheet_reference: str
) -> tuple[Path, str, Literal["source", "output"]]:
    if sheet_reference in session.selected_targets:
        path, sheet_name = session.selected_targets[sheet_reference]
        return path, sheet_name, "source"
    if SOURCE_SEPARATOR in sheet_reference:
        raise ValueError(f"计划试图操作未选择的工作表：{sheet_reference}")
    return session.root / OUTPUT_FILE_NAME, sheet_reference, "output"


def _load_for_write(path: Path) -> Workbook:
    if path.exists():
        workbook = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    else:
        workbook = Workbook()
    active = _OPEN_WRITE_WORKBOOKS.get()
    if active is not None:
        active.append(workbook)
    return workbook


def _get_or_create_sheet(workbook: Workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]
    if (
        len(workbook.sheetnames) == 1
        and workbook.active.title == "Sheet"
        and workbook.active.max_row == 1
        and workbook.active.max_column == 1
        and workbook.active["A1"].value is None
    ):
        workbook.active.title = sheet_name
        return workbook.active
    return workbook.create_sheet(sheet_name)


def _values_equal(actual: object, expected: object) -> bool:
    if actual in (None, "") and expected in (None, ""):
        return True
    if (
        isinstance(actual, (int, float))
        and not isinstance(actual, bool)
        and isinstance(expected, (int, float))
        and not isinstance(expected, bool)
    ):
        actual_number = float(actual)
        expected_number = float(expected)
        return actual_number == expected_number or abs(
            actual_number - expected_number
        ) <= math.ulp(1.0) * 16 * max(
            1.0, abs(actual_number), abs(expected_number)
        )
    return actual == expected


def _normalized_range(address: str) -> str:
    local_address = address.rsplit("!", 1)[-1]
    return local_address.replace("$", "").strip().upper()


def _verification_value(value: object):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def _matrix_verification_criteria(
    path: Path,
    sheet_name: str,
    matrix: list[list[object]],
) -> list[tuple[Path, VerificationCriterion]]:
    if not matrix or not matrix[0]:
        return [(path, WorksheetExistsCriterion(type="worksheetExists", sheet=sheet_name))]
    width = len(matrix[0])
    normalized = [
        [_verification_value(value) for value in row]
        for row in matrix
    ]
    criteria: list[tuple[Path, VerificationCriterion]] = []
    for offset in range(0, len(normalized), 500):
        expected = normalized[offset : offset + 500]
        start_row = offset + 1
        end_row = start_row + len(expected) - 1
        criteria.append(
            (
                path,
                RangeEqualsCriterion(
                    type="rangeEquals",
                    sheet=sheet_name,
                    range=(
                        f"A{start_row}:{get_column_letter(width)}{end_row}"
                    ),
                    expected=expected,
                ),
            )
        )
    return criteria


def _compare_sort_values(left: object, right: object) -> int:
    left_blank = left in (None, "")
    right_blank = right in (None, "")
    if left_blank or right_blank:
        if left_blank and right_blank:
            return 0
        return 1 if left_blank else -1
    if (
        isinstance(left, (int, float))
        and not isinstance(left, bool)
        and isinstance(right, (int, float))
        and not isinstance(right, bool)
    ):
        return (float(left) > float(right)) - (float(left) < float(right))
    if isinstance(left, bool) and isinstance(right, bool):
        return int(left) - int(right)
    left_text = str(left).casefold()
    right_text = str(right).casefold()
    return (left_text > right_text) - (left_text < right_text)


def _compare_sorted_rows(left: list[object], right: list[object], keys) -> int:
    for key in keys:
        left_value = left[key.column]
        right_value = right[key.column]
        comparison = _compare_sort_values(left_value, right_value)
        if comparison == 0:
            continue
        if left_value in (None, "") or right_value in (None, ""):
            return comparison
        return comparison if key.ascending else -comparison
    return 0


def _range_is_sorted(values: list[list[object]], keys, has_headers: bool) -> bool:
    rows = values[1:] if has_headers else values
    if not keys or any(
        key.column < 0 or any(key.column >= len(row) for row in rows)
        for key in keys
    ):
        return False
    return all(
        _compare_sorted_rows(rows[index - 1], rows[index], keys) <= 0
        for index in range(1, len(rows))
    )


def _normalized_color(value: object) -> str:
    rgb = getattr(value, "rgb", value)
    text = str(rgb or "").removeprefix("#")
    return f"#{text[-6:].upper()}" if len(text) >= 6 else text.upper()


def _openpyxl_border_style(style: str, weight: str) -> str | None:
    if style == "none":
        return None
    if style == "continuous":
        return {
            "hairline": "hair",
            "thin": "thin",
            "medium": "medium",
            "thick": "thick",
        }[weight]
    if style == "dash":
        return "mediumDashed" if weight in {"medium", "thick"} else "dashed"
    if style == "dashDot":
        return "mediumDashDot" if weight in {"medium", "thick"} else "dashDot"
    if style == "dot":
        return "dotted"
    return "double"


def _range_format_matches(sheet, criterion) -> bool:
    cells = [cell for row in _iter_range(sheet, criterion.range) for cell in row]
    if criterion.fillColor is not None and not all(
        cell.fill.fill_type == "solid"
        and _normalized_color(cell.fill.fgColor) == _normalized_color(criterion.fillColor)
        for cell in cells
    ):
        return False
    if criterion.bold is not None and not all(
        cell.font.bold == criterion.bold for cell in cells
    ):
        return False
    if criterion.fontColor is not None and not all(
        _normalized_color(cell.font.color) == _normalized_color(criterion.fontColor)
        for cell in cells
    ):
        return False
    if criterion.numberFormat is not None and not all(
        cell.number_format == criterion.numberFormat for cell in cells
    ):
        return False
    expected_horizontal = {
        "centerAcrossSelection": "centerContinuous"
    }.get(criterion.horizontal, criterion.horizontal)
    if criterion.horizontal is not None and not all(
        cell.alignment.horizontal == expected_horizontal for cell in cells
    ):
        return False
    if criterion.vertical is not None and not all(
        cell.alignment.vertical == criterion.vertical for cell in cells
    ):
        return False
    if criterion.wrapText is not None and not all(
        cell.alignment.wrap_text == criterion.wrapText for cell in cells
    ):
        return False
    min_col, min_row, max_col, max_row = range_boundaries(criterion.range)
    if criterion.rowHeight is not None and not all(
        sheet.row_dimensions[row].height is not None
        and abs(sheet.row_dimensions[row].height - criterion.rowHeight) <= 0.1
        for row in range(min_row, max_row + 1)
    ):
        return False
    if criterion.columnWidth is not None and not all(
        sheet.column_dimensions[get_column_letter(column)].width is not None
        and abs(
            sheet.column_dimensions[get_column_letter(column)].width
            - criterion.columnWidth
        )
        <= 0.1
        for column in range(min_col, max_col + 1)
    ):
        return False
    return True


def _borders_match(sheet, criterion) -> bool:
    side_names = {
        "top": "top",
        "bottom": "bottom",
        "left": "left",
        "right": "right",
        "insideHorizontal": "horizontal",
        "insideVertical": "vertical",
    }
    expected_style = _openpyxl_border_style(criterion.style, criterion.weight)
    for row in _iter_range(sheet, criterion.range):
        for cell in row:
            for side in criterion.sides:
                actual = getattr(cell.border, side_names[side])
                if actual.style != expected_style:
                    return False
                if (
                    expected_style is not None
                    and _normalized_color(actual.color)
                    != _normalized_color(criterion.color)
                ):
                    return False
    return True


def _data_validation_matches(sheet, criterion) -> bool:
    expected_type = {
        "wholeNumber": "whole",
        "textLength": "textLength",
    }.get(criterion.validationType, criterion.validationType)
    expected_operator = {
        "equalTo": "equal",
        "notEqualTo": "notEqual",
        "greaterThanOrEqualTo": "greaterThanOrEqual",
        "lessThanOrEqualTo": "lessThanOrEqual",
    }.get(criterion.operator, criterion.operator)
    expected_formula1 = (
        '"' + ",".join(str(value) for value in criterion.values) + '"'
        if expected_type == "list"
        else criterion.formula1
    )
    for validation in sheet.data_validations.dataValidation:
        ranges = {
            _normalized_range(str(cell_range))
            for cell_range in validation.sqref.ranges
        }
        if _normalized_range(criterion.range) not in ranges:
            continue
        return (
            validation.type == expected_type
            and validation.operator == expected_operator
            and str(
                "" if validation.formula1 is None else validation.formula1
            )
            == str("" if expected_formula1 is None else expected_formula1)
            and str(
                "" if validation.formula2 is None else validation.formula2
            )
            == str("" if criterion.formula2 is None else criterion.formula2)
            and bool(validation.allow_blank) == criterion.allowBlank
            and (
                criterion.prompt is None
                or validation.prompt == criterion.prompt
            )
            and (
                criterion.errorMessage is None
                or validation.error == criterion.errorMessage
            )
        )
    return False


def _freeze_panes_position(sheet) -> tuple[int, int]:
    location = sheet.freeze_panes
    if location is None:
        return 0, 0
    coordinate = getattr(location, "coordinate", str(location))
    min_col, min_row, _, _ = range_boundaries(coordinate)
    return min_row - 1, min_col - 1


def _iter_range(sheet, address: str):
    min_column, min_row, max_column, max_row = range_boundaries(address)
    return sheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_column,
        max_col=max_column,
    )


def _write_matrix(sheet, address: str, values: list[list[object]]) -> None:
    min_column, min_row, max_column, max_row = range_boundaries(address)
    if (
        max_row - min_row + 1 != len(values)
        or not values
        or max_column - min_column + 1 != len(values[0])
    ):
        raise ValueError(f"{address} 与写入数据尺寸不一致")
    for row_offset, row in enumerate(values):
        for column_offset, value in enumerate(row):
            sheet.cell(
                row=min_row + row_offset,
                column=min_column + column_offset,
                value=value,
            )


def _matrix_would_change(sheet, address: str, values: list[list[object]]) -> bool:
    min_column, min_row, max_column, max_row = range_boundaries(address)
    if (
        max_row - min_row + 1 != len(values)
        or not values
        or max_column - min_column + 1 != len(values[0])
    ):
        raise ValueError(f"{address} 与写入数据尺寸不一致")
    return any(
        not _values_equal(
            sheet.cell(row=min_row + row_offset, column=min_column + column_offset).value,
            value,
        )
        for row_offset, row in enumerate(values)
        for column_offset, value in enumerate(row)
    )


def _force_text_if_formula(value: object) -> object:
    """防公式注入：openpyxl 会把 "=..." 字符串当公式写入（data_type='f'）。

    LLM 或外部数据可能把用户输入原样填进"值"单元格；对以 = 开头的字符串
    加前导撇号强制按文本存储，避免值路径被注入可执行公式。
    """
    if isinstance(value, str) and value.startswith("="):
        return "'" + value
    return value


def _folder_mode_unsupported(action_type: str) -> None:
    raise ValueError(
        f"文件夹模式暂不支持 {action_type}；请打开工作簿后在 Excel 内执行该动作"
    )


def _preflight_range_shape(address: str) -> tuple[int, int]:
    min_column, min_row, max_column, max_row = range_boundaries(address)
    if min_column < 1 or min_row < 1 or max_column > 16_384 or max_row > 1_048_576:
        raise ValueError("超出 Excel 工作表边界")
    return max_row - min_row + 1, max_column - min_column + 1


def _action_range_fields(action) -> list[tuple[str, str]]:
    fields: list[tuple[str, str]] = []
    for attribute, label in (
        ("range", "区域"),
        ("startCell", "起始位置"),
        ("cell", "单元格"),
        ("sourceRange", "源区域"),
        ("targetRange", "目标区域"),
        ("destinationCell", "目标位置"),
    ):
        value = getattr(action, attribute, None)
        if value:
            fields.append((label, value))
    return fields


def validate_plan_for_folder(
    plan: AnalysisPlan,
    folder_session: _FolderSession,
    authorized_sheets: list[str] | None = None,
) -> None:
    session = folder_session
    unsupported = {
        "insertRange",
        "deleteRange",
        "removeDuplicates",
        "createPivotTable",
        "addShape",
        # computeValues 靠 Excel 端在执行时读实时数据求值，文件夹模式没有那条通路。
        "computeValues",
    }
    authorized = {
        name.casefold()
        for name in (
            authorized_sheets
            if authorized_sheets is not None
            else session.selected_targets.keys()
        )
    }
    created_sheets: set[tuple[Path, str]] = set()
    sheet_names: dict[Path, set[str]] = {}
    object_names: dict[Path, set[str]] = {}

    def names_for(path: Path) -> set[str]:
        existing = sheet_names.get(path)
        if existing is not None:
            return existing
        if path.exists():
            workbook = load_workbook(
                path,
                read_only=True,
                data_only=False,
                keep_vba=path.suffix.lower() == ".xlsm",
            )
            try:
                existing = set(workbook.sheetnames)
            finally:
                workbook.close()
        else:
            existing = {"Sheet"}
        sheet_names[path] = existing
        return existing

    def objects_for(path: Path) -> set[str]:
        existing = object_names.get(path)
        if existing is not None:
            return existing
        if not path.exists():
            existing = set()
        else:
            workbook = load_workbook(
                path,
                read_only=False,
                data_only=False,
                keep_vba=path.suffix.lower() == ".xlsm",
            )
            try:
                existing = {
                    name.casefold() for name in workbook.defined_names.keys()
                }
                existing.update(
                    table.name.casefold()
                    for worksheet in workbook.worksheets
                    for table in worksheet.tables.values()
                )
            finally:
                workbook.close()
        object_names[path] = existing
        return existing

    issues: list[str] = []
    for index, action in enumerate(plan.actions):
        if getattr(action, "filters", []):
            issues.append(
                f"第 {index + 1} 步：文件夹模式暂不支持 {action.type} 动作的筛选条件"
                "（filters）。请在 Excel 中手动筛选后再操作，或使用当前工作簿模式。"
            )
            continue
        if action.type in unsupported:
            issues.append(
                f"第 {index + 1} 步：文件夹模式暂不支持 {action.type}"
            )
            continue
        if (
            action.type == "createChart"
            and action.chartType not in _FOLDER_CHART_TYPES
        ):
            issues.append(
                f"第 {index + 1} 步：文件夹模式暂不支持图表类型"
                f"「{action.chartType}」。请改用当前工作簿模式，或选择受支持的图表类型。"
            )
            continue

        path, sheet_name, _ = _target(session, action.sheet)
        current_names = names_for(path)
        sheet_key = (path, sheet_name.casefold())
        existed_before_plan = (
            path.exists()
            and sheet_name in current_names
            and sheet_key not in created_sheets
        )
        if existed_before_plan and action.sheet.casefold() not in authorized:
            operation = "读取" if action.type == "splitGroupAggregate" else "写入"
            issues.append(
                f"第 {index + 1} 步：动作 {action.type} 尝试{operation}未授权的工作表"
                f"「{action.sheet}」。"
            )
            continue
        creates_sheet = not path.exists() or sheet_name not in current_names
        invalid_range = False
        shapes: dict[str, tuple[int, int]] = {}
        for label, address in _action_range_fields(action):
            try:
                shapes[address] = _preflight_range_shape(address)
            except (TypeError, ValueError) as error:
                issues.append(
                    f"第 {index + 1} 步：{label}「{address}」无效（{error}）"
                )
                invalid_range = True
        if invalid_range:
            continue

        if action.type in {"writeValues", "writeFormulas"}:
            matrix = (
                action.values
                if action.type == "writeValues"
                else action.formulas
            )
            expected_shape = (len(matrix), len(matrix[0]))
            if shapes[action.range] != expected_shape:
                rows, columns = shapes[action.range]
                issues.append(
                    f"第 {index + 1} 步：目标区域 {action.range} 是"
                    f" {rows}×{columns}，但写入矩阵是"
                    f" {expected_shape[0]}×{expected_shape[1]}"
                )
        if action.type == "writeTable" and shapes[action.startCell] != (1, 1):
            issues.append(
                f"第 {index + 1} 步：表格起始位置必须是单个单元格"
            )
        if action.type == "sortRange":
            columns = shapes[action.range][1]
            for key in action.keys:
                if key.column >= columns:
                    issues.append(
                        f"第 {index + 1} 步：排序列索引 {key.column} "
                        f"超出 {action.range} 的 {columns} 列范围"
                    )
        if action.type == "filterRange":
            columns = shapes[action.range][1]
            if action.column >= columns:
                issues.append(
                    f"第 {index + 1} 步：筛选列索引 {action.column} "
                    f"超出 {action.range} 的 {columns} 列范围"
                )

        if action.type in {"createTable", "addNamedRange"}:
            object_name = (
                action.name
                if action.type == "addNamedRange" or action.name
                else f"ExcelBroTable{index + 1}"
            )
            names = objects_for(path)
            if object_name.casefold() in names:
                issues.append(
                    f"第 {index + 1} 步：表格或命名区域名称"
                    f"「{object_name}」已存在"
                )
            else:
                names.add(object_name.casefold())

        if action.type == "deleteWorksheet":
            if sheet_name not in current_names:
                issues.append(
                    f"第 {index + 1} 步：未找到工作表「{sheet_name}」"
                )
            elif len(current_names) == 1:
                issues.append(
                    f"第 {index + 1} 步：不能删除工作簿中唯一的工作表"
                )
            else:
                current_names.remove(sheet_name)
            continue

        if (
            action.type == "splitGroupAggregate"
            and sheet_name not in current_names
        ):
            issues.append(
                f"第 {index + 1} 步：未找到源工作表「{sheet_name}」"
            )
            continue

        if action.type == "copyRange":
            source_path, source_name, _ = _target(
                session, action.sourceSheet
            )
            source_names = names_for(source_path)
            source_key = (source_path, source_name.casefold())
            source_existed_before_plan = (
                source_path.exists()
                and source_name in source_names
                and source_key not in created_sheets
            )
            if (
                source_existed_before_plan
                and action.sourceSheet.casefold() not in authorized
            ):
                issues.append(
                    f"第 {index + 1} 步：动作 {action.type} 尝试读取未授权的工作表"
                    f"「{action.sourceSheet}」。"
                )
                continue
            if source_name not in source_names:
                issues.append(
                    f"第 {index + 1} 步：未找到复制源工作表"
                    f"「{action.sourceSheet}」"
                )
                continue

        if action.type != "splitGroupAggregate":
            if creates_sheet:
                created_sheets.add(sheet_key)
            current_names.add(sheet_name)

    if issues:
        raise ValueError("执行前检查未通过：" + "；".join(issues))


def _verify_folder_plan(
    session: _FolderSession,
    plan: AnalysisPlan,
    dynamic_criteria_by_action: dict[
        int, list[tuple[Path, VerificationCriterion]]
    ] | None = None,
) -> VerificationReport:
    dynamic_criteria_by_action = dynamic_criteria_by_action or {}
    criterion_sources = [
        (
            _target(session, criterion.sheet)[0],
            _target(session, criterion.sheet)[1],
            criterion,
        )
        for criterion in plan.acceptanceCriteria
    ]
    criterion_sources.extend(
        (path, criterion.sheet, criterion)
        for criteria in dynamic_criteria_by_action.values()
        for path, criterion in criteria
    )
    criteria = [criterion for _, _, criterion in criterion_sources]
    readers: dict[Path, Workbook] = {}
    checks: list[VerificationCheck] = []
    try:
        for path, sheet_name, criterion in criterion_sources:
            if not path.exists():
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=False,
                        message=f"文件不存在：{path.name}",
                    )
                )
                continue

            workbook = readers.setdefault(
                path,
                load_workbook(
                    path,
                    data_only=False,
                    keep_vba=path.suffix.lower() == ".xlsm",
                ),
            )
            if criterion.type == "worksheetMissing":
                missing = sheet_name not in workbook.sheetnames
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=missing,
                        message=(
                            f"工作表「{sheet_name}」已删除"
                            if missing
                            else f"工作表「{sheet_name}」仍然存在"
                        ),
                    )
                )
                continue
            if sheet_name not in workbook.sheetnames:
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=False,
                        message=f"未找到工作表「{sheet_name}」",
                    )
                )
                continue

            if criterion.type == "worksheetExists":
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=True,
                        message=f"工作表「{sheet_name}」存在",
                    )
                )
                continue

            sheet = workbook[sheet_name]
            if criterion.type == "filterApplied":
                filter_columns = {
                    item.colId: item for item in sheet.auto_filter.filterColumn
                }
                filter_column = filter_columns.get(criterion.column)
                actual_values = (
                    list(filter_column.filters.filter)
                    if filter_column is not None
                    and filter_column.filters is not None
                    else []
                )
                passed = (
                    _normalized_range(sheet.auto_filter.ref or "")
                    == _normalized_range(criterion.range)
                    and sorted(str(value) for value in actual_values)
                    == sorted(str(value) for value in criterion.values)
                )
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=passed,
                        message=(
                            f"「{sheet_name}」{criterion.range} 第 "
                            f"{criterion.column + 1} 列筛选条件一致"
                            if passed
                            else f"「{sheet_name}」{criterion.range} 的筛选范围或条件与预期不一致"
                        ),
                    )
                )
                continue
            if criterion.type == "filterCleared":
                passed = not sheet.auto_filter.filterColumn
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=passed,
                        message=(
                            f"工作表「{sheet_name}」的筛选条件已清除"
                            if passed
                            else f"工作表「{sheet_name}」仍有筛选条件"
                        ),
                    )
                )
                continue
            if criterion.type == "tableExists":
                expected_range = _normalized_range(criterion.range)
                matching_table = next(
                    (
                        table
                        for table in sheet.tables.values()
                        if (
                            criterion.name is None
                            or table.name.casefold() == criterion.name.casefold()
                        )
                        and _normalized_range(table.ref) == expected_range
                    ),
                    None,
                )
                actual_headers = (
                    matching_table is not None
                    and getattr(matching_table, "headerRowCount", 1) != 0
                )
                passed = (
                    matching_table is not None
                    and actual_headers == criterion.hasHeaders
                )
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=passed,
                        message=(
                            f"「{sheet_name}」{criterion.range} 的表格范围与表头状态一致"
                            if passed
                            else f"「{sheet_name}」{criterion.range} 未找到符合预期的表格"
                        ),
                    )
                )
                continue
            if criterion.type == "rangeFormatMatches":
                passed = _range_format_matches(sheet, criterion)
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=passed,
                        message=(
                            f"「{sheet_name}」{criterion.range} 的格式属性一致"
                            if passed
                            else f"「{sheet_name}」{criterion.range} 的格式属性与预期不一致"
                        ),
                    )
                )
                continue
            if criterion.type == "bordersMatch":
                passed = _borders_match(sheet, criterion)
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=passed,
                        message=(
                            f"「{sheet_name}」{criterion.range} 的边框属性一致"
                            if passed
                            else f"「{sheet_name}」{criterion.range} 的边框属性与预期不一致"
                        ),
                    )
                )
                continue
            if criterion.type == "dataValidationMatches":
                passed = _data_validation_matches(sheet, criterion)
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=passed,
                        message=(
                            f"「{sheet_name}」{criterion.range} 的数据验证规则一致"
                            if passed
                            else f"「{sheet_name}」{criterion.range} 的数据验证规则与预期不一致"
                        ),
                    )
                )
                continue
            if criterion.type == "freezePanesMatches":
                actual_rows, actual_columns = _freeze_panes_position(sheet)
                passed = (
                    actual_rows == criterion.rows
                    and actual_columns == criterion.columns
                )
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=passed,
                        message=(
                            f"工作表「{sheet_name}」的冻结窗格位置一致"
                            if passed
                            else f"工作表「{sheet_name}」的冻结窗格位置与预期不一致"
                        ),
                    )
                )
                continue
            if criterion.type in {"chartExists", "pivotTableExists"}:
                checks.append(
                    VerificationCheck(
                        criterion=criterion,
                        passed=False,
                        message=(
                            "文件夹模式当前不能独立验证"
                            + (
                                "图表"
                                if criterion.type == "chartExists"
                                else "数据透视表"
                            )
                            + "对象"
                        ),
                    )
                )
                continue

            min_column, min_row, max_column, max_row = range_boundaries(
                criterion.range
            )
            actual = [
                [
                    _verification_value(
                        sheet.cell(row=row, column=column).value
                    )
                    for column in range(min_column, max_column + 1)
                ]
                for row in range(min_row, max_row + 1)
            ]
            if criterion.type == "rangeEmpty":
                passed = all(
                    value in (None, "") for row in actual for value in row
                )
            elif criterion.type == "rangeSorted":
                passed = _range_is_sorted(
                    actual, criterion.keys, criterion.hasHeaders
                )
            else:
                passed = len(actual) == len(criterion.expected) and all(
                    len(actual_row) == len(expected_row)
                    and all(
                        _values_equal(
                            actual_value,
                            (
                                _force_text_if_formula(expected_value)
                                if criterion.type == "rangeEquals"
                                else expected_value
                            ),
                        )
                        for actual_value, expected_value in zip(
                            actual_row, expected_row
                        )
                    )
                    for actual_row, expected_row in zip(actual, criterion.expected)
                )
            checks.append(
                VerificationCheck(
                    criterion=criterion,
                    passed=passed,
                    message=(
                        (
                            f"「{sheet_name}」{criterion.range} 已清空"
                            if criterion.type == "rangeEmpty"
                            else (
                                f"「{sheet_name}」{criterion.range} 排序顺序一致"
                                if criterion.type == "rangeSorted"
                                else (
                                    f"「{sheet_name}」{criterion.range} 公式一致"
                                    if criterion.type == "formulasEqual"
                                    else f"「{sheet_name}」{criterion.range} 写入值一致"
                                )
                            )
                        )
                        if passed
                        else f"「{sheet_name}」{criterion.range} 与预期不一致"
                    ),
                    actual=actual,
                )
            )
    finally:
        for workbook in readers.values():
            workbook.close()

    def same_range(criterion, sheet: str, address: str) -> bool:
        return (
            hasattr(criterion, "range")
            and criterion.sheet.casefold() == sheet.casefold()
            and _normalized_range(criterion.range) == _normalized_range(address)
        )

    def has_criterion(predicate) -> bool:
        return any(predicate(criterion) for criterion in criteria)

    unverified_actions: list[UnverifiedAction] = []
    for index, action in enumerate(plan.actions):
        verified = False
        later_actions = plan.actions[index + 1 :]

        def later_same_range(action_type: str):
            if not hasattr(action, "range"):
                return []
            return [
                later
                for later in later_actions
                if later.type == action_type
                and later.sheet.casefold() == action.sheet.casefold()
                and hasattr(later, "range")
                and _normalized_range(later.range)
                == _normalized_range(action.range)
            ]

        if action.type == "createWorksheet":
            verified = has_criterion(
                lambda criterion: criterion.type == "worksheetExists"
                and criterion.sheet.casefold() == action.sheet.casefold()
            )
        elif action.type == "deleteWorksheet":
            verified = has_criterion(
                lambda criterion: criterion.type == "worksheetMissing"
                and criterion.sheet.casefold() == action.sheet.casefold()
            )
        elif action.type == "writeValues":
            verified = has_criterion(
                lambda criterion: criterion.type == "rangeEquals"
                and same_range(criterion, action.sheet, action.range)
            )
        elif action.type == "writeFormulas":
            verified = has_criterion(
                lambda criterion: criterion.type == "formulasEqual"
                and same_range(criterion, action.sheet, action.range)
            )
        elif action.type == "writeTable":
            start_column, start_row, _, _ = range_boundaries(
                action.startCell
            )
            target_range = (
                f"{get_column_letter(start_column)}{start_row}:"
                f"{get_column_letter(start_column + len(action.headers) - 1)}"
                f"{start_row + len(action.rows)}"
            )
            verified = has_criterion(
                lambda criterion: criterion.type == "rangeEquals"
                and same_range(criterion, action.sheet, target_range)
            )
        elif action.type == "clearRange":
            verified = action.applyTo in {"all", "contents"} and has_criterion(
                lambda criterion: criterion.type == "rangeEmpty"
                and same_range(criterion, action.sheet, action.range)
            )
        elif action.type == "copyRange":
            verified = has_criterion(
                lambda criterion: criterion.type
                in {"rangeEquals", "formulasEqual"}
                and same_range(criterion, action.sheet, action.targetRange)
            )
        elif action.type == "splitGroupAggregate":
            verified = bool(dynamic_criteria_by_action.get(index))
        elif action.type == "sortRange":
            verified = has_criterion(
                lambda criterion: criterion.type == "rangeSorted"
                and same_range(criterion, action.sheet, action.range)
            )
        elif action.type == "filterRange":
            verified = any(
                later.type in {"filterRange", "clearFilter"}
                and later.sheet.casefold() == action.sheet.casefold()
                for later in later_actions
            ) or has_criterion(
                lambda criterion: criterion.type == "filterApplied"
                and same_range(criterion, action.sheet, action.range)
                and criterion.column == action.column
            )
        elif action.type == "clearFilter":
            verified = any(
                later.type in {"filterRange", "clearFilter"}
                and later.sheet.casefold() == action.sheet.casefold()
                for later in later_actions
            ) or has_criterion(
                lambda criterion: criterion.type == "filterCleared"
                and criterion.sheet.casefold() == action.sheet.casefold()
            )
        elif action.type == "createTable":
            verified = has_criterion(
                lambda criterion: criterion.type == "tableExists"
                and same_range(criterion, action.sheet, action.range)
            )
        elif action.type == "setFill":
            verified = bool(later_same_range("setFill")) or has_criterion(
                lambda criterion: criterion.type == "rangeFormatMatches"
                and same_range(criterion, action.sheet, action.range)
                and criterion.fillColor is not None
                and criterion.fillColor.casefold() == action.color.casefold()
            )
        elif action.type == "setFont":
            verified = (
                action.bold is None
                or any(
                    later.bold is not None
                    for later in later_same_range("setFont")
                )
                or has_criterion(
                    lambda criterion: criterion.type == "rangeFormatMatches"
                    and same_range(criterion, action.sheet, action.range)
                    and criterion.bold == action.bold
                )
            ) and (
                action.color is None
                or any(
                    later.color is not None
                    for later in later_same_range("setFont")
                )
                or has_criterion(
                    lambda criterion: criterion.type == "rangeFormatMatches"
                    and same_range(criterion, action.sheet, action.range)
                    and criterion.fontColor is not None
                    and criterion.fontColor.casefold() == action.color.casefold()
                )
            )
        elif action.type == "setNumberFormat":
            verified = bool(
                later_same_range("setNumberFormat")
            ) or has_criterion(
                lambda criterion: criterion.type == "rangeFormatMatches"
                and same_range(criterion, action.sheet, action.range)
                and criterion.numberFormat == action.formatCode
            )
        elif action.type == "setAlignment":
            verified = (
                action.horizontal is None
                or any(
                    later.horizontal is not None
                    for later in later_same_range("setAlignment")
                )
                or has_criterion(
                    lambda criterion: criterion.type == "rangeFormatMatches"
                    and same_range(criterion, action.sheet, action.range)
                    and criterion.horizontal == action.horizontal
                )
            ) and (
                action.vertical is None
                or any(
                    later.vertical is not None
                    for later in later_same_range("setAlignment")
                )
                or has_criterion(
                    lambda criterion: criterion.type == "rangeFormatMatches"
                    and same_range(criterion, action.sheet, action.range)
                    and criterion.vertical == action.vertical
                )
            ) and (
                action.wrapText is None
                or any(
                    later.wrapText is not None
                    for later in later_same_range("setAlignment")
                )
                or has_criterion(
                    lambda criterion: criterion.type == "rangeFormatMatches"
                    and same_range(criterion, action.sheet, action.range)
                    and criterion.wrapText == action.wrapText
                )
            )
        elif action.type == "resizeRange":
            verified = (
                action.rowHeight is None
                or any(
                    later.rowHeight is not None
                    for later in later_same_range("resizeRange")
                )
                or has_criterion(
                    lambda criterion: criterion.type == "rangeFormatMatches"
                    and same_range(criterion, action.sheet, action.range)
                    and criterion.rowHeight == action.rowHeight
                )
            ) and (
                action.columnWidth is None
                or any(
                    later.columnWidth is not None
                    for later in later_same_range("resizeRange")
                )
                or has_criterion(
                    lambda criterion: criterion.type == "rangeFormatMatches"
                    and same_range(criterion, action.sheet, action.range)
                    and criterion.columnWidth == action.columnWidth
                )
            )
        elif action.type == "setBorders":
            verified = has_criterion(
                lambda criterion: criterion.type == "bordersMatch"
                and same_range(criterion, action.sheet, action.range)
            )
        elif action.type == "setDataValidation":
            verified = bool(
                later_same_range("setDataValidation")
            ) or has_criterion(
                lambda criterion: criterion.type == "dataValidationMatches"
                and same_range(criterion, action.sheet, action.range)
            )
        elif action.type == "freezePanes":
            verified = any(
                later.type == "freezePanes"
                and later.sheet.casefold() == action.sheet.casefold()
                for later in later_actions
            ) or has_criterion(
                lambda criterion: criterion.type == "freezePanesMatches"
                and criterion.sheet.casefold() == action.sheet.casefold()
                and criterion.rows == action.rows
                and criterion.columns == action.columns
            )
        elif action.type == "createChart":
            verified = has_criterion(
                lambda criterion: criterion.type == "chartExists"
                and criterion.sheet.casefold() == action.sheet.casefold()
                and _normalized_range(criterion.sourceRange)
                == _normalized_range(action.sourceRange)
            )
        elif action.type == "createPivotTable":
            verified = has_criterion(
                lambda criterion: criterion.type == "pivotTableExists"
                and criterion.sheet.casefold() == action.sheet.casefold()
                and criterion.name.casefold() == action.name.casefold()
            )
        if not verified:
            unverified_actions.append(
                UnverifiedAction(
                    index=index,
                    type=action.type,
                    sheet=action.sheet,
                    message=(
                        f"第 {index + 1} 步 {action.type} 已执行，"
                        "但当前验收协议不能独立验证该操作的具体效果"
                    ),
                )
            )

    if any(not check.passed for check in checks):
        status = "failed"
    elif unverified_actions or not checks:
        status = "executed_unverified"
    else:
        status = "verified"
    return VerificationReport(
        status=status,
        passed=status == "verified",
        checks=checks,
        unverifiedActions=unverified_actions,
    )


def _is_locked_for_writing(path: Path) -> bool:
    """目标文件是否被别的进程占着写锁（最典型的是正在 Excel 里打开）。

    直接试着以读写方式打开，比检测 Excel 的 `~$` 锁文件可靠——后者在 Excel 崩溃
    后会留下陈旧残留，造成误判。这里只打开不写，没有副作用。
    """
    if not path.exists():
        return False
    try:
        with path.open("r+b"):
            return False
    except OSError:
        return True


def _save_workbooks_atomically(
    session: _FolderSession,
    workbooks: dict[Path, Workbook],
    modified_paths: set[Path],
    execution_id: str,
    plan_id: str,
    plan_title: str,
    on_progress: FolderProgressCallback | None = None,
) -> tuple[
    list[str],
    list[str],
    list[_FolderUndoTarget],
    Path | None,
    list[str],
]:
    targets = [
        (path, workbooks[path])
        for path in workbooks
        if path in modified_paths
    ]
    # 在碰磁盘之前拦住被占用的目标文件。
    #
    # 文件正在 Excel 里打开时，os.replace() 会被 Windows 拒绝（WinError 5）。原先
    # 这个失败发生在写完临时文件、做完备份之后，要走一遍回滚，最后抛 RuntimeError
    # 逃逸成 500——响应不带 CORS 头，前端只看到 Failed to fetch，什么都查不出来。
    #
    # 而且即使替换成功也不安全：Excel 手里有自己的内存副本，它下次保存会把我们的
    # 修改整个盖掉。所以这里是明确拒绝，不是想办法绕过去。
    locked = [path for path, _ in targets if _is_locked_for_writing(path)]
    if locked:
        for workbook in workbooks.values():
            workbook.close()
        names = "、".join(sorted(path.name for path in locked))
        raise ValueError(
            f"以下文件正在被其他程序占用（通常是正在 Excel 中打开）：{names}。"
            "请先在 Excel 里关闭这些文件后重试；"
            "如果要修改当前打开的工作簿，请把数据来源切换为「当前工作簿」。"
        )
    if not targets:
        for workbook in workbooks.values():
            workbook.close()
        return [], [], [], None, []
    existing_source_bytes = sum(
        path.stat().st_size for path, _ in targets if path.exists()
    )
    try:
        # 执行前先为待保存结果和原文件副本各预留一份近似空间；
        # 临时结果落盘后还会用实际剩余空间再检查一次。
        _require_folder_write_space(
            session.root, existing_source_bytes * 2, "安全保存"
        )
    except Exception:
        for workbook in workbooks.values():
            workbook.close()
        raise
    backup_root = _backup_root(session.root)
    _ensure_safe_backup_storage(session.root, "executions")
    temporary_paths: dict[Path, Path] = {}
    existed_before = {path: path.exists() for path, _ in targets}
    _emit_folder_progress(
        on_progress,
        {
            "executionId": execution_id,
            "phase": "saving",
            "stepId": "folder.prepare-files",
            "status": "started",
            "title": "正在准备待写入文件",
            "current": 0,
            "total": len(targets),
            "unit": "files",
        },
    )
    try:
        for target_index, (path, workbook) in enumerate(targets, start=1):
            temporary = path.with_name(
                f".{path.name}.excel-bro-tmp-{uuid.uuid4().hex}{path.suffix}"
            )
            workbook.save(temporary)
            temporary_paths[path] = temporary
            _emit_folder_progress(
                on_progress,
                {
                    "executionId": execution_id,
                    "phase": "saving",
                    "stepId": "folder.prepare-files",
                    "status": "progress",
                    "title": "正在准备待写入文件",
                    "summary": f"已准备 {target_index} / {len(targets)} 个文件。",
                    "current": target_index,
                    "total": len(targets),
                    "unit": "files",
                    "target": {"file": str(path.relative_to(session.root))},
                },
            )
    except Exception:
        for temporary in temporary_paths.values():
            temporary.unlink(missing_ok=True)
        raise
    finally:
        for workbook in workbooks.values():
            workbook.close()

    try:
        # 临时结果已经占用磁盘；此时确认仍能完整保存所有既有文件的执行前副本。
        _require_folder_write_space(
            session.root, existing_source_bytes, "执行前备份"
        )
    except Exception:
        for temporary in temporary_paths.values():
            temporary.unlink(missing_ok=True)
        raise

    backup_directory = backup_root / "executions" / execution_id
    backups_by_path: dict[Path, Path] = {}
    manifest_files: list[dict[str, object]] = []
    for path, _ in targets:
        relative = path.relative_to(session.root)
        manifest_files.append(
            {
                "originalRelativePath": relative.as_posix(),
                "backupRelativePath": relative.as_posix() if existed_before[path] else None,
                "recoveryAction": "restore" if existed_before[path] else "remove",
                "existedBefore": existed_before[path],
                "priorFingerprint": None,
                "postFingerprint": None,
                "sizeBytes": 0,
                "state": "pending",
            }
        )
    manifest: dict[str, object] = {
        "version": BACKUP_MANIFEST_VERSION,
        "undoId": execution_id,
        "planId": plan_id,
        "planTitle": plan_title,
        "createdAt": datetime.now().astimezone().isoformat(),
        "status": "preparing",
        "revision": 1,
        "updatedAt": datetime.now().astimezone().isoformat(),
        "totalBytes": 0,
        "files": manifest_files,
    }
    try:
        backup_directory.mkdir(parents=True, exist_ok=False)
        if _has_reparse_point(backup_directory, backup_root):
            raise ValueError("本次备份目录不能是 reparse point")
        _atomic_write_manifest(backup_directory, manifest)
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "backup",
                "stepId": "folder.backup",
                "status": "started",
                "title": "正在集中保存执行前备份",
                "current": 0,
                "total": len(targets),
                "unit": "files",
            },
        )
        for target_index, (path, _) in enumerate(targets, start=1):
            if existed_before[path]:
                backup = backup_directory / path.relative_to(session.root)
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, backup)
                backups_by_path[path] = backup
            backup = backups_by_path.get(path)
            prior_fingerprint = (
                _file_content_fingerprint(backup) if backup is not None else None
            )
            post_fingerprint = _file_content_fingerprint(temporary_paths[path])
            size_bytes = backup.stat().st_size if backup is not None else 0
            manifest_files[target_index - 1]["priorFingerprint"] = prior_fingerprint
            manifest_files[target_index - 1]["postFingerprint"] = post_fingerprint
            manifest_files[target_index - 1]["sizeBytes"] = size_bytes
            manifest_files[target_index - 1]["state"] = "prepared"
            _emit_folder_progress(
                on_progress,
                {
                    "executionId": execution_id,
                    "phase": "backup",
                    "stepId": "folder.backup",
                    "status": "progress",
                    "title": "正在集中保存执行前备份",
                    "summary": f"已处理 {target_index} / {len(targets)} 个文件的恢复点。",
                    "current": target_index,
                    "total": len(targets),
                    "unit": "files",
                    "target": {"file": str(path.relative_to(session.root))},
                },
            )
        manifest["totalBytes"] = sum(
            int(item.get("sizeBytes", 0) or 0) for item in manifest_files
        )
        _set_manifest_status(backup_directory, manifest, "prepared")
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "backup",
                "stepId": "folder.backup",
                "status": "completed",
                "title": "执行前备份已保存",
                "summary": f"已为 {len(backups_by_path)} 个现有文件保存集中恢复副本。",
                "current": len(targets),
                "total": len(targets),
                "unit": "files",
            },
        )
    except Exception:
        for temporary in temporary_paths.values():
            temporary.unlink(missing_ok=True)
        if backup_directory.exists():
            _delete_execution_backup_best_effort(
                backup_directory, backup_root, protected_ids=set()
            )
        raise

    replaced: list[Path] = []
    try:
        _set_manifest_status(backup_directory, manifest, "applying")
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "saving",
                "stepId": "folder.save",
                "status": "started",
                "title": "正在原子保存工作簿",
                "current": 0,
                "total": len(targets),
                "unit": "files",
            },
        )
        for target_index, (path, _) in enumerate(targets, start=1):
            temporary_paths[path].replace(path)
            replaced.append(path)
            actual_post_fingerprint = _file_content_fingerprint(path)
            expected_post_fingerprint = manifest_files[target_index - 1].get(
                "postFingerprint"
            )
            if actual_post_fingerprint != expected_post_fingerprint:
                raise RuntimeError(f"保存后文件指纹不一致：{path.name}")
            manifest_files[target_index - 1]["state"] = "applied"
            _set_manifest_status(backup_directory, manifest, "applying")
            _emit_folder_progress(
                on_progress,
                {
                    "executionId": execution_id,
                    "phase": "saving",
                    "stepId": "folder.save",
                    "status": "progress",
                    "title": "正在原子保存工作簿",
                    "summary": f"已保存 {target_index} / {len(targets)} 个文件。",
                    "current": target_index,
                    "total": len(targets),
                    "unit": "files",
                    "target": {"file": str(path.relative_to(session.root))},
                },
            )
        undo_targets: list[_FolderUndoTarget] = []
        total_bytes = 0
        for index, (path, _) in enumerate(targets):
            backup = backups_by_path.get(path)
            prior_fingerprint = manifest_files[index].get("priorFingerprint")
            post_fingerprint = manifest_files[index].get("postFingerprint")
            size_bytes = int(manifest_files[index].get("sizeBytes", 0) or 0)
            total_bytes += size_bytes
            undo_targets.append(
                _FolderUndoTarget(
                    path=path,
                    backup_path=backup,
                    prior_fingerprint=prior_fingerprint,
                    post_fingerprint=post_fingerprint,
                )
            )
        manifest["totalBytes"] = total_bytes
        _set_manifest_status(backup_directory, manifest, "ready", completed=True)
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "saving",
                "stepId": "folder.save",
                "status": "completed",
                "title": "工作簿已安全保存",
                "summary": f"已原子保存 {len(targets)} 个文件。",
                "current": len(targets),
                "total": len(targets),
                "unit": "files",
            },
        )
    except Exception as replace_error:
        rollback_errors: list[str] = []
        for path in reversed(replaced):
            try:
                backup = backups_by_path.get(path)
                if backup is not None:
                    shutil.copy2(backup, path)
                elif not existed_before[path]:
                    path.unlink(missing_ok=True)
                manifest_files[[candidate for candidate, _ in targets].index(path)][
                    "state"
                ] = "rolled_back"
            except Exception as rollback_error:
                rollback_errors.append(f"{path.name}: {rollback_error}")
                manifest_files[[candidate for candidate, _ in targets].index(path)][
                    "state"
                ] = "unknown"
        for path, temporary in temporary_paths.items():
            if path not in replaced:
                temporary.unlink(missing_ok=True)
        try:
            _set_manifest_status(
                backup_directory,
                manifest,
                "needs_attention" if rollback_errors else "rolled_back",
                completed=not rollback_errors,
                reason=(
                    f"写入失败且部分文件回滚失败：{'；'.join(rollback_errors)}"
                    if rollback_errors
                    else f"写入失败，已核对并回滚：{replace_error}"
                ),
            )
        except OSError:
            _logger.exception("无法持久记录执行失败状态，恢复资料将继续保留")
        detail = (
            f"；恢复失败：{'；'.join(rollback_errors)}"
            if rollback_errors
            else ""
        )
        message = (
            f"替换工作簿失败，部分文件未能回滚；恢复资料已保留：{replace_error}{detail}"
            if rollback_errors
            else f"替换工作簿失败，已恢复先前文件：{replace_error}"
        )
        raise RuntimeError(message) from replace_error

    modified = [str(path.relative_to(session.root)) for path, _ in targets]
    backups = [
        str(backups_by_path[path].relative_to(session.root))
        for path, _ in targets
        if path in backups_by_path
    ]
    warnings: list[str] = []
    if int(manifest.get("totalBytes", 0) or 0) > capability_int(
        "undo", "folderBackupMaxBytes"
    ):
        warnings.append(
            "本次备份超过备份仓库容量软上限，但已完整保留，可在备份管理中清理旧备份。"
        )
    return modified, backups, undo_targets, backup_directory, warnings


def _register_folder_undo(
    session: _FolderSession,
    session_id: str,
    undo_id: str,
    plan_id: str,
    backup_directory: Path,
    targets: list[_FolderUndoTarget],
) -> FolderUndoReference | None:
    if not targets:
        return None
    record = _FolderUndoRecord(
        undo_id=undo_id,
        plan_id=plan_id,
        backup_directory=backup_directory,
        targets=tuple(targets),
    )
    max_entries = capability_int("undo", "maxEntries")
    combined = [*session.undo_stack, record]
    session.undo_stack = combined[-max_entries:]
    return FolderUndoReference(undoId=record.undo_id, sessionId=session_id)


def _restore_folder_undo_impl(request: FolderUndoRequest) -> FolderUndoResponse:
    session = _session(request.sessionId)
    record_index = next(
        (
            index
            for index, record in enumerate(session.undo_stack)
            if record.undo_id == request.undoId
        ),
        -1,
    )
    if record_index < 0:
        raise ValueError("这次文件夹执行的撤销数据已失效或已使用")

    records = session.undo_stack[record_index:]
    return _restore_folder_records(session, records)


def _restore_folder_records(
    session: _FolderSession,
    records: list[_FolderUndoRecord],
    *,
    allowed_manifest_statuses: frozenset[str] = frozenset({"ready"}),
    recovery_operation_id: str | None = None,
) -> FolderUndoResponse:
    for record in records:
        manifest = _read_manifest(record.backup_directory)
        if manifest.get("version") != BACKUP_MANIFEST_VERSION:
            raise ValueError(
                "这份旧版恢复点缺少可验证的直接撤销链，请从备份管理取出副本后人工核对"
            )
        if manifest.get("status") not in allowed_manifest_statuses:
            raise ValueError("恢复资料当前不处于可撤销状态，请刷新备份管理后核对")
    effective_fingerprints: dict[Path, str | None] = {}
    final_sources: dict[Path, Path | None] = {}
    for record in reversed(records):
        for target in record.targets:
            if target.path in effective_fingerprints:
                current_fingerprint = effective_fingerprints[target.path]
            else:
                current_fingerprint = (
                    _file_content_fingerprint(target.path)
                    if target.path.exists()
                    else None
                )
            if current_fingerprint == target.prior_fingerprint:
                effective_fingerprints[target.path] = target.prior_fingerprint
                continue
            if current_fingerprint != target.post_fingerprint:
                raise ValueError(
                    f"无法撤销「{target.path.name}」：执行后文件又发生了变化。"
                    "为避免覆盖后续编辑，请从“备份管理”导出执行前副本后人工核对。"
                )
            if target.backup_path is not None:
                if not target.backup_path.exists():
                    raise ValueError(
                        f"无法撤销「{target.path.name}」：自动备份文件已不存在"
                    )
                backup_fingerprint = _file_content_fingerprint(
                    target.backup_path
                )
                if backup_fingerprint != target.prior_fingerprint:
                    raise ValueError(
                        f"无法撤销「{target.path.name}」：自动备份文件已发生变化"
                    )
            effective_fingerprints[target.path] = target.prior_fingerprint
            final_sources[target.path] = target.backup_path

    paths = sorted(final_sources, key=lambda path: str(path).casefold())
    if not paths:
        consumed_ids = {record.undo_id for record in records}
        for record in records:
            try:
                manifest = _read_manifest(record.backup_directory)
                if manifest.get("status") != "restored":
                    _set_manifest_status(
                        record.backup_directory, manifest, "restored", completed=True
                    )
            except (OSError, ValueError):
                _logger.exception("无法持久记录已恢复状态：%s", record.undo_id)
        backup_root = _backup_root(session.root)
        for candidate_session in _sessions.values():
            if _backup_root(candidate_session.root) == backup_root:
                candidate_session.undo_stack = [
                    record
                    for record in candidate_session.undo_stack
                    if record.undo_id not in consumed_ids
                ]
        return FolderUndoResponse(
            filesRestored=[],
            executionsUndone=len(records),
            remainingUndoCount=len(session.undo_stack),
        )
    locked = [path for path in paths if _is_locked_for_writing(path)]
    if locked:
        names = "、".join(path.name for path in locked)
        raise ValueError(
            f"以下文件正在被其他程序占用，无法撤销：{names}。"
            "请先关闭文件后重试。"
        )

    restore_temporaries: dict[Path, Path] = {}
    rollback_temporaries: dict[Path, Path] = {}
    existed_before_restore = {path: path.exists() for path in paths}
    operations_root = _ensure_safe_backup_storage(session.root, "operations")
    operations_root.mkdir(parents=True, exist_ok=True)
    operation_id = recovery_operation_id or uuid.uuid4().hex
    operation_directory = operations_root / operation_id
    if operation_directory.exists():
        if (
            not operation_directory.is_dir()
            or _has_reparse_point(operation_directory, _backup_root(session.root))
        ):
            raise ValueError("恢复事务目录不安全，已停止恢复")
        try:
            previous_operation = json.loads(
                (operation_directory / "manifest.json").read_text(encoding="utf-8")
            )
        except (OSError, json.JSONDecodeError) as error:
            raise ValueError("已有恢复事务记录不可读，已停止恢复") from error
        if (
            not isinstance(previous_operation, dict)
            or previous_operation.get("operationId") != operation_id
            or previous_operation.get("operationType") != "undo"
            or previous_operation.get("sourceUndoIds")
            != [record.undo_id for record in records]
        ):
            raise ValueError("已有恢复事务与本次请求不一致，已停止恢复")
    else:
        operation_directory.mkdir(parents=False, exist_ok=False)
    operation_files: list[dict[str, object]] = []
    operation_file_by_path: dict[Path, dict[str, object]] = {}
    for path in paths:
        relative = path.relative_to(session.root)
        item: dict[str, object] = {
            "originalRelativePath": relative.as_posix(),
            "rollbackRelativePath": (
                (Path("rollback") / relative).as_posix()
                if existed_before_restore[path]
                else None
            ),
            "state": "pending",
        }
        operation_files.append(item)
        operation_file_by_path[path] = item
    operation_manifest: dict[str, object] = {
        "version": 1,
        "operationId": operation_id,
        "operationType": "undo",
        "status": "preparing",
        "createdAt": datetime.now().astimezone().isoformat(),
        "sourceUndoIds": [record.undo_id for record in records],
        "files": operation_files,
    }
    _atomic_write_manifest(operation_directory, operation_manifest)
    for record in records:
        try:
            manifest = _read_manifest(record.backup_directory)
            _set_manifest_status(record.backup_directory, manifest, "recovering")
        except (OSError, ValueError):
            _logger.exception("无法持久记录撤销开始状态：%s", record.undo_id)
    try:
        for path in paths:
            source = final_sources[path]
            if source is not None:
                temporary = path.with_name(
                    f".{path.name}.excel-bro-undo-"
                    f"{uuid.uuid4().hex}{path.suffix}"
                )
                shutil.copy2(source, temporary)
                restore_temporaries[path] = temporary
            if path.exists():
                rollback = operation_directory / "rollback" / path.relative_to(
                    session.root
                )
                rollback.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(path, rollback)
                rollback_temporaries[path] = rollback
        operation_manifest["status"] = "prepared"
        operation_manifest["updatedAt"] = datetime.now().astimezone().isoformat()
        _atomic_write_manifest(operation_directory, operation_manifest)
    except Exception:
        for temporary in restore_temporaries.values():
            temporary.unlink(missing_ok=True)
        operation_manifest["status"] = "aborted"
        operation_manifest["updatedAt"] = datetime.now().astimezone().isoformat()
        operation_manifest["completedAt"] = operation_manifest["updatedAt"]
        operation_manifest["statusReason"] = "撤销准备失败，当前工作文件尚未改变。"
        try:
            _atomic_write_manifest(operation_directory, operation_manifest)
        except OSError:
            _logger.exception("无法持久记录撤销准备失败状态：%s", operation_id)
        if operation_directory.exists() and not _has_reparse_point(
            operation_directory, _backup_root(session.root)
        ):
            shutil.rmtree(operation_directory, ignore_errors=True)
            if operation_directory.exists():
                _logger.warning("撤销准备目录仍待清理：%s", operation_directory)
        for record in records:
            try:
                manifest = _read_manifest(record.backup_directory)
                _set_manifest_status(
                    record.backup_directory,
                    manifest,
                    "ready",
                    reason="撤销准备失败，当前工作文件尚未改变。",
                )
            except (OSError, ValueError):
                _logger.exception("无法恢复撤销准备失败状态：%s", record.undo_id)
        raise

    restored: list[Path] = []
    try:
        operation_manifest["status"] = "applying"
        operation_manifest["updatedAt"] = datetime.now().astimezone().isoformat()
        _atomic_write_manifest(operation_directory, operation_manifest)
        for path in paths:
            temporary = restore_temporaries.get(path)
            if temporary is None:
                path.unlink(missing_ok=True)
            else:
                temporary.replace(path)
            restored.append(path)
            operation_file_by_path[path]["state"] = "restored"
            operation_manifest["updatedAt"] = datetime.now().astimezone().isoformat()
            _atomic_write_manifest(operation_directory, operation_manifest)
    except Exception as restore_error:
        rollback_errors: list[str] = []
        for path in reversed(restored):
            try:
                rollback = rollback_temporaries.get(path)
                if rollback is not None:
                    shutil.copy2(rollback, path)
                elif not existed_before_restore[path]:
                    path.unlink(missing_ok=True)
                operation_file_by_path[path]["state"] = "rolled_back"
            except Exception as rollback_error:
                rollback_errors.append(f"{path.name}: {rollback_error}")
                operation_file_by_path[path]["state"] = "unknown"
        detail = (
            f"；回滚失败：{'；'.join(rollback_errors)}"
            if rollback_errors
            else ""
        )
        operation_manifest["status"] = (
            "needs_attention" if rollback_errors else "rolled_back"
        )
        operation_manifest["updatedAt"] = datetime.now().astimezone().isoformat()
        operation_manifest["completedAt"] = operation_manifest["updatedAt"]
        operation_manifest["statusReason"] = (
            f"撤销失败且部分回滚失败：{'；'.join(rollback_errors)}"
            if rollback_errors
            else f"撤销失败，已恢复撤销前状态：{restore_error}"
        )
        try:
            _atomic_write_manifest(operation_directory, operation_manifest)
        except OSError:
            _logger.exception("无法持久记录撤销事务失败状态：%s", operation_id)
        for record in records:
            try:
                manifest = _read_manifest(record.backup_directory)
                _set_manifest_status(
                    record.backup_directory,
                    manifest,
                    "needs_attention" if rollback_errors else "ready",
                    reason=(
                        f"撤销失败且部分回滚失败：{'；'.join(rollback_errors)}"
                        if rollback_errors
                        else f"撤销失败，已恢复撤销前状态：{restore_error}"
                    ),
                )
            except (OSError, ValueError):
                _logger.exception("无法持久记录撤销失败状态：%s", record.undo_id)
        message = (
            f"撤销文件夹执行失败，部分文件未能回滚；撤销前副本已保存在统一备份目录：{restore_error}{detail}"
            if rollback_errors
            else f"撤销文件夹执行失败，已恢复撤销前文件：{restore_error}"
        )
        if not rollback_errors and operation_directory.exists():
            shutil.rmtree(operation_directory, ignore_errors=True)
        raise RuntimeError(message) from restore_error
    finally:
        for temporary in restore_temporaries.values():
            temporary.unlink(missing_ok=True)

    operation_manifest["status"] = "completed"
    operation_manifest["updatedAt"] = datetime.now().astimezone().isoformat()
    operation_manifest["completedAt"] = operation_manifest["updatedAt"]
    operation_manifest.pop("statusReason", None)
    try:
        _atomic_write_manifest(operation_directory, operation_manifest)
    except OSError:
        _logger.exception("无法持久记录撤销事务完成状态：%s", operation_id)
    if operation_directory.exists():
        shutil.rmtree(operation_directory, ignore_errors=True)
        if operation_directory.exists():
            _logger.warning("已完成的撤销事务目录仍待清理：%s", operation_directory)

    for record in records:
        try:
            manifest = _read_manifest(record.backup_directory)
            _set_manifest_status(
                record.backup_directory, manifest, "restored", completed=True
            )
        except (OSError, ValueError):
            # 文件已经完整恢复；manifest 更新失败只影响后续自动清理，不能反转撤销结果。
            pass
    backup_root = _backup_root(session.root)
    consumed_ids = {record.undo_id for record in records}
    for candidate_session in _sessions.values():
        if _backup_root(candidate_session.root) == backup_root:
            candidate_session.undo_stack = [
                record
                for record in candidate_session.undo_stack
                if record.undo_id not in consumed_ids
            ]
    selected_paths = {path for path, _ in session.selected_targets.values()}
    if selected_paths and all(path.exists() for path in selected_paths):
        session.source_fingerprint = _selected_files_fingerprint(selected_paths)
    return FolderUndoResponse(
        filesRestored=[
            str(path.relative_to(session.root)) for path in paths
        ],
        executionsUndone=len(records),
        remainingUndoCount=len(session.undo_stack),
    )


def restore_managed_folder_backup(
    root: Path,
    *,
    undo_id: str,
    expected_revision: int,
    recovery_operation_id: str,
) -> FolderUndoResponse:
    """Restore one persisted recovery point without relying on an in-memory session."""

    root = validate_managed_work_directory(root)
    if _UNDO_ID_PATTERN.fullmatch(undo_id) is None:
        raise ValueError("恢复点标识无效")
    if _MANAGEMENT_OPERATION_ID_PATTERN.fullmatch(recovery_operation_id) is None:
        raise ValueError("恢复操作标识无效")
    backup_root = _backup_root(root)
    backup_directory = backup_root / "executions" / undo_id
    manifest, validation_error = _validate_execution_directory(
        backup_directory, backup_root
    )
    if manifest is None:
        raise ValueError(f"恢复资料无效：{validation_error}")
    if manifest.get("version") != BACKUP_MANIFEST_VERSION:
        raise ValueError("旧版恢复资料只能导出副本后人工核对")
    revision = max(int(manifest.get("revision", 0) or 0), 0)
    if revision != expected_revision:
        raise ValueError("恢复资料状态已变化，请刷新后重新确认")
    allowed_statuses = frozenset(
        {
            "preparing",
            "prepared",
            "applying",
            "ready",
            "recovering",
            "needs_attention",
            "restored",
        }
    )
    if manifest.get("status") not in allowed_statuses:
        raise ValueError("恢复资料当前状态不支持自动恢复，请先导出副本核对")

    targets: list[_FolderUndoTarget] = []
    raw_files = manifest.get("files")
    if not isinstance(raw_files, list) or not raw_files:
        raise ValueError("恢复资料没有可核对的文件清单")
    root_absolute = Path(os.path.abspath(root))
    for raw_item in raw_files:
        if not isinstance(raw_item, dict):
            raise ValueError("恢复资料包含无效文件项")
        relative = _safe_manifest_relative_path(raw_item.get("originalRelativePath"))
        if relative is None:
            raise ValueError("恢复资料包含无效目标路径")
        target = Path(os.path.abspath(root / relative))
        if (
            not _is_relative_to(target, root_absolute)
            or _has_reparse_point(target, root)
        ):
            raise ValueError("恢复目标路径越界或包含 reparse point")
        post_fingerprint = raw_item.get("postFingerprint")
        prior_fingerprint = raw_item.get("priorFingerprint")
        if (
            not isinstance(post_fingerprint, str)
            or re.fullmatch(r"[0-9a-f]{64}", post_fingerprint) is None
        ):
            raise ValueError("恢复资料缺少有效的执行后文件指纹")
        if prior_fingerprint is not None and (
            not isinstance(prior_fingerprint, str)
            or re.fullmatch(r"[0-9a-f]{64}", prior_fingerprint) is None
        ):
            raise ValueError("恢复资料包含无效的执行前文件指纹")
        recovery_action = raw_item.get("recoveryAction")
        backup_relative = _safe_manifest_relative_path(
            raw_item.get("backupRelativePath")
        )
        if recovery_action not in {"restore", "remove"}:
            recovery_action = "restore" if backup_relative is not None else "remove"
        backup_path: Path | None = None
        if recovery_action == "restore":
            if backup_relative is None or prior_fingerprint is None:
                raise ValueError("恢复资料缺少执行前副本或指纹")
            backup_path = Path(os.path.abspath(backup_directory / backup_relative))
            if (
                not _is_relative_to(backup_path, backup_directory)
                or _has_reparse_point(backup_path, backup_directory)
            ):
                raise ValueError("执行前副本路径越界或包含 reparse point")
        elif prior_fingerprint is not None:
            raise ValueError("新建文件恢复项不应包含执行前文件指纹")
        targets.append(
            _FolderUndoTarget(
                path=target,
                backup_path=backup_path,
                prior_fingerprint=prior_fingerprint,
                post_fingerprint=post_fingerprint,
            )
        )

    record = _FolderUndoRecord(
        undo_id=undo_id,
        plan_id=str(manifest.get("planId", undo_id)),
        backup_directory=backup_directory,
        targets=tuple(targets),
    )
    transient_session = _FolderSession(root, {})
    with folder_write_guard(root):
        with _backup_cleanup_lock:
            if undo_id in _deleting_backup_ids:
                raise ValueError("恢复资料正在删除，无法同时恢复")
            if undo_id in _restoring_backup_ids:
                raise ValueError("这份恢复资料正在处理，请勿重复操作")
            _restoring_backup_ids.add(undo_id)
        try:
            return _restore_folder_records(
                transient_session,
                [record],
                allowed_manifest_statuses=allowed_statuses,
                recovery_operation_id=recovery_operation_id,
            )
        finally:
            with _backup_cleanup_lock:
                _restoring_backup_ids.discard(undo_id)


def undo_folder_execution(request: FolderUndoRequest) -> FolderUndoResponse:
    session = _session(request.sessionId)
    record_index = next(
        (
            index
            for index, record in enumerate(session.undo_stack)
            if record.undo_id == request.undoId
        ),
        -1,
    )
    if record_index < 0:
        raise ValueError("这次文件夹执行的撤销数据已失效或已使用")
    restore_ids = {
        record.undo_id for record in session.undo_stack[record_index:]
    }
    with folder_write_guard(session.root):
        with _backup_cleanup_lock:
            if restore_ids & _deleting_backup_ids:
                raise ValueError("所需备份正在删除，无法同时撤销")
            if restore_ids & _restoring_backup_ids:
                raise ValueError("这次文件夹执行正在撤销，请勿重复操作")
            _restoring_backup_ids.update(restore_ids)
        try:
            return _restore_folder_undo_impl(request)
        finally:
            with _backup_cleanup_lock:
                _restoring_backup_ids.difference_update(restore_ids)


def _execute_folder_plan(
    request: FolderExecuteRequest,
    on_progress: FolderProgressCallback | None = None,
) -> FolderExecuteResponse:
    session = _session(request.sessionId)
    execution_id = uuid.uuid4().hex
    _emit_folder_progress(
        on_progress,
        {
            "executionId": execution_id,
            "phase": "preflight",
            "stepId": "folder.preflight",
            "status": "started",
            "title": "正在执行文件夹写入前检查",
            "summary": f"正在检查 {len(request.plan.actions)} 步操作的来源指纹、授权范围和动作能力。",
            "current": 0,
            "total": len(request.plan.actions),
            "unit": "actions",
        },
    )
    if session.source_fingerprint is not None:
        if request.plan.sourceFingerprint is None:
            raise ValueError("计划缺少数据来源指纹，请重新生成预览")
        if request.plan.sourceFingerprint != session.source_fingerprint:
            raise ValueError("计划的数据来源指纹已失效，请重新生成预览")
        current_fingerprint = _selected_files_fingerprint(
            {path for path, _ in session.selected_targets.values()}
        )
        if current_fingerprint != session.source_fingerprint:
            raise ValueError("预览后所选文件已发生变化，请重新生成预览后再执行")
    validate_plan_for_folder(
        request.plan,
        session,
        authorized_sheets=list(session.selected_targets),
    )
    _emit_folder_progress(
        on_progress,
        {
            "executionId": execution_id,
            "phase": "preflight",
            "stepId": "folder.preflight",
            "status": "completed",
            "title": "文件夹写入前检查已通过",
            "summary": f"已确认 {len(request.plan.actions)} 步操作均在当前选择的文件和工作表范围内。",
            "current": len(request.plan.actions),
            "total": len(request.plan.actions),
            "unit": "actions",
            "trace": {
                "kind": "evidence",
                "summary": "文件夹写入前检查已通过。",
                "basis": "来源指纹、已选工作表授权、动作白名单与目标有效性检查",
            },
        },
    )
    execution_started = time.perf_counter()
    workbooks: dict[Path, Workbook] = {}
    modified_paths: set[Path] = set()
    action_results: list[ActionExecutionResult] = []
    dynamic_criteria_by_action: dict[
        int, list[tuple[Path, VerificationCriterion]]
    ] = {}

    def complete_action(
        index: int,
        action,
        path: Path,
        sheet_name: str,
        summary: str | None = None,
    ) -> None:
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "executing",
                "stepId": f"folder.action.{index + 1}",
                "status": "completed",
                "title": f"第 {index + 1} / {len(request.plan.actions)} 步已完成",
                "summary": summary or f"{action.type} 已应用到 {action.sheet}",
                "current": index + 1,
                "total": len(request.plan.actions),
                "unit": "actions",
                "actionIndex": index + 1,
                "actionCount": len(request.plan.actions),
                "actionType": action.type,
                "target": {
                    "file": str(path.relative_to(session.root)),
                    "sheet": sheet_name,
                },
            },
        )

    for index, action in enumerate(request.plan.actions):
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "executing",
                "stepId": f"folder.action.{index + 1}",
                "status": "started",
                "title": f"正在执行第 {index + 1} / {len(request.plan.actions)} 步",
                "summary": f"{action.type} · {action.sheet}",
                "current": index,
                "total": len(request.plan.actions),
                "unit": "actions",
                "actionIndex": index + 1,
                "actionCount": len(request.plan.actions),
                "actionType": action.type,
                "target": {"sheet": action.sheet},
            },
        )
        path, sheet_name, _ = _target(session, action.sheet)
        workbook = workbooks.setdefault(path, _load_for_write(path))

        if action.type == "deleteWorksheet":
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"未找到工作表「{sheet_name}」")
            if len(workbook.sheetnames) == 1:
                raise ValueError("不能删除工作簿中唯一的工作表")
            workbook.remove(workbook[sheet_name])
            modified_paths.add(path)
            action_results.append(
                ActionExecutionResult(
                    index=index,
                    type=action.type,
                    sheet=action.sheet,
                    status="succeeded",
                    changed=True,
                )
            )
            complete_action(index, action, path, sheet_name)
            continue

        if action.type == "splitGroupAggregate":
            action_criteria = dynamic_criteria_by_action.setdefault(index, [])
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"未找到源工作表「{sheet_name}」")
            source_sheet = workbook[sheet_name]
            if action.sourceRange:
                min_col, min_row, max_col, max_row = range_boundaries(
                    action.sourceRange
                )
            else:
                min_col, min_row = 1, 1
                max_col, max_row = source_sheet.max_column, source_sheet.max_row
            values = [
                [
                    source_sheet.cell(row=row, column=column).value
                    for column in range(min_col, max_col + 1)
                ]
                for row in range(min_row, max_row + 1)
            ]

            def normalized_header(value) -> str:
                return str(value or "").strip().casefold()

            required_fields = [
                action.splitBy,
                *action.groupBy,
                *[
                    metric.field
                    for metric in action.metrics
                    if metric.field is not None
                ],
            ]
            required_normalized = {
                normalized_header(field) for field in required_fields
            }
            header_offset = next(
                (
                    row_index
                    for row_index, row in enumerate(values)
                    if required_normalized.issubset(
                        {
                            normalized_header(value)
                            for value in row
                            if normalized_header(value)
                        }
                    )
                ),
                None,
            )
            if header_offset is None:
                raise ValueError(
                    f"未找到包含这些字段的表头：{'、'.join(required_fields)}"
                )
            header_indexes: dict[str, int] = {}
            for column_index, value in enumerate(values[header_offset]):
                normalized = normalized_header(value)
                if normalized and normalized not in header_indexes:
                    header_indexes[normalized] = column_index

            def field_index(field: str) -> int:
                try:
                    return header_indexes[normalized_header(field)]
                except KeyError as error:
                    raise ValueError(f"未找到字段「{field}」") from error

            split_index = field_index(action.splitBy)
            group_indexes = [field_index(field) for field in action.groupBy]
            metric_indexes = [
                field_index(metric.field) if metric.field else None
                for metric in action.metrics
            ]

            def is_blank(value) -> bool:
                return value is None or (
                    isinstance(value, str) and not value.strip()
                )

            def value_key(value):
                return (type(value).__name__, str(value).strip())

            groups: dict[tuple, dict] = {}
            splits: dict[tuple, dict] = {}
            for row in values[header_offset + 1 :]:
                split_value = row[split_index]
                group_values = [row[column] for column in group_indexes]
                if (
                    not action.includeBlankSplitValues
                    and is_blank(split_value)
                ) or any(is_blank(value) for value in group_values):
                    continue
                group_key = tuple(value_key(value) for value in group_values)
                split_key = (value_key(split_value),)
                group = groups.setdefault(
                    group_key,
                    {
                        "values": group_values,
                        "totals": [0.0 for _ in action.metrics],
                    },
                )
                split = splits.setdefault(
                    split_key, {"value": split_value, "pairs": {}}
                )
                pair = split["pairs"].setdefault(
                    group_key, [0.0 for _ in action.metrics]
                )
                for metric_index, metric in enumerate(action.metrics):
                    cell = (
                        None
                        if metric_indexes[metric_index] is None
                        else row[metric_indexes[metric_index]]
                    )
                    if metric.operation == "countRows":
                        increment = 1.0
                    elif metric.operation == "countNonBlank":
                        increment = 0.0 if is_blank(cell) else 1.0
                    else:
                        try:
                            increment = float(
                                str(cell or "").replace(",", "").strip()
                            )
                        except ValueError:
                            increment = 0.0
                    group["totals"][metric_index] += increment
                    pair[metric_index] += increment

            if not splits:
                raise ValueError("表头下没有可拆分的有效数据行")
            if len(splits) > action.maxOutputSheets:
                raise ValueError(
                    f"将生成 {len(splits)} 个工作表，"
                    f"超过安全上限 {action.maxOutputSheets} 个"
                )
            _emit_folder_progress(
                on_progress,
                {
                    "executionId": execution_id,
                    "phase": "executing",
                    "stepId": f"folder.action.{index + 1}.sheets",
                    "status": "progress",
                    "title": "正在生成拆分工作表",
                    "summary": f"已计算出 {len(splits)} 张结果工作表，正在写入内存工作簿。",
                    "current": 0,
                    "total": len(splits),
                    "unit": "sheets",
                    "actionIndex": index + 1,
                    "actionCount": len(request.plan.actions),
                    "actionType": action.type,
                    "target": {
                        "file": str(path.relative_to(session.root)),
                        "sheet": sheet_name,
                    },
                },
            )

            headers = [*action.groupBy, action.splitBy]
            ratio_columns: list[int] = []
            for metric in action.metrics:
                headers.append(metric.outputName)
                if metric.ratioOutputName:
                    headers.append(metric.ratioOutputName)
                    ratio_columns.append(len(headers))

            existing_names = {
                name.casefold(): name for name in workbook.sheetnames
            }
            generated_names: set[str] = set()
            split_changed = False

            def safe_name(value) -> str:
                cleaned = re.sub(
                    r"\s+",
                    " ",
                    re.sub(
                        r"[\x00-\x1f:：\\＼/／?？*＊\[\]［］]",
                        " ",
                        str(value or ""),
                    ),
                ).strip().strip("'").strip()
                cleaned = cleaned or "未命名"
                if cleaned.casefold() == "history":
                    cleaned = "History 结果"
                return cleaned[:31]

            def renamed(base: str) -> str:
                if base.casefold() not in existing_names:
                    return base
                suffix = 2
                while True:
                    marker = f" ({suffix})"
                    candidate = f"{base[:31 - len(marker)]}{marker}"
                    if candidate.casefold() not in existing_names:
                        return candidate
                    suffix += 1

            for split_number, split in enumerate(splits.values(), start=1):
                base_name = safe_name(split["value"])
                existing_name = existing_names.get(base_name.casefold())
                collides_with_current_run = base_name.casefold() in generated_names
                if (
                    existing_name
                    and action.existingSheetPolicy == "skip"
                    and not collides_with_current_run
                ):
                    action_criteria.append(
                        (
                            path,
                            WorksheetExistsCriterion(
                                type="worksheetExists", sheet=existing_name
                            ),
                        )
                    )
                    _emit_folder_progress(
                        on_progress,
                        {
                            "executionId": execution_id,
                            "phase": "executing",
                            "stepId": f"folder.action.{index + 1}.sheets",
                            "status": "progress",
                            "title": "正在生成拆分工作表",
                            "summary": f"已处理 {split_number} / {len(splits)} 张结果工作表。",
                            "current": split_number,
                            "total": len(splits),
                            "unit": "sheets",
                            "actionIndex": index + 1,
                            "actionCount": len(request.plan.actions),
                            "actionType": action.type,
                            "target": {
                                "file": str(path.relative_to(session.root)),
                                "sheet": existing_name,
                            },
                        },
                    )
                    continue
                split_changed = True
                if (
                    existing_name
                    and action.existingSheetPolicy == "replace"
                    and not collides_with_current_run
                ):
                    if existing_name.casefold() == sheet_name.casefold():
                        raise ValueError(
                            f"不能用拆分结果覆盖源工作表「{sheet_name}」"
                        )
                    position = workbook.sheetnames.index(existing_name)
                    workbook.remove(workbook[existing_name])
                    target_sheet = workbook.create_sheet(base_name, position)
                    target_name = base_name
                else:
                    target_name = renamed(base_name)
                    target_sheet = workbook.create_sheet(target_name)
                existing_names[target_name.casefold()] = target_name
                generated_names.add(target_name.casefold())

                matrix: list[list[object]] = [list(headers)]
                for group_key, aggregates in split["pairs"].items():
                    group = groups[group_key]
                    output_row = [*group["values"], split["value"]]
                    for metric_index, metric in enumerate(action.metrics):
                        aggregate = aggregates[metric_index]
                        if metric.operation != "sum" and aggregate.is_integer():
                            aggregate = int(aggregate)
                        output_row.append(aggregate)
                        if metric.ratioOutputName:
                            total = group["totals"][metric_index]
                            output_row.append(
                                None if total == 0 else aggregate / total
                            )
                    matrix.append(output_row)
                for output_row in matrix:
                    target_sheet.append(output_row)
                for cell in target_sheet[1]:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill("solid", fgColor="DFF3E4")
                for column in ratio_columns:
                    for row in range(2, target_sheet.max_row + 1):
                        target_sheet.cell(row=row, column=column).number_format = (
                            "0.00%"
                        )
                for column in range(1, target_sheet.max_column + 1):
                    letter = get_column_letter(column)
                    width = max(
                        len(str(target_sheet.cell(row=row, column=column).value or ""))
                        for row in range(1, target_sheet.max_row + 1)
                    )
                    target_sheet.column_dimensions[letter].width = min(width + 2, 40)
                action_criteria.extend(
                    _matrix_verification_criteria(path, target_name, matrix)
                )
                _emit_folder_progress(
                    on_progress,
                    {
                        "executionId": execution_id,
                        "phase": "executing",
                        "stepId": f"folder.action.{index + 1}.sheets",
                        "status": "progress",
                        "title": "正在生成拆分工作表",
                        "summary": f"已处理 {split_number} / {len(splits)} 张结果工作表。",
                        "current": split_number,
                        "total": len(splits),
                        "unit": "sheets",
                        "actionIndex": index + 1,
                        "actionCount": len(request.plan.actions),
                        "actionType": action.type,
                        "target": {
                            "file": str(path.relative_to(session.root)),
                            "sheet": target_name,
                        },
                    },
                )

            if split_changed:
                modified_paths.add(path)
            action_results.append(
                ActionExecutionResult(
                    index=index,
                    type=action.type,
                    sheet=action.sheet,
                    status="succeeded",
                    changed=split_changed,
                )
            )
            complete_action(
                index,
                action,
                path,
                sheet_name,
                (
                    f"已生成或处理 {len(splits)} 张拆分工作表。"
                    if split_changed
                    else "所有结果工作表均已存在并按策略跳过，文件内容未变化。"
                ),
            )
            continue

        sheet_existed = sheet_name in workbook.sheetnames
        active_sheet_before = workbook.active.title
        sheet = _get_or_create_sheet(workbook, sheet_name)
        if action.type in {"createWorksheet", "activateWorksheet"}:
            action_changed = not sheet_existed
            if action.type == "activateWorksheet":
                if active_sheet_before != sheet.title:
                    workbook.active = workbook.sheetnames.index(sheet.title)
                    action_changed = True
            if action_changed:
                modified_paths.add(path)
            action_results.append(
                ActionExecutionResult(
                    index=index,
                    type=action.type,
                    sheet=action.sheet,
                    status="succeeded",
                    changed=action_changed,
                )
            )
            complete_action(
                index,
                action,
                path,
                sheet_name,
                None if action_changed else "目标状态已经符合要求，文件内容未变化。",
            )
            continue
        action_changed = True
        if action.type == "writeTable":
            rows = [action.headers, *action.rows]
            start_column, start_row, _, _ = range_boundaries(action.startCell)
            normalized_rows = [
                [_force_text_if_formula(value) for value in row] for row in rows
            ]
            end_column = start_column + len(normalized_rows[0]) - 1
            end_row = start_row + len(normalized_rows) - 1
            target_range = (
                f"{get_column_letter(start_column)}{start_row}:"
                f"{get_column_letter(end_column)}{end_row}"
            )
            action_changed = not sheet_existed or _matrix_would_change(
                sheet, target_range, normalized_rows
            )
            if action_changed:
                _write_matrix(sheet, target_range, normalized_rows)
        elif action.type == "writeValues":
            normalized_values = [
                [_force_text_if_formula(v) for v in row] for row in action.values
            ]
            action_changed = not sheet_existed or _matrix_would_change(
                sheet, action.range, normalized_values
            )
            if action_changed:
                _write_matrix(sheet, action.range, normalized_values)
        elif action.type == "writeFormulas":
            action_changed = not sheet_existed or _matrix_would_change(
                sheet, action.range, action.formulas
            )
            if action_changed:
                _write_matrix(sheet, action.range, action.formulas)
        elif action.type == "clearRange":
            cells = [cell for row in _iter_range(sheet, action.range) for cell in row]
            if action.applyTo == "contents":
                action_changed = not sheet_existed or any(
                    cell.value not in (None, "") for cell in cells
                )
            elif action.applyTo == "hyperlinks":
                action_changed = not sheet_existed or any(
                    cell.hyperlink is not None for cell in cells
                )
            if action.applyTo in {"all", "contents"}:
                for cell in cells:
                    cell.value = None
            if action.applyTo in {"all", "formats"}:
                for cell in cells:
                    cell._style = copy(sheet["A1"]._style)
                    cell.number_format = "General"
                    cell.alignment = Alignment()
            if action.applyTo in {"all", "hyperlinks"}:
                for cell in cells:
                    cell.hyperlink = None
        elif action.type in {"insertRange", "deleteRange"}:
            _folder_mode_unsupported(action.type)
        elif action.type == "copyRange":
            source_path, source_name, _ = _target(session, action.sourceSheet)
            source_book = workbooks.setdefault(source_path, _load_for_write(source_path))
            if source_name not in source_book.sheetnames:
                raise ValueError(f"未找到源工作表「{source_name}」")
            source_sheet = source_book[source_name]
            source_cells = [list(row) for row in _iter_range(source_sheet, action.sourceRange)]
            target_min_col, target_min_row, _, _ = range_boundaries(action.targetRange)
            if action.transpose:
                source_cells = [list(row) for row in zip(*source_cells)]
            for row_offset, row in enumerate(source_cells):
                for column_offset, source_cell in enumerate(row):
                    target_cell = sheet.cell(
                        row=target_min_row + row_offset,
                        column=target_min_col + column_offset,
                    )
                    if action.skipBlanks and source_cell.value in (None, ""):
                        continue
                    if action.copyType in {"all", "values", "formulas"}:
                        target_cell.value = source_cell.value
                    if action.copyType in {"all", "formats"}:
                        target_cell._style = copy(source_cell._style)
                        target_cell.number_format = source_cell.number_format
                    if action.copyType == "link":
                        target_cell.value = (
                            f"='{source_name}'!{source_cell.coordinate}"
                        )
        elif action.type == "sortRange":
            min_col, min_row, max_col, max_row = range_boundaries(action.range)
            first_data_row = min_row + (1 if action.hasHeaders else 0)
            rows = [
                [sheet.cell(row=row, column=col).value for col in range(min_col, max_col + 1)]
                for row in range(first_data_row, max_row + 1)
            ]
            if rows and any(key.column >= len(rows[0]) for key in action.keys):
                raise ValueError("排序列索引超出范围")
            sorted_rows = sorted(
                rows,
                key=cmp_to_key(
                    lambda left, right: _compare_sorted_rows(
                        left, right, action.keys
                    )
                )
            )
            action_changed = not sheet_existed or rows != sorted_rows
            if action_changed:
                for row_offset, values in enumerate(sorted_rows):
                    for column_offset, value in enumerate(values):
                        sheet.cell(
                            row=first_data_row + row_offset,
                            column=min_col + column_offset,
                            value=value,
                        )
        elif action.type == "filterRange":
            sheet.auto_filter.ref = action.range
            sheet.auto_filter.filterColumn = [
                FilterColumn(
                    colId=action.column,
                    filters=Filters(filter=[str(value) for value in action.values]),
                )
            ]
        elif action.type == "clearFilter":
            action_changed = not sheet_existed or bool(
                sheet.auto_filter.ref or sheet.auto_filter.filterColumn
            )
            sheet.auto_filter.ref = None
            sheet.auto_filter.filterColumn = []
        elif action.type == "setDataValidation":
            operator_map = {
                "equalTo": "equal",
                "notEqualTo": "notEqual",
                "greaterThanOrEqualTo": "greaterThanOrEqual",
                "lessThanOrEqualTo": "lessThanOrEqual",
            }
            validation_type = {
                "wholeNumber": "whole",
                "textLength": "textLength",
            }.get(action.validationType, action.validationType)
            if validation_type == "list":
                formula1 = '"' + ",".join(str(value) for value in action.values) + '"'
            else:
                formula1 = action.formula1
            validation = DataValidation(
                type=validation_type,
                operator=operator_map.get(action.operator, action.operator),
                formula1=formula1,
                formula2=action.formula2,
                allow_blank=action.allowBlank,
                prompt=action.prompt,
                error=action.errorMessage,
            )
            sheet.add_data_validation(validation)
            validation.add(action.range)
        elif action.type == "setConditionalFormat":
            fill = PatternFill(
                "solid", fgColor=(action.color or "#FFF2CC").removeprefix("#")
            )
            if action.ruleType == "colorScale":
                rule = ColorScaleRule(
                    start_type="min",
                    start_color=(action.minColor or "#F8696B").removeprefix("#"),
                    mid_type="percentile" if action.midColor else None,
                    mid_value=50 if action.midColor else None,
                    mid_color=action.midColor.removeprefix("#") if action.midColor else None,
                    end_type="max",
                    end_color=(action.maxColor or "#63BE7B").removeprefix("#"),
                )
            elif action.ruleType == "custom":
                rule = FormulaRule(formula=[str(action.formula1 or "")], fill=fill)
            else:
                rule = CellIsRule(
                    operator={
                        "equalTo": "equal",
                        "notEqualTo": "notEqual",
                        "greaterThanOrEqualTo": "greaterThanOrEqual",
                        "lessThanOrEqualTo": "lessThanOrEqual",
                    }.get(action.operator or "equalTo", action.operator or "equal"),
                    formula=[
                        str(action.formula1 or ""),
                        *(
                            [str(action.formula2)]
                            if action.formula2 is not None
                            else []
                        ),
                    ],
                    fill=fill,
                )
            sheet.conditional_formatting.add(action.range, rule)
        elif action.type == "setNumberFormat":
            cells = [cell for row in _iter_range(sheet, action.range) for cell in row]
            action_changed = not sheet_existed or any(
                cell.number_format != action.formatCode for cell in cells
            )
            if action_changed:
                for cell in cells:
                    cell.number_format = action.formatCode
        elif action.type == "setBorders":
            side = Side(
                style=_openpyxl_border_style(action.style, action.weight),
                color=action.color.removeprefix("#"),
            )
            for row in _iter_range(sheet, action.range):
                for cell in row:
                    current = cell.border
                    cell.border = Border(
                        left=side if "left" in action.sides else current.left,
                        right=side if "right" in action.sides else current.right,
                        top=side if "top" in action.sides else current.top,
                        bottom=side if "bottom" in action.sides else current.bottom,
                        vertical=side if "insideVertical" in action.sides else current.vertical,
                        horizontal=side if "insideHorizontal" in action.sides else current.horizontal,
                    )
        elif action.type == "setAlignment":
            for row in _iter_range(sheet, action.range):
                for cell in row:
                    cell.alignment = copy(cell.alignment)
                    if action.horizontal is not None:
                        cell.alignment = Alignment(
                            horizontal={
                                "centerAcrossSelection": "centerContinuous"
                            }.get(action.horizontal, action.horizontal),
                            vertical=cell.alignment.vertical,
                            wrap_text=cell.alignment.wrap_text,
                        )
                    if action.vertical is not None or action.wrapText is not None:
                        cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=action.vertical or cell.alignment.vertical,
                            wrap_text=(
                                action.wrapText
                                if action.wrapText is not None
                                else cell.alignment.wrap_text
                            ),
                        )
        elif action.type == "mergeCells":
            sheet.merge_cells(action.range)
        elif action.type == "unmergeCells":
            sheet.unmerge_cells(action.range)
        elif action.type == "resizeRange":
            min_col, min_row, max_col, max_row = range_boundaries(action.range)
            if action.rowHeight is not None:
                for row in range(min_row, max_row + 1):
                    sheet.row_dimensions[row].height = action.rowHeight
            if action.columnWidth is not None:
                for column in range(min_col, max_col + 1):
                    sheet.column_dimensions[get_column_letter(column)].width = action.columnWidth
        elif action.type == "freezePanes":
            current_freeze = sheet.freeze_panes
            if hasattr(current_freeze, "coordinate"):
                current_freeze = current_freeze.coordinate
            if action.rows == 0 and action.columns == 0:
                desired_freeze = None
            else:
                desired_freeze = sheet.cell(
                    row=action.rows + 1, column=action.columns + 1
                ).coordinate
            action_changed = not sheet_existed or current_freeze != desired_freeze
            if action_changed:
                sheet.freeze_panes = desired_freeze
        elif action.type == "setHyperlink":
            for row in _iter_range(sheet, action.range):
                for cell in row:
                    cell.hyperlink = action.address
                    if action.text is not None:
                        cell.value = action.text
        elif action.type in {"addComment", "addNote"}:
            sheet[action.cell].comment = Comment(action.text, "Excel Bro")
        elif action.type == "createTable":
            table = Table(
                displayName=action.name or f"ExcelBroTable{index + 1}",
                ref=action.range,
            )
            if not action.hasHeaders:
                table.headerRowCount = 0
            table.tableStyleInfo = TableStyleInfo(
                name=action.style or "TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            sheet.add_table(table)
        elif action.type == "createChart":
            chart_class = _FOLDER_CHART_CLASSES.get(action.chartType)
            if chart_class is None:
                raise ValueError(
                    f"文件夹模式暂不支持图表类型「{action.chartType}」"
                )
            chart = chart_class()
            if action.chartType in {"ColumnClustered", "BarClustered"}:
                chart.type = (
                    "col" if action.chartType == "ColumnClustered" else "bar"
                )
                chart.grouping = "clustered"
            min_col, min_row, max_col, max_row = range_boundaries(action.sourceRange)
            chart.add_data(
                Reference(
                    sheet,
                    min_col=min_col,
                    max_col=max_col,
                    min_row=min_row,
                    max_row=max_row,
                ),
                titles_from_data=True,
            )
            if action.title:
                chart.title = action.title
            sheet.add_chart(chart, action.targetRange or f"{get_column_letter(max_col + 2)}{min_row}")
        elif action.type == "createPivotTable":
            _folder_mode_unsupported(action.type)
        elif action.type == "addNamedRange":
            workbook.defined_names.add(
                DefinedName(
                    action.name,
                    attr_text=(
                        f"{quote_sheetname(sheet.title)}!"
                        f"{absolute_coordinate(action.range)}"
                    ),
                    comment=action.comment,
                )
            )
        elif action.type == "addImage":
            try:
                from openpyxl.drawing.image import Image
                image = Image(io.BytesIO(base64.b64decode(action.base64)))
            except Exception as error:
                raise ValueError(f"无法解析图片数据：{error}") from error
            sheet.add_image(image, action.targetRange.split(":")[0])
        elif action.type == "addShape":
            _folder_mode_unsupported(action.type)
        elif action.type == "setFill":
            min_column, min_row, max_column, max_row = range_boundaries(action.range)
            fill = PatternFill("solid", fgColor=action.color.removeprefix("#"))
            for row in sheet.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_column,
                max_col=max_column,
            ):
                for cell in row:
                    cell.fill = fill
        elif action.type == "setFont":
            min_column, min_row, max_column, max_row = range_boundaries(action.range)
            for row in sheet.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_column,
                max_col=max_column,
            ):
                for cell in row:
                    cell.font = Font(
                        name=cell.font.name,
                        size=cell.font.size,
                        bold=action.bold
                        if action.bold is not None
                        else cell.font.bold,
                        italic=cell.font.italic,
                        color=action.color.removeprefix("#")
                        if action.color
                        else cell.font.color,
                    )
        elif action.type == "autofit":
            min_column, _, max_column, _ = range_boundaries(action.range)
            for column in range(min_column, max_column + 1):
                letter = get_column_letter(column)
                width = max(
                    (
                        len(str(cell.value))
                        for cell in sheet[letter]
                        if cell.value is not None
                    ),
                    default=8,
                )
                sheet.column_dimensions[letter].width = min(width + 2, 60)
        else:
            raise ValueError(f"文件夹模式未实现动作：{action.type}")
        if action_changed:
            modified_paths.add(path)
        action_results.append(
            ActionExecutionResult(
                index=index,
                type=action.type,
                sheet=action.sheet,
                status="succeeded",
                changed=action_changed,
            )
        )
        complete_action(
            index,
            action,
            path,
            sheet_name,
            None if action_changed else "目标状态已经符合要求，文件内容未变化。",
        )

    with _backup_cleanup_lock:
        _active_backup_ids.add(execution_id)
    try:
        (
            modified,
            backups,
            undo_targets,
            backup_directory,
            backup_warnings,
        ) = _save_workbooks_atomically(
            session,
            workbooks,
            modified_paths,
            execution_id,
            request.plan.id,
            request.plan.title,
            on_progress,
        )
        undo_reference = (
            _register_folder_undo(
                session,
                request.sessionId,
                execution_id,
                request.plan.id,
                backup_directory,
                undo_targets,
            )
            if backup_directory is not None
            else None
        )

        verification_started = time.perf_counter()
        execution_ms = (verification_started - execution_started) * 1000
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "verifying",
                "stepId": "folder.verify",
                "status": "started",
                "title": "正在重新打开文件并独立验证结果",
            },
        )
        verification = _verify_folder_plan(
            session, request.plan, dynamic_criteria_by_action
        )
        verification_ms = (time.perf_counter() - verification_started) * 1000
        passed_checks = sum(1 for check in verification.checks if check.passed)
        _emit_folder_progress(
            on_progress,
            {
                "executionId": execution_id,
                "phase": "verifying",
                "stepId": "folder.verify",
                "status": "warning" if verification.status == "failed" else "completed",
                "title": (
                    "执行已完成，但验证发现不一致"
                    if verification.status == "failed"
                    else "文件写入结果验证完成"
                ),
                "summary": f"{passed_checks} / {len(verification.checks)} 项检查符合预期。",
                "current": passed_checks,
                "total": len(verification.checks),
                "unit": "checks",
                "trace": {
                    "kind": "conflict" if verification.status == "failed" else "evidence",
                    "summary": f"{passed_checks} / {len(verification.checks)} 项检查符合预期。",
                    "basis": "保存后重新打开目标工作簿并按计划验收条件比较",
                },
            },
        )
        return FolderExecuteResponse(
            executionId=execution_id,
            filesModified=modified,
            backups=backups,
            backupDirectory=(
                str(_backup_root(session.root)) if backup_directory is not None else None
            ),
            backupWarnings=backup_warnings,
            actionResults=action_results,
            verification=verification,
            undo=undo_reference,
            executionMs=execution_ms,
            verificationMs=verification_ms,
        )
    finally:
        with _backup_cleanup_lock:
            _active_backup_ids.discard(execution_id)


def execute_folder_plan(
    request: FolderExecuteRequest,
    on_progress: FolderProgressCallback | None = None,
) -> FolderExecuteResponse:
    session = _session(request.sessionId)
    with folder_write_guard(session.root):
        opened: list[Workbook] = []
        token = _OPEN_WRITE_WORKBOOKS.set(opened)
        try:
            return _execute_folder_plan(request, on_progress)
        finally:
            for workbook in opened:
                try:
                    workbook.close()
                except Exception:
                    pass
            _OPEN_WRITE_WORKBOOKS.reset(token)

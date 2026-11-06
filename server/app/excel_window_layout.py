"""Local Excel review layout and workbook-scoped cleanup.

The compatibility copy is archived locally before closing the same verified COM
workbook. No window close messages, application creation or arbitrary execution.
"""
from __future__ import annotations

import hashlib
import ntpath
import re
import zipfile
import secrets
import sys
import threading
import time
from dataclasses import dataclass

from .app_data import application_data_root
from .capabilities import capability_float, capability_int
from .models import (
    ExcelReviewCopyCloseRequest,
    ExcelReviewCopyCloseResult,
    ExcelReviewLayoutPrepareRequest,
    ExcelReviewLayoutResult,
    ExcelReviewPaneMeasurement,
)


class WindowLayoutError(Exception):
    def __init__(self, code: str):
        self.code = code
        # Never expose COM exceptions, titles or full paths.
        super().__init__("无法安全完成本次窗口排列；原表未因此写入，请手动并排查看。")


def path_hash(path: str) -> str:
    canonical = ntpath.normpath(path).replace("\\", "/").lower()
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


@dataclass(frozen=True)
class WindowKey:
    hwnd: int
    pid: int


@dataclass(frozen=True)
class WindowInfo:
    key: WindowKey
    workbook_hash: str
    sheet_hash: str
    sheet_name: str
    unsaved: bool
    review_copy_id: str = ""


@dataclass(frozen=True)
class CopyCloseReceipt:
    accepted: bool
    recovery_path: str | None = None


@dataclass
class _Ticket:
    origin: str
    expires: float
    request: ExcelReviewLayoutPrepareRequest
    original: WindowKey
    baseline: frozenset[WindowKey]


class ExcelWindowLayout:
    def __init__(self, backend=None, clock=time.monotonic, sleep=time.sleep):
        self.backend = backend  # Injected fake in tests; native created lazily.
        self.clock, self.sleep = clock, sleep
        self._tickets: dict[str, _Ticket] = {}
        self._lock = threading.Lock()

    def _native(self):
        if self.backend is None:
            self.backend = WindowsExcelWindows()
        return self.backend

    def prepare(self, request: ExcelReviewLayoutPrepareRequest, origin: str) -> str:
        with self._lock:
            self._tickets = {k: v for k, v in self._tickets.items() if v.expires > self.clock()}
            if len(self._tickets) >= capability_int("excelWindowLayout", "maxTickets"):
                raise WindowLayoutError("LAYOUT_BUSY")
            native = self._native()
            baseline = frozenset(native.keys())
            matches = [key for key in baseline if self._original(native.describe(key), request)]
            if len(matches) != 1:
                raise WindowLayoutError("LAYOUT_ORIGINAL_AMBIGUOUS")
            token = secrets.token_urlsafe(32)
            self._tickets[token] = _Ticket(origin, self.clock() + capability_int("excelWindowLayout", "ticketTtlSeconds"),
                                           request.model_copy(deep=True), matches[0], baseline)
            return token

    @staticmethod
    def _original(info: WindowInfo | None, request: ExcelReviewLayoutPrepareRequest) -> bool:
        return bool(info and not info.unsaved and info.workbook_hash == request.workbookPathHash
                    and info.sheet_hash == request.worksheetNameHash)

    @staticmethod
    def _copy(info: WindowInfo | None, request) -> bool:
        if not info:
            return False
        if info.review_copy_id:
            return (request.copyIdentityVersion == 2 and info.review_copy_id == request.copySheetName
                    and (info.sheet_hash == request.worksheetNameHash or info.sheet_name == request.copySheetName))
        # Existing pending reviews remain addressable with their old visible nonce.
        return info.sheet_name == request.copySheetName

    def measure(self, request: ExcelReviewLayoutPrepareRequest) -> ExcelReviewPaneMeasurement:
        # Read-only, no ticket creation/consumption, window mutation or persistence.
        # A unique saved original AND this unsaved nonce copy must still exist.
        with self._lock:
            native = self._native()
            observations = [native.describe(key) for key in native.keys()]
            originals = [info for info in observations if self._original(info, request)]
            copies = [info for info in observations if info and info.unsaved and self._copy(info, request)]
            if len(originals) != 1 or len(copies) != 1:
                raise WindowLayoutError("LAYOUT_MEASUREMENT_AMBIGUOUS")
            original, copy = originals[0], copies[0]
            width = native.client_width_css_px(original.key)
            work_area = native.work_area_width_css_px(original.key)
            if native.describe(original.key) != original or native.describe(copy.key) != copy:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            return ExcelReviewPaneMeasurement(clientWidthCssPx=width, workAreaWidthCssPx=work_area)

    def view_copy(self, request: ExcelReviewLayoutPrepareRequest, action: str):
        """Reuse or park an existing nonce copy. Never close, save or write cells."""
        with self._lock:
            native = self._native()
            observations = [native.describe(key) for key in native.keys()]
            originals = [info for info in observations if self._original(info, request)]
            copies = [info for info in observations if self._copy(info, request)]
            if len(originals) != 1 or len(copies) != 1 or not copies[0].unsaved:
                raise WindowLayoutError("LAYOUT_COPY_AMBIGUOUS")
            original, copy = originals[0], copies[0]
            if native.describe(original.key) != original or native.describe(copy.key) != copy:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            if action == "resume":
                if request.copyIdentityVersion == 2 and copy.sheet_name == request.copySheetName:
                    if not hasattr(native, 'promote_copy_label'):
                        raise WindowLayoutError("LAYOUT_UNAVAILABLE")
                    native.promote_copy_label(original, copy, request.copySheetName)
                    copy = native.describe(copy.key)
                    if not self._copy(copy, request) or copy.sheet_hash != original.sheet_hash:
                        raise WindowLayoutError("LAYOUT_COPY_CHANGED")
                return native.position_pair(original.key, copy.key, original, copy)
            if action != "suspend" or not hasattr(native, "suspend_pair"):
                raise WindowLayoutError("LAYOUT_UNAVAILABLE")
            native.suspend_pair(original, copy)
            return {"status": "suspended", "copyRetained": True}

    def select_copy_sheet(self, request: ExcelReviewLayoutPrepareRequest):
        """Select a named snapshot tab in an owned workbook; never move windows or write cells."""
        with self._lock:
            native = self._native()
            if request.copyIdentityVersion != 2:
                raise WindowLayoutError("LAYOUT_COPY_CHANGED")
            observations = [native.describe(key) for key in native.keys()]
            originals = [info for info in observations if self._original(info, request)]
            copies = [info for info in observations if info and info.review_copy_id == request.copySheetName]
            if len(originals) != 1 or len(copies) != 1 or not copies[0].unsaved:
                raise WindowLayoutError("LAYOUT_COPY_AMBIGUOUS")
            original, copy = originals[0], copies[0]
            if native.describe(original.key) != original or native.describe(copy.key) != copy:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            native.select_pair_sheet(original, copy)
            selected = native.describe(copy.key)
            if native.describe(original.key) != original or not self._copy(selected, request):
                raise WindowLayoutError("LAYOUT_COPY_CHANGED")
            return {"status": "selected"}

    @staticmethod
    def _copy_for_close(info: WindowInfo | None, request) -> bool:
        if not info:
            return False
        if request.copyIdentityVersion == 2:
            return info.review_copy_id == request.copySheetName
        return not info.review_copy_id and info.sheet_name == request.copySheetName

    def close_copy(self, request: ExcelReviewLayoutPrepareRequest | ExcelReviewCopyCloseRequest) -> ExcelReviewCopyCloseResult:
        """Close only the identified copy; retain a recovery file before discarding it.

        Protocol v4 is required at the HTTP boundary so a new pane cannot invoke
        an older bridge with unsafe window-level or unarchived close behavior.
        """
        native = self._native()
        if not hasattr(native, "close_workbook"):
            # The platform/host offers no verified workbook-close primitive. Degrade to
            # manual close rather than guessing (no TerminateProcess, no scripts, and
            # deliberately **no** WM_CLOSE fallback — that message targets a window,
            # not a workbook, and once aimed wrong it prompts to save the user's
            # original workbook).
            return ExcelReviewCopyCloseResult(status="unsupported", layoutVersion=4, copyClosed=False)
        observations = [native.describe(key) for key in native.keys()]
        # A saved workbook whose active sheet carries this review's nonce is the user's
        # own "save as"/replace of the copy. That file is theirs: never auto-discard it.
        if any(info and not info.unsaved and self._copy_for_close(info, request) for info in observations):
            return ExcelReviewCopyCloseResult(status="user_saved", layoutVersion=4, copyClosed=False)
        copies = [info for info in observations if info and info.unsaved and self._copy_for_close(info, request)]
        if len(copies) > 1:
            # Two distinct windows claim this session's nonce: refusing is the only safe
            # answer (guessing which is ours could close an unrelated workbook).
            return ExcelReviewCopyCloseResult(status="ambiguous", layoutVersion=4, copyClosed=False)
        if not copies:
            if request.copyIdentityVersion == 2 and any(info and info.unsaved and
                    (info.review_copy_id == request.copySheetName or info.sheet_hash == request.worksheetNameHash)
                    for info in observations):
                return ExcelReviewCopyCloseResult(status="identity_lost", layoutVersion=4, copyClosed=False)
            # Idempotent: repeat cleanup, or the user already closed the copy themselves.
            return ExcelReviewCopyCloseResult(status="already_closed", layoutVersion=4, copyClosed=True)
        originals = [info for info in observations if bool(info and not info.unsaved and info.workbook_hash == request.workbookPathHash)]
        if len(originals) != 1:
            return ExcelReviewCopyCloseResult(status="identity_lost", layoutVersion=4, copyClosed=False)
        copy, original = copies[0], originals[0]
        # Re-verify both identities before handing the copy to the workbook adapter.
        if native.describe(copy.key) != copy or native.describe(original.key) != original:
            return ExcelReviewCopyCloseResult(status="identity_lost", layoutVersion=4, copyClosed=False)
        if hasattr(native, "prepare_close_pair"):
            # Restore the verified original window before the owned copy disappears.
            # This activates the window, never changes either active worksheet.
            native.prepare_close_pair(original, copy)
            if native.describe(copy.key) != copy or native.describe(original.key) != original:
                return ExcelReviewCopyCloseResult(status="identity_lost", layoutVersion=4, copyClosed=False)
        receipt = native.close_workbook(copy.key, copy)
        if not receipt.accepted:
            return ExcelReviewCopyCloseResult(status="failed", layoutVersion=4, copyClosed=False,
                                             recoveryCopyPath=receipt.recovery_path)
        deadline = self.clock() + capability_float("reviewCopyLifecycle", "closeWaitSeconds")
        interval = capability_int("reviewCopyLifecycle", "closePollIntervalMs") / 1000
        while self.clock() < deadline:
            if native.describe(copy.key) is None:
                return ExcelReviewCopyCloseResult(status="closed", layoutVersion=4, copyClosed=True, recoveryCopyPath=receipt.recovery_path)
            self.sleep(interval)
        # Still on screen after the deadline. We have NO dialog evidence, so we must not
        # infer that the host is asking to save: a slow close or an unrelated modal looks
        # exactly the same from here. Report "close not confirmed" and let the user look.
        return ExcelReviewCopyCloseResult(status="close_unconfirmed", layoutVersion=4, copyClosed=False, recoveryCopyPath=receipt.recovery_path)

    def arrange(self, token: str, origin: str) -> ExcelReviewLayoutResult:
        with self._lock:
            ticket = self._tickets.get(token)
            if not ticket or ticket.origin != origin or ticket.expires <= self.clock():
                raise WindowLayoutError("LAYOUT_TICKET_UNAVAILABLE")
            # One use, even on partial positioning failure. Never replay old HWNDs.
            del self._tickets[token]
            native = self._native()
            deadline = min(ticket.expires, self.clock() + capability_float("excelWindowLayout", "copyWaitSeconds"))
            while True:
                if not self._original(native.describe(ticket.original), ticket.request):
                    raise WindowLayoutError("LAYOUT_ORIGINAL_CHANGED")
                candidates = []
                for key in set(native.keys()) - ticket.baseline:
                    info = native.describe(key)
                    if info and info.unsaved and self._copy(info, ticket.request):
                        candidates.append(key)
                if len(candidates) > 1:
                    raise WindowLayoutError("LAYOUT_COPY_AMBIGUOUS")
                if candidates:
                    copy = candidates[0]
                    # Recheck both identities immediately before any window mutation.
                    original_info = native.describe(ticket.original)
                    if not self._original(original_info, ticket.request):
                        raise WindowLayoutError("LAYOUT_ORIGINAL_CHANGED")
                    info = native.describe(copy)
                    if not info or not info.unsaved or not self._copy(info, ticket.request):
                        raise WindowLayoutError("LAYOUT_COPY_CHANGED")
                    if self.clock() >= deadline:
                        raise WindowLayoutError("LAYOUT_TICKET_UNAVAILABLE")
                    result = native.position_pair(ticket.original, copy, original_info, info,
                                                  lambda: self.clock() < deadline)
                    if not isinstance(result, ExcelReviewLayoutResult):
                        raise WindowLayoutError("LAYOUT_POSITION_UNVERIFIED")
                    return result
                if self.clock() >= deadline:
                    raise WindowLayoutError("LAYOUT_COPY_NOT_FOUND")
                self.sleep(capability_int("excelWindowLayout", "pollIntervalMs") / 1000)


class WindowsExcelWindows:
    def __init__(self):
        if sys.platform != "win32":
            raise WindowLayoutError("LAYOUT_WINDOWS_ONLY")
        try:
            import win32api
            import win32con
            import win32gui
            import win32process
            import pythoncom
            from win32com.client import dynamic
        except ImportError:
            raise WindowLayoutError("LAYOUT_DEPENDENCY_MISSING") from None
        self.api, self.con, self.gui, self.process = win32api, win32con, win32gui, win32process
        self.com, self.dynamic = pythoncom, dynamic
        self._review_placements = {}
        self._review_pairs = {}

    def _only_review_windows(self, collection, original: WindowInfo, copy: WindowInfo) -> bool:
        handles = {int(collection.Item(i).Hwnd) for i in range(1, int(collection.Count) + 1)}
        if not {original.key.hwnd, copy.key.hwnd}.issubset(handles):
            return False
        owned = getattr(self, '_review_pairs', {})
        for hwnd in handles - {original.key.hwnd, copy.key.hwnd}:
            matches = [(key, info) for key, (workbook_hash, info) in owned.items()
                       if key.hwnd == hwnd and workbook_hash == original.workbook_hash]
            if len(matches) != 1 or self.describe(matches[0][0]) != matches[0][1]:
                return False
        return True

    def suspend_pair(self, original: WindowInfo, copy: WindowInfo):
        if self.describe(original.key) != original or self.describe(copy.key) != copy:
            raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
        # Break only this application's verified pair, never an unrelated comparison.
        if original.key.pid == copy.key.pid:
            children = []
            self.gui.EnumChildWindows(original.key.hwnd, lambda h, _: children.append(h)
                                      if self.gui.GetClassName(h) == "EXCEL7" else None, None)
            if not children:
                raise WindowLayoutError("LAYOUT_UNAVAILABLE")
            self.com.CoInitialize()
            try:
                def leave(window):
                    collection = window.Application.Windows
                    if not self._only_review_windows(collection, original, copy):
                        raise WindowLayoutError("LAYOUT_COMPARISON_AMBIGUOUS")
                    if self._window_metadata(window, original.key) != original:
                        raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
                    collection.BreakSideBySide()
                    window.Activate()
                self._document_window(children[0], leave)
            finally:
                self.com.CoUninitialize()
        if self.describe(original.key) != original or self.describe(copy.key) != copy:
            raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
        self.gui.ShowWindow(copy.key.hwnd, self.con.SW_MINIMIZE)
        placement = getattr(self, '_review_placements', {}).get((original.key, copy.key))
        if placement:
            self.gui.SetWindowPlacement(original.key.hwnd, placement)
        else:
            self.gui.ShowWindow(original.key.hwnd, self.con.SW_MAXIMIZE)
        if not self.gui.IsIconic(copy.key.hwnd):
            raise WindowLayoutError("LAYOUT_POSITION_UNVERIFIED")

    def prepare_close_pair(self, original: WindowInfo, copy: WindowInfo):
        # Uses the same identity-checked break/placement restoration as explicit exit.
        self.suspend_pair(original, copy)

    def _is_excel(self, hwnd: int) -> WindowKey | None:
        import ctypes as ct
        from ctypes import wintypes as wt
        try:
            if not self.gui.IsWindow(hwnd) or self.gui.GetClassName(hwnd) != "XLMAIN":
                return None
            pid = self.process.GetWindowThreadProcessId(hwnd)[1]
            kernel = ct.WinDLL("kernel32", use_last_error=True)
            kernel.OpenProcess.argtypes, kernel.OpenProcess.restype = [wt.DWORD, wt.BOOL, wt.DWORD], wt.HANDLE
            kernel.CloseHandle.argtypes = [wt.HANDLE]
            kernel.QueryFullProcessImageNameW.argtypes = [wt.HANDLE, wt.DWORD, wt.LPWSTR, ct.POINTER(wt.DWORD)]
            handle = kernel.OpenProcess(0x1000, False, pid)
            if not handle:
                return None
            try:
                text, size = ct.create_unicode_buffer(32768), wt.DWORD(32768)
                if not kernel.QueryFullProcessImageNameW(handle, 0, text, ct.byref(size)):
                    return None
                return WindowKey(hwnd, pid) if ntpath.basename(text.value).lower() == "excel.exe" else None
            finally:
                kernel.CloseHandle(handle)
        except Exception:
            return None

    def request_close(self, key: WindowKey) -> bool:
        """**已停用**：面向窗口的 `WM_CLOSE`，不是面向文档的关闭。

        保留仅为让历史测试与文档里的旧结论仍可追溯。**任何调用点都不得使用它** ——
        顶层窗口句柄与副本工作簿在 Excel 真实拓扑下不是 1:1，实机上曾打到**原表**
        的保存提示（见 `docs/REVIEW_COPY_LIFECYCLE_ACCEPTANCE_2026-10-14.md` §11）。
        现行关闭原语是 `close_workbook`。
        """
        raise WindowLayoutError("LAYOUT_WINDOW_CLOSE_PRIMITIVE_DISABLED")

    @staticmethod
    def _book_review_id(book) -> str:
        # A hidden constant string, never evaluated and never a Range/Cells read.
        try:
            marker = book.Names.Item("_EB_ReviewCopyId")
            if bool(marker.Visible):
                return ""
            match = re.fullmatch(r'=\"(EB_[0-9a-f]{20})\"', str(marker.RefersTo))
            return match.group(1) if match else ""
        except Exception:
            return ""

    @staticmethod
    def _copy_book_matches(book, expected: WindowInfo) -> bool:
        # Read the retained Workbook itself, never reacquire ActiveWorkbook/Parent.
        # New copies can contain multiple snapshot sheets; legacy unmarked copies remain single-sheet.
        identity_matches = (WindowsExcelWindows._book_review_id(book) == expected.review_copy_id
                            if expected.review_copy_id else re.fullmatch(r"EB_[0-9a-f]{20}", expected.sheet_name) is not None)
        names = [str(book.Sheets.Item(i).Name) for i in range(1, int(book.Sheets.Count) + 1)]
        return (expected.unsaved and identity_matches
                and str(book.Path) == "" and (bool(expected.review_copy_id) or len(names) == 1)
                and names.count(expected.sheet_name) == 1
                and int(book.FileFormat) == 51)

    def select_pair_sheet(self, original: WindowInfo, copy: WindowInfo):
        if not copy.review_copy_id or self.describe(original.key) != original or self.describe(copy.key) != copy:
            raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
        def children_of(key):
            children = []
            self.gui.EnumChildWindows(key.hwnd, lambda h, _: children.append(h)
                                      if self.gui.GetClassName(h) == "EXCEL7" else None, None)
            if not children:
                raise WindowLayoutError("LAYOUT_UNAVAILABLE")
            return children
        child, original_child = children_of(copy.key)[0], children_of(original.key)[0]
        self.com.CoInitialize()
        try:
            book = self._document_workbook(child, copy.key)
            if book is None or not self._copy_book_matches(book, copy):
                raise WindowLayoutError("LAYOUT_COPY_CHANGED")
            sheets = [book.Worksheets.Item(i) for i in range(1, int(book.Worksheets.Count) + 1)]
            matches = [sheet for sheet in sheets if hashlib.sha256(str(sheet.Name).encode("utf-8")).hexdigest() == original.sheet_hash]
            if len(matches) != 1 or self.describe(original.key) != original or self.describe(copy.key) != copy:
                raise WindowLayoutError("LAYOUT_COPY_CHANGED")
            matches[0].Activate()
            def refocus(window):
                if self._window_metadata(window, original.key) != original:
                    raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
                window.Activate()
            self._document_window(original_child, refocus)
            selected = self.describe(copy.key)
            if not selected or selected.review_copy_id != copy.review_copy_id or selected.sheet_hash != original.sheet_hash:
                raise WindowLayoutError("LAYOUT_COPY_CHANGED")
            if hasattr(self, '_review_pairs'):
                self._review_pairs[copy.key] = (original.workbook_hash, selected)
        finally:
            self.com.CoUninitialize()

    def promote_copy_label(self, original: WindowInfo, copy: WindowInfo, nonce: str):
        """Upgrade a retained legacy copy's label, without rebuilding or changing cells."""
        if (not copy.unsaved or copy.sheet_name != nonce or original.unsaved
                or self.describe(original.key) != original or self.describe(copy.key) != copy):
            raise WindowLayoutError("LAYOUT_COPY_CHANGED")
        children = []
        self.gui.EnumChildWindows(copy.key.hwnd, lambda h, _: children.append(h)
                                  if self.gui.GetClassName(h) == "EXCEL7" else None, None)
        if not children:
            raise WindowLayoutError("LAYOUT_UNAVAILABLE")
        self.com.CoInitialize()
        try:
            book = self._document_workbook(children[0], copy.key)
            if book is None or not self._copy_book_matches(book, copy):
                raise WindowLayoutError("LAYOUT_COPY_CHANGED")
            if self.describe(original.key) != original or self.describe(copy.key) != copy:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            self._promote_book_label(book, copy, original.sheet_name, nonce)
        finally:
            self.com.CoUninitialize()

    @staticmethod
    def _promote_book_label(book, expected: WindowInfo, sheet_name: str, nonce: str):
        if expected.sheet_name != nonce or not WindowsExcelWindows._copy_book_matches(book, expected):
            raise WindowLayoutError("LAYOUT_COPY_CHANGED")
        # Add the nonce first. A rename failure retains an identifiable, retryable copy.
        if not expected.review_copy_id:
            book.Names.Add(Name="_EB_ReviewCopyId", RefersTo=f'="{nonce}"', Visible=False)
        if WindowsExcelWindows._book_review_id(book) != nonce:
            raise WindowLayoutError("LAYOUT_COPY_CHANGED")
        book.Sheets.Item(1).Name = sheet_name
        if str(book.Sheets.Item(1).Name) != sheet_name or str(book.Path):
            raise WindowLayoutError("LAYOUT_COPY_CHANGED")

    @staticmethod
    def _archive_review_book(book) -> str:
        # Server-chosen unique path only. Never overwrite or delete a user's file.
        base = application_data_root().resolve()
        root = base / "review-copies"
        root.mkdir(parents=True, exist_ok=True)
        if root.resolve() != root or root.is_symlink():
            raise WindowLayoutError("COPY_ARCHIVE_PATH_UNSAFE")
        directory = root / secrets.token_hex(16)
        directory.mkdir(exist_ok=False)
        path = directory / "review-copy.xlsx"
        book.SaveCopyAs(str(path))
        # A returned COM call is not enough: require a readable XLSX archive.
        with zipfile.ZipFile(path) as package:
            if "xl/workbook.xml" not in package.namelist() or package.testzip() is not None:
                raise WindowLayoutError("COPY_ARCHIVE_INVALID")
        return str(path)

    def close_workbook(self, key: WindowKey, expected: WindowInfo) -> CopyCloseReceipt:
        """Archive then close the SAME verified Workbook, never a window frame.

        SaveCopyAs preserves any extra edits without changing the open book's path.
        Failure to archive or recheck identity leaves the workbook open. Saved-as
        copies belong to the user and are never closed here. No global prompt flags.
        """
        recovery_path = None
        try:
            if self._is_excel(key.hwnd) != key or not expected.unsaved:
                return CopyCloseReceipt(False)
            children: list[int] = []
            self.gui.EnumChildWindows(key.hwnd, lambda h, _: children.append(h)
                                       if self.gui.GetClassName(h) == "EXCEL7" else None, None)
            if not children:
                return CopyCloseReceipt(False)
            self.com.CoInitialize()
            try:
                if any(self._document_metadata(child, key) != expected for child in children):
                    return CopyCloseReceipt(False)
                book = self._document_workbook(children[0], key)
                if book is None or not self._copy_book_matches(book, expected):
                    return CopyCloseReceipt(False)
                recovery_path = self._archive_review_book(book)
                if self._is_excel(key.hwnd) != key or not self._copy_book_matches(book, expected):
                    return CopyCloseReceipt(False, recovery_path)
                book.Close(False)
                getattr(self, '_review_pairs', {}).pop(key, None)
                placements = getattr(self, '_review_placements', {})
                for pair in list(placements):
                    if pair[1] == key:
                        del placements[pair]
                return CopyCloseReceipt(True, recovery_path)
            finally:
                self.com.CoUninitialize()
        except Exception:
            return CopyCloseReceipt(False, recovery_path)

    def _document_workbook(self, child: int, key: WindowKey):
        """取该文档窗口的 **Workbook** COM 对象（`ActiveSheet.Parent`）。

        这条链与 `_window_metadata` 读的是同一批白名单成员，所以它给出的就是
        已被复核过身份的那一本工作簿，而不是"某个活动工作簿"。
        """
        holder: list = []
        window = self._document_window(child, lambda w: holder.append(w) or w)
        if not holder:
            return None
        return holder[0].ActiveSheet.Parent

    def keys(self) -> list[WindowKey]:
        keys: list[WindowKey] = []
        def visit(hwnd, _):
            key = self._is_excel(hwnd)
            if key:
                keys.append(key)
        self.gui.EnumWindows(visit, None)
        return keys

    def describe(self, key: WindowKey) -> WindowInfo | None:
        if self._is_excel(key.hwnd) != key or not self.gui.IsWindowVisible(key.hwnd):
            return None
        children: list[int] = []
        self.gui.EnumChildWindows(key.hwnd, lambda h, _: children.append(h)
                                   if self.gui.GetClassName(h) == "EXCEL7" else None, None)
        self.com.CoInitialize()
        try:
            observations = []
            for child in children:
                info = self._document_metadata(child, key)
                if info:
                    observations.append(info)
            # Multiple panes are acceptable only if all report the same identity.
            return observations[0] if observations and all(o == observations[0] for o in observations) else None
        finally:
            self.com.CoUninitialize()

    def _document_metadata(self, child: int, key: WindowKey) -> WindowInfo | None:
        return self._document_window(child, lambda window: self._window_metadata(window, key))

    def _document_window(self, child: int, visit):
        import ctypes as ct
        from ctypes import wintypes as wt
        import uuid
        pointer = ct.c_void_p()
        oleacc = ct.OleDLL("oleacc")
        oleacc.AccessibleObjectFromWindow.argtypes = [wt.HWND, wt.DWORD, ct.c_void_p, ct.POINTER(ct.c_void_p)]
        iid = (ct.c_ubyte * 16).from_buffer_copy(uuid.UUID("00020400-0000-0000-C000-000000000046").bytes_le)
        try:
            # Microsoft documents EXCEL7 + OBJID_NATIVEOM as an Excel Window.
            oleacc.AccessibleObjectFromWindow(child, 0xFFFFFFF0, ct.byref(iid), ct.byref(pointer))
            if not pointer.value:
                return None
            dispatch = self.com.ObjectFromAddress(pointer.value, self.com.IID_IDispatch)
            window = self.dynamic.Dispatch(dispatch)
            return visit(window)
        except Exception:
            return None
        finally:
            # ObjectFromAddress QI adds its own reference. Release the original
            # AccessibleObjectFromWindow reference, including on exceptions.
            if pointer.value:
                table = ct.cast(pointer, ct.POINTER(ct.POINTER(ct.c_void_p))).contents
                ct.WINFUNCTYPE(wt.ULONG, ct.c_void_p)(table[2])(pointer)

    @staticmethod
    def _window_metadata(window, key: WindowKey) -> WindowInfo:
        # Only this whitelist of identity metadata is read. No Range/Cells.
        sheet = window.ActiveSheet
        book = sheet.Parent
        name, folder = str(sheet.Name), str(book.Path)
        full_name = str(book.FullName)
        return WindowInfo(key, path_hash(full_name) if folder else "", hashlib.sha256(name.encode("utf-8")).hexdigest(), name, not folder,
                          WindowsExcelWindows._book_review_id(book))

    def enable_synchronous_comparison(self, original: WindowKey, copy: WindowKey,
                                     original_info: WindowInfo | None = None, copy_info: WindowInfo | None = None,
                                     may_mutate=None) -> str:
        # Native Excel comparison is scoped to one Application, not two processes.
        # Never create another Application, reopen a file, hook or poll scrolling.
        if original.pid != copy.pid:
            return "different_instance"
        if not original_info or not copy_info or self.describe(original) != original_info or self.describe(copy) != copy_info:
            return "unavailable"
        children: dict[int, list[int]] = {}
        for key in (original, copy):
            children[key.hwnd] = []
            self.gui.EnumChildWindows(key.hwnd, lambda h, _, key=key: children[key.hwnd].append(h)
                                      if self.gui.GetClassName(h) == "EXCEL7" else None, None)
        if not all(children.values()):
            return "unavailable"
        self.com.CoInitialize()
        try:
            result = self._document_window(children[original.hwnd][0], lambda left:
                self._document_window(children[copy.hwnd][0], lambda right:
                    self._compare_verified_windows(left, right, original_info, copy_info, may_mutate)))
            return result if result in {"enabled", "unavailable"} else "unavailable"
        finally:
            self.com.CoUninitialize()

    def _compare_verified_windows(self, original, copy, expected_original: WindowInfo, expected_copy: WindowInfo,
                                  may_mutate=None) -> str:
        def require_available():
            if may_mutate is not None and not may_mutate():
                raise WindowLayoutError("LAYOUT_TICKET_UNAVAILABLE")
        def identities_match():
            return (self._is_excel(expected_original.key.hwnd) == expected_original.key
                    and self._is_excel(expected_copy.key.hwnd) == expected_copy.key
                    and int(original.Hwnd) == expected_original.key.hwnd
                    and int(copy.Hwnd) == expected_copy.key.hwnd
                    and self._window_metadata(original, expected_original.key) == expected_original
                    and self._window_metadata(copy, expected_copy.key) == expected_copy)
        if expected_original.key.pid != expected_copy.key.pid or expected_original.unsaved or not expected_copy.unsaved or not identities_match():
            raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
        collection = original.Application.Windows
        # Comparing is application-wide. Do not replace an unrelated comparison
        # or affect a third workbook's view. More than this pair stays manual.
        if not self._only_review_windows(collection, expected_original, expected_copy):
            raise WindowLayoutError("LAYOUT_COMPARISON_AMBIGUOUS")
        caption = str(copy.Caption)
        row, column = int(original.ScrollRow), int(original.ScrollColumn)
        if not caption or row < 1 or column < 1 or not identities_match():
            raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
        require_available()
        copy.ScrollRow, copy.ScrollColumn = row, column
        require_available()
        original.Activate()
        require_available()
        accepted = collection.CompareSideBySideWith(caption)
        if not isinstance(accepted, (bool, int)) or not accepted:
            return "unavailable"
        if not identities_match():
            raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
        require_available()
        collection.SyncScrollingSideBySide = True
        if not bool(collection.SyncScrollingSideBySide):
            return "unavailable"
        if (int(original.ScrollRow), int(original.ScrollColumn)) != (int(copy.ScrollRow), int(copy.ScrollColumn)):
            return "unavailable"
        return "enabled"

    def client_width_css_px(self, key: WindowKey) -> float:
        import ctypes as ct
        from ctypes import wintypes as wt
        user = ct.WinDLL("user32", use_last_error=True)
        user.SetThreadDpiAwarenessContext.argtypes = [ct.c_void_p]
        user.SetThreadDpiAwarenessContext.restype = ct.c_void_p
        user.GetDpiForWindow.argtypes, user.GetDpiForWindow.restype = [wt.HWND], ct.c_uint
        user.GetClientRect.argtypes, user.GetClientRect.restype = [wt.HWND, ct.POINTER(wt.RECT)], wt.BOOL
        previous = user.SetThreadDpiAwarenessContext(ct.c_void_p(-4))
        if not previous:
            raise WindowLayoutError("LAYOUT_DPI_UNAVAILABLE")
        try:
            if self._is_excel(key.hwnd) != key or not self.gui.IsWindowVisible(key.hwnd) or self.gui.IsIconic(key.hwnd):
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            rect, dpi = wt.RECT(), user.GetDpiForWindow(key.hwnd)
            if not dpi or not user.GetClientRect(key.hwnd, ct.byref(rect)) or rect.right <= rect.left:
                raise WindowLayoutError("LAYOUT_SIZE_UNAVAILABLE")
            return (rect.right - rect.left) * 96 / dpi
        finally:
            user.SetThreadDpiAwarenessContext(previous)

    def work_area_width_css_px(self, key: WindowKey) -> float:
        """The whole monitor work area (taskbar excluded) in 96-DPI logical pixels.

        This is the single shared reference for the AI pane fraction on both hosts,
        so it must NOT be a single document window's client width.
        """
        import ctypes as ct
        from ctypes import wintypes as wt
        user = ct.WinDLL("user32", use_last_error=True)
        user.SetThreadDpiAwarenessContext.argtypes = [ct.c_void_p]
        user.SetThreadDpiAwarenessContext.restype = ct.c_void_p
        user.GetDpiForWindow.argtypes, user.GetDpiForWindow.restype = [ct.c_void_p], ct.c_uint
        previous = user.SetThreadDpiAwarenessContext(ct.c_void_p(-4))
        if not previous:
            raise WindowLayoutError("LAYOUT_DPI_UNAVAILABLE")
        try:
            if self._is_excel(key.hwnd) != key:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            dpi = user.GetDpiForWindow(key.hwnd)
            if not dpi:
                raise WindowLayoutError("LAYOUT_SIZE_UNAVAILABLE")
            monitor = self.api.MonitorFromWindow(key.hwnd, self.con.MONITOR_DEFAULTTONEAREST)
            left, _, right, _ = self.api.GetMonitorInfo(monitor)["Work"]
            if right - left <= 0:
                raise WindowLayoutError("LAYOUT_SIZE_UNAVAILABLE")
            return (right - left) * 96 / dpi
        finally:
            user.SetThreadDpiAwarenessContext(previous)

    @staticmethod
    def _visible_rect(hwnd: int) -> tuple[int, int, int, int]:
        import ctypes as ct
        from ctypes import wintypes as wt
        dwm = ct.WinDLL("dwmapi")
        dwm.DwmGetWindowAttribute.argtypes = [wt.HWND, wt.DWORD, ct.c_void_p, wt.DWORD]
        dwm.DwmGetWindowAttribute.restype = ct.c_long
        rect = wt.RECT()
        if dwm.DwmGetWindowAttribute(hwnd, 9, ct.byref(rect), ct.sizeof(rect)) != 0:
            raise WindowLayoutError("LAYOUT_VISIBLE_FRAME_UNAVAILABLE")
        return rect.left, rect.top, rect.right, rect.bottom

    def position_pair(self, original: WindowKey, copy: WindowKey,
                      original_info: WindowInfo | None = None, copy_info: WindowInfo | None = None,
                      may_mutate=None) -> ExcelReviewLayoutResult:
        import ctypes as ct
        # The background service must use the original monitor's physical DPI.
        user = ct.WinDLL("user32", use_last_error=True)
        user.SetThreadDpiAwarenessContext.argtypes = [ct.c_void_p]
        user.SetThreadDpiAwarenessContext.restype = ct.c_void_p
        user.GetDpiForWindow.argtypes = [ct.c_void_p]
        user.GetDpiForWindow.restype = ct.c_uint
        previous = user.SetThreadDpiAwarenessContext(ct.c_void_p(-4))
        if not previous:
            raise WindowLayoutError("LAYOUT_DPI_UNAVAILABLE")
        placements = {}
        try:
            if self._is_excel(original.hwnd) != original or self._is_excel(copy.hwnd) != copy:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            monitor = self.api.MonitorFromWindow(original.hwnd, self.con.MONITOR_DEFAULTTONEAREST)
            x, y, right, bottom = self.api.GetMonitorInfo(monitor)["Work"]
            width, height = right - x, bottom - y
            dpi = user.GetDpiForWindow(original.hwnd)
            # Single source of truth for the outer split: the shared copy fraction.
            # The AI pane fraction is applied by the frontend against this same work area.
            fraction = 1.0 - capability_float("reviewPaneSizing", "reviewCopyWidthFraction")
            gap = int(round(capability_float("excelWindowLayout", "windowGapDip") * dpi / 96))
            if not dpi or not 0 < fraction < 1 or height <= 0 or gap < 0 or width <= gap:
                raise WindowLayoutError("LAYOUT_SIZE_UNAVAILABLE")
            original_width = int((width - gap) * fraction)
            copy_width = width - gap - original_width
            if original_width <= 0 or copy_width <= 0:
                raise WindowLayoutError("LAYOUT_SIZE_UNAVAILABLE")
            # Review copy LEFT; original with its AI task pane RIGHT, wider.
            # Targets describe VISIBLE frames, not invisible Win32 resize borders.
            targets = {original.hwnd: (right - original_width, y, right, bottom),
                       copy.hwnd: (x, y, x + copy_width, bottom)}
            for key in (original, copy):
                placements[key.hwnd] = self.gui.GetWindowPlacement(key.hwnd)
            if not hasattr(self, '_review_placements'):
                self._review_placements = {}
            # Every sheet copy shares the original window's pre-review placement.
            # A later copy must not save the already split frame as the normal size.
            baseline = next((placement for (source, _), placement in self._review_placements.items()
                             if source == original), placements[original.hwnd])
            self._review_placements.setdefault((original, copy), baseline)
            if not hasattr(self, '_review_pairs'):
                self._review_pairs = {}
            if original_info and copy_info:
                self._review_pairs[copy] = (original_info.workbook_hash, copy_info)
            synchronous_scrolling = self.enable_synchronous_comparison(original, copy, original_info, copy_info, may_mutate)
            for key in (original, copy):
                if self._is_excel(key.hwnd) != key:
                    raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
                if original_info is not None and (self.describe(original) != original_info or self.describe(copy) != copy_info):
                    raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
                if may_mutate is not None and not may_mutate():
                    raise WindowLayoutError("LAYOUT_TICKET_UNAVAILABLE")
                left, top, edge, down = targets[key.hwnd]
                self.gui.ShowWindow(key.hwnd, self.con.SW_RESTORE)
                raw, visible = self.gui.GetWindowRect(key.hwnd), self._visible_rect(key.hwnd)
                if visible[2] <= visible[0] or visible[3] <= visible[1]:
                    raise WindowLayoutError("LAYOUT_VISIBLE_FRAME_UNAVAILABLE")
                insets = tuple(v-r for v, r in zip(visible, raw))
                max_inset = capability_float("excelWindowLayout", "maxFrameInsetDip") * dpi / 96
                if any(abs(inset) > max_inset for inset in insets):
                    raise WindowLayoutError("LAYOUT_VISIBLE_FRAME_UNAVAILABLE")
                target_raw = tuple(v-inset for v, inset in zip((left, top, edge, down), insets))
                rx, ry, rr, rb = target_raw
                if may_mutate is not None and not may_mutate():
                    raise WindowLayoutError("LAYOUT_TICKET_UNAVAILABLE")
                self.gui.SetWindowPos(key.hwnd, self.con.HWND_TOP, rx, ry, rr-rx, rb-ry,
                                      self.con.SWP_NOACTIVATE | self.con.SWP_NOOWNERZORDER)
            tolerance = capability_int("excelWindowLayout", "positionTolerancePixels")
            for hwnd, target in targets.items():
                if any(abs(a-b) > tolerance for a, b in zip(self._visible_rect(hwnd), target)):
                    raise WindowLayoutError("LAYOUT_POSITION_UNVERIFIED")
            # A versioned receipt comes from verified native geometry, not the
            # HTTP handler assuming that any successful call used the new order.
            if self._is_excel(original.hwnd) != original or self._is_excel(copy.hwnd) != copy:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            if original_info is not None and (self.describe(original) != original_info or self.describe(copy) != copy_info):
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            actual_original = self._visible_rect(original.hwnd)
            actual_copy = self._visible_rect(copy.hwnd)
            if actual_copy[2] > actual_original[0] or actual_copy[0] >= actual_original[0]:
                raise WindowLayoutError("LAYOUT_POSITION_UNVERIFIED")
            return ExcelReviewLayoutResult(status="arranged", layoutVersion=3, copySide="left", originalSide="right",
                                           synchronousScrolling=synchronous_scrolling)
        except Exception:
            # Best effort restore, ONLY while original handles still belong to
            # the same processes. A partial layout must never look successful.
            for key in (original, copy):
                if key.hwnd in placements and self._is_excel(key.hwnd) == key:
                    try:
                        self.gui.SetWindowPlacement(key.hwnd, placements[key.hwnd])
                    except Exception:
                        pass
            raise WindowLayoutError("LAYOUT_POSITION_FAILED") from None
        finally:
            if previous:
                user.SetThreadDpiAwarenessContext(previous)


excel_window_layout = ExcelWindowLayout()

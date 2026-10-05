"""Bounded Windows-only layout of an original window and one NEW review copy.

No document writes, application creation, cell access, scripts, keystrokes,
subprocesses or persistence. Native Office accessibility reads identity metadata
and performs bounded window comparison/viewport operations on the verified pair.
Positioning or scrolling is never evidence of cell verification.
"""
from __future__ import annotations

import hashlib
import ntpath
import secrets
import sys
import threading
import time
from dataclasses import dataclass

from .capabilities import capability_float, capability_int
from .models import ExcelReviewLayoutPrepareRequest, ExcelReviewLayoutResult, ExcelReviewPaneMeasurement


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

    def measure(self, request: ExcelReviewLayoutPrepareRequest) -> ExcelReviewPaneMeasurement:
        # Read-only, no ticket creation/consumption, window mutation or persistence.
        # A unique saved original AND this unsaved nonce copy must still exist.
        with self._lock:
            native = self._native()
            observations = [native.describe(key) for key in native.keys()]
            originals = [info for info in observations if self._original(info, request)]
            copies = [info for info in observations if info and info.unsaved and info.sheet_name == request.copySheetName]
            if len(originals) != 1 or len(copies) != 1:
                raise WindowLayoutError("LAYOUT_MEASUREMENT_AMBIGUOUS")
            original, copy = originals[0], copies[0]
            width = native.client_width_css_px(original.key)
            if native.describe(original.key) != original or native.describe(copy.key) != copy:
                raise WindowLayoutError("LAYOUT_WINDOW_CHANGED")
            return ExcelReviewPaneMeasurement(clientWidthCssPx=width)

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
                    if info and info.unsaved and info.sheet_name == ticket.request.copySheetName:
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
                    if not info or not info.unsaved or info.sheet_name != ticket.request.copySheetName:
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
        return WindowInfo(key, path_hash(full_name) if folder else "", hashlib.sha256(name.encode("utf-8")).hexdigest(), name, not folder)

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
        if int(collection.Count) != 2 or {int(collection.Item(i).Hwnd) for i in (1, 2)} != {
                expected_original.key.hwnd, expected_copy.key.hwnd}:
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
            fraction = capability_float("excelWindowLayout", "originalWidthFraction")
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

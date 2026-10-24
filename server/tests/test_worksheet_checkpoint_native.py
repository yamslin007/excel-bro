"""Exercise native attachment, not the service-level workbook replacement."""
from types import SimpleNamespace
from unittest.mock import Mock, PropertyMock
import sys
import pytest
from server.app.worksheet_checkpoint import NativeWorksheetCheckpoint
from server.app.excel_window_layout import WindowKey, path_hash

PATH = r"C:\work\example.xlsx"

@pytest.fixture
def native(monkeypatch):
    monkeypatch.setattr("server.app.worksheet_checkpoint.time.sleep", lambda _: None)
    com = SimpleNamespace(CoInitialize=Mock(), CoUninitialize=Mock())
    monkeypatch.setitem(sys.modules, "pythoncom", com)
    client = SimpleNamespace(GetActiveObject=Mock(side_effect=AssertionError("no active workbook fallback")))
    monkeypatch.setitem(sys.modules, "win32com", SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    key = WindowKey(10, 20)
    gui = SimpleNamespace(EnumChildWindows=lambda hwnd, callback, arg: callback(11, arg), GetClassName=lambda hwnd: "EXCEL7", IsWindowVisible=lambda hwnd: True)
    adapter = SimpleNamespace(keys=lambda: [key], describe=lambda k: SimpleNamespace(unsaved=False, workbook_hash=path_hash(PATH)),
                             gui=gui, _is_excel=Mock(return_value=key), read=Mock())
    adapter._document_window = lambda child, visit: visit(SimpleNamespace(ActiveSheet=SimpleNamespace(Parent=book))) if (book := adapter.read()) is not None else None
    monkeypatch.setattr("server.app.excel_window_layout.WindowsExcelWindows", lambda: adapter)
    book = SimpleNamespace(FullName=PATH, Path=r"C:\work")
    return adapter, book, com

def attach():
    return NativeWorksheetCheckpoint().workbook(SimpleNamespace(host="excel", workbookPath=PATH))

def test_transient_native_object_failure_retries_read_and_yields_retained_book(native):
    adapter, book, com = native
    adapter.read.side_effect = [None, book]
    with attach() as actual:
        assert actual is book
    assert adapter.read.call_count == 2
    com.CoUninitialize.assert_called_once()

def test_persistent_native_object_failure_is_unavailable_not_identity_change(native):
    adapter, _, com = native
    adapter.read.return_value = None
    with pytest.raises(ValueError, match="对象暂不可读取"):
        with attach():
            pytest.fail("must not execute without a workbook object")
    assert adapter.read.call_count == 3
    com.CoUninitialize.assert_called_once()

def test_actual_identity_mismatch_is_not_retried(native):
    adapter, book, _ = native
    changing = type("ChangingBook", (), {"Path": r"C:\work", "FullName": PropertyMock(side_effect=[PATH, r"C:\other.xlsx"])})()
    adapter.read.side_effect = [changing, book]
    with pytest.raises(ValueError, match="身份发生变化"):
        with attach():
            pytest.fail("wrong workbook was accepted")
    assert adapter.read.call_count == 1

def test_reused_window_between_read_attempts_cannot_be_adopted(native):
    adapter, book, _ = native
    adapter._is_excel.side_effect = [WindowKey(10, 20), WindowKey(10, 99)]
    adapter.read.side_effect = [None, book]
    with pytest.raises(ValueError, match="窗口身份发生变化"):
        with attach():
            pytest.fail("reused window was accepted")
    assert adapter.read.call_count == 1

def test_failure_after_yield_does_not_replay_operation(native):
    adapter, book, _ = native
    adapter.read.return_value = book
    calls = []
    with pytest.raises(RuntimeError, match="operation failed"):
        with attach():
            calls.append("write")
            raise RuntimeError("operation failed")
    assert calls == ["write"]
    assert adapter.read.call_count == 1

def test_identity_metadata_and_execution_use_one_native_object(native):
    adapter, book, _ = native
    adapter.describe = Mock(side_effect=AssertionError("must not reacquire after describe"))
    adapter.read.side_effect = [book, None]
    with attach() as actual:
        assert actual is book
    adapter.describe.assert_not_called()
    assert adapter.read.call_count == 1

def test_duplicate_matching_windows_remain_ambiguous(native):
    adapter, book, _ = native
    adapter.keys = lambda: [WindowKey(10, 20), WindowKey(12, 20)]
    adapter._is_excel.side_effect = lambda hwnd: WindowKey(hwnd, 20)
    adapter.read.return_value = book
    with pytest.raises(ValueError, match="无法唯一定位"):
        with attach():
            pytest.fail("ambiguous workbook accepted")

def test_retry_budget_covers_two_transient_failures_without_replaying_writes(native):
    adapter, book, _ = native
    adapter.read.side_effect = [None, None, book]
    with attach() as actual:
        assert actual is book
    assert adapter.read.call_count == 3

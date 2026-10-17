"""Real adapter control flow; only the COM workbook and OS window provider are fake."""
from types import SimpleNamespace
import zipfile

import pytest
from fastapi.testclient import TestClient

from server.app import main
from server.app.excel_window_layout import WindowsExcelWindows, WindowInfo, WindowKey


KEY = WindowKey(20, 100)
NAME = 'EB_' + 'a' * 20
EXPECTED = WindowInfo(KEY, '', '', NAME, True)


class Book:
    def __init__(self, *, path='', name=NAME):
        self.Path = path
        self.FileFormat = 51
        self.Sheets = SimpleNamespace(Count=1, Item=lambda _: SimpleNamespace(Name=name))
        self.calls = []
        self.extra_edit = 'USER_EDIT_OUTSIDE_PLAN'
        self.fail_archive = False
        self.after_archive = lambda: None

    def SaveCopyAs(self, path):
        self.calls.append(('archive', path))
        if self.fail_archive:
            raise OSError('disk full')
        with zipfile.ZipFile(path, 'w') as out:
            out.writestr('xl/workbook.xml', '<workbook/>')
            out.writestr('xl/worksheets/sheet1.xml', self.extra_edit)
        self.after_archive()

    def Close(self, save):
        self.calls.append(('close', save))


def adapter(monkeypatch, tmp_path, book):
    monkeypatch.setenv('EXCEL_BRO_CONFIG_DIR', str(tmp_path))
    native = WindowsExcelWindows.__new__(WindowsExcelWindows)
    native.gui = SimpleNamespace(EnumChildWindows=lambda hwnd, visit, ctx: visit(21, ctx),
                                 GetClassName=lambda _: 'EXCEL7')
    native.com = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    native._is_excel = lambda _: KEY
    native._document_metadata = lambda *args: EXPECTED
    native._document_workbook = lambda *args: book
    return native


def test_retained_workbook_is_archived_with_extra_edits_then_closed_without_original_access(monkeypatch, tmp_path):
    book, original = Book(), Book(path='E:/original.xlsx')
    native = adapter(monkeypatch, tmp_path, book)
    acquired = []
    def acquire(*args):
        acquired.append(True)
        return book if len(acquired) == 1 else original
    native._document_workbook = acquire
    result = native.close_workbook(KEY, EXPECTED)
    assert result.accepted and result.recovery_path
    assert [call[0] for call in book.calls] == ['archive', 'close']
    assert book.calls[-1] == ('close', False)
    assert len(acquired) == 1 and original.calls == []
    with zipfile.ZipFile(result.recovery_path) as archive:
        assert archive.read('xl/worksheets/sheet1.xml') == b'USER_EDIT_OUTSIDE_PLAN'
    assert book.Path == ''  # SaveCopyAs never turns the review book into the archive.


@pytest.mark.parametrize('failure', ['saved_original', 'wrong_nonce', 'multiple_sheets', 'unsupported_format',
                                     'disk_full', 'empty_archive', 'identity_changes_after_archive'])
def test_failure_never_closes_the_book(monkeypatch, tmp_path, failure):
    book = Book(path='E:/original.xlsx' if failure == 'saved_original' else '',
                name='Other' if failure == 'wrong_nonce' else NAME)
    native = adapter(monkeypatch, tmp_path, book)
    if failure == 'multiple_sheets': book.Sheets.Count = 2
    if failure == 'unsupported_format': book.FileFormat = 56
    if failure == 'disk_full': book.fail_archive = True
    if failure == 'empty_archive': book.SaveCopyAs = lambda path: None
    if failure == 'identity_changes_after_archive':
        book.after_archive = lambda: setattr(book, 'Path', 'E:/saved-by-user.xlsx')
    result = native.close_workbook(KEY, EXPECTED)
    assert not result.accepted
    assert not any(call[0] == 'close' for call in book.calls)
    if failure == 'identity_changes_after_archive':
        assert result.recovery_path is not None


def test_close_error_keeps_recovery_file(monkeypatch, tmp_path):
    book = Book()
    def fail_close(save): raise RuntimeError('COM busy')
    book.Close = fail_close
    result = adapter(monkeypatch, tmp_path, book).close_workbook(KEY, EXPECTED)
    assert not result.accepted and zipfile.is_zipfile(result.recovery_path)


def test_v3_request_is_rejected_before_invoking_bridge(monkeypatch):
    calls = []
    monkeypatch.setattr(main.excel_window_layout, 'close_copy', lambda payload: calls.append(payload))
    client = TestClient(main.app, headers={'origin': 'https://localhost:3000'})
    result = client.post('/api/office/review-layout/close', json={
        'workbookPathHash': 'a'*64, 'worksheetNameHash': 'b'*64, 'copySheetName': NAME})
    assert result.status_code == 422 and calls == []

import hashlib
from types import SimpleNamespace

import pytest

from server.app import main
from server.app.excel_window_layout import ExcelWindowLayout, WindowInfo, WindowLayoutError, WindowsExcelWindows
from server.tests.test_excel_window_layout import FakeWindows, SOURCE, COPY, REQUEST, NONCE, client


def prepared_pair():
    native = FakeWindows()
    native.windows[COPY] = WindowInfo(COPY, '', REQUEST.worksheetNameHash, 'Ledger', True, NONCE)
    return native, ExcelWindowLayout(native)


def test_switch_tabs_keeps_same_pair_without_arrange_suspend_or_close(monkeypatch):
    native, service = prepared_pair()
    monkeypatch.setattr(main, 'excel_window_layout', service)
    for name in ['结果3', '结果1', '结果2']:
        digest = hashlib.sha256(name.encode()).hexdigest()
        native.windows[SOURCE] = WindowInfo(SOURCE, REQUEST.workbookPathHash, digest, name, False)
        request = REQUEST.model_copy(update={'copyIdentityVersion': 2, 'worksheetNameHash': digest})
        response = client().post('/api/office/review-layout/select-sheet', json=request.model_dump())
        assert response.status_code == 200, response.text
        assert response.json() == {'status': 'selected'}
        assert native.windows[COPY].sheet_name == name
        assert set(native.windows) == {SOURCE, COPY}
    assert native.moves == []
    assert native.close_posted == []


@pytest.mark.parametrize('problem', ['saved', 'marker', 'source', 'v1'])
def test_switch_never_guesses_a_workbook(problem):
    native, service = prepared_pair()
    request = REQUEST.model_copy(update={'copyIdentityVersion': 2})
    if problem == 'saved':
        native.windows[COPY] = WindowInfo(COPY, '', REQUEST.worksheetNameHash, 'Ledger', False, NONCE)
    if problem == 'marker':
        native.windows[COPY] = WindowInfo(COPY, '', REQUEST.worksheetNameHash, 'Ledger', True, '')
    if problem == 'source':
        native.windows.pop(SOURCE)
    if problem == 'v1':
        request = REQUEST
    with pytest.raises(WindowLayoutError):
        service.select_copy_sheet(request)
    assert native.moves == []
    assert native.close_posted == []


def test_multisheet_workbook_close_requires_owned_hidden_identity_and_archivable_xlsx():
    sheets = [SimpleNamespace(Name=name) for name in ['结果1', '结果2', '结果3']]
    book = SimpleNamespace(Path='', FileFormat=51, Sheets=SimpleNamespace(Count=3, Item=lambda i: sheets[i-1]),
                           Names=SimpleNamespace(Item=lambda name: SimpleNamespace(Visible=False, RefersTo=f'="{NONCE}"')))
    expected = WindowInfo(COPY, '', '', '结果3', True, NONCE)
    assert WindowsExcelWindows._copy_book_matches(book, expected)
    book.Path = 'C:/saved'
    assert not WindowsExcelWindows._copy_book_matches(book, expected)
    book.Path = ''
    book.Names = SimpleNamespace(Item=lambda name: SimpleNamespace(Visible=False, RefersTo='="unowned"'))
    assert not WindowsExcelWindows._copy_book_matches(book, expected)


def test_select_endpoint_rejects_untrusted_origin_and_extra_payload(monkeypatch):
    native, service = prepared_pair()
    monkeypatch.setattr(main, 'excel_window_layout', service)
    payload = REQUEST.model_dump() | {'copyIdentityVersion': 2}
    assert client(origin='https://untrusted.invalid').post('/api/office/review-layout/select-sheet', json=payload).status_code == 403
    assert client().post('/api/office/review-layout/select-sheet', json=payload | {'values': [['secret']]}).status_code == 422
    assert native.moves == []


def test_native_selector_activates_only_the_matching_sheet_and_refocuses_original():
    native, _ = prepared_pair()
    original = WindowInfo(SOURCE, REQUEST.workbookPathHash, hashlib.sha256('结果3'.encode()).hexdigest(), '结果3', False)
    native.windows[SOURCE] = original
    copy = native.windows[COPY]
    activated = []
    def sheet(name):
        def activate():
            activated.append(name)
            native.windows[COPY] = WindowInfo(COPY, '', hashlib.sha256(name.encode()).hexdigest(), name, True, NONCE)
        return SimpleNamespace(Name=name, Activate=activate)
    sheets = [sheet(name) for name in ['Ledger', '结果2', '结果3']]
    collection = SimpleNamespace(Count=3, Item=lambda i: sheets[i-1])
    book = SimpleNamespace(Path='', FileFormat=51, Sheets=collection, Worksheets=collection,
                           Names=SimpleNamespace(Item=lambda _: SimpleNamespace(Visible=False, RefersTo=f'="{NONCE}"')))
    bridge = object.__new__(WindowsExcelWindows)
    bridge.describe = native.describe
    bridge.gui = SimpleNamespace(EnumChildWindows=lambda hwnd, callback, arg: callback(hwnd+100, arg), GetClassName=lambda _: 'EXCEL7')
    bridge.com = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    bridge._document_workbook = lambda child, key: book
    bridge._document_window = lambda child, callback: callback(SimpleNamespace(Activate=lambda: activated.append('original')))
    bridge._window_metadata = lambda window, key: original
    bridge._review_pairs = {}
    bridge.select_pair_sheet(original, copy)
    assert activated == ['结果3', 'original']
    assert native.windows[SOURCE] == original
    assert bridge._review_pairs[COPY] == (original.workbook_hash, native.windows[COPY])

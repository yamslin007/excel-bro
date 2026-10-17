from __future__ import annotations

import hashlib
from types import SimpleNamespace

import pytest
from fastapi.testclient import TestClient

from server.app.excel_window_layout import CopyCloseReceipt, ExcelWindowLayout, WindowInfo, WindowKey, WindowLayoutError, WindowsExcelWindows, path_hash
from server.app.models import ExcelReviewCopyCloseRequest, ExcelReviewLayoutPrepareRequest, ExcelReviewLayoutResult
import server.app.main as main


ORIGIN = 'https://localhost:3000'
SOURCE, COPY, OTHER = WindowKey(10, 100), WindowKey(20, 200), WindowKey(30, 300)
NONCE = 'EB_' + 'a' * 20
HASH = hashlib.sha256(b'Ledger').hexdigest()
def closed_handles(backend):
    """Only the **handles** we asked to close.

    Identity is asserted separately (see the wrong-identity test); keeping this helper
    handle-only means the existing "只关了副本" assertions stay readable.
    """
    return [key for key, _identity in backend.close_posted]


REQUEST = ExcelReviewLayoutPrepareRequest(workbookPathHash=path_hash('E:/Original.xlsx'), worksheetNameHash=HASH, copySheetName=NONCE)


class FakeWindows:
    def __init__(self):
        self.windows = {SOURCE: WindowInfo(SOURCE, REQUEST.workbookPathHash, HASH, 'Ledger', False)}
        self.moves = []
        self.inspect = None
        self.error = None
        # Close-path controls. `close_posted` records every requested handle so a
        # test can prove that only the verified copy was ever asked to close.
        self.close_posted = []
        self.close_accepted = True
        self.window_stays_open = False
        # Set by a test to simulate "the workbook behind that handle is no longer ours".
        self.close_identity_override = None
        # A slow close: the window only disappears after N witness reads. Used to prove the
        # close loop keeps polling instead of declaring a result on the first look.
        self.remove_after_polls = None
        self.pending_close = None

    def keys(self):
        return list(self.windows)

    def describe(self, key):
        if self.inspect:
            self.inspect(key)
        pending = self.pending_close
        if pending and key == pending[0]:
            remaining = pending[1] - 1
            if remaining <= 0:
                self.pending_close = None
                self.windows.pop(key, None)
            else:
                self.pending_close = (key, remaining)
        return self.windows.get(key)

    def close_workbook(self, key, expected):
        # Workbook-level close: the handle **and** the registered workbook identity must
        # both still match. Recording `(key, expected)` lets a test prove the original's
        # identity was never the thing we asked to close.
        self.close_posted.append((key, expected))
        if self.close_identity_override is not None:
            return CopyCloseReceipt(False)
        if self.describe(key) != expected:
            return CopyCloseReceipt(False)
        if not self.close_accepted:
            return CopyCloseReceipt(False)
        if self.window_stays_open:
            return CopyCloseReceipt(True, "E:/recovery/review-copy.xlsx")
        if self.remove_after_polls:
            self.pending_close = (key, self.remove_after_polls)
            return CopyCloseReceipt(True, "E:/recovery/review-copy.xlsx")
        self.windows.pop(key, None)
        return CopyCloseReceipt(True, "E:/recovery/review-copy.xlsx")

    def request_close(self, key):
        # The window-message primitive is disabled in the real backend; a fake that still
        # offers it must not silently become the path under test.
        raise AssertionError('close path must not use the window-message primitive')

    def position_pair(self, original, copy, original_info=None, copy_info=None, may_mutate=None):
        if self.error:
            raise self.error
        self.moves.append((original, copy))
        return ExcelReviewLayoutResult(status='arranged', layoutVersion=3, copySide='left', originalSide='right', synchronousScrolling='different_instance')

    def add_copy(self, key=COPY, *, name=NONCE, saved=False):
        self.windows[key] = WindowInfo(key, '', '', name, not saved)

    def client_width_css_px(self, key):
        assert key == SOURCE
        return 1200

    def work_area_width_css_px(self, key):
        assert key == SOURCE
        return 1600


def setup():
    now = [0.0]
    backend = FakeWindows()
    service = ExcelWindowLayout(backend, clock=lambda: now[0], sleep=lambda seconds: now.__setitem__(0, now[0] + seconds))
    return service, backend, now


def test_measure_only_verified_live_pair_no_mutations_or_tickets():
    service, backend, _ = setup()
    backend.add_copy()
    assert service.measure(REQUEST).model_dump() == {'clientWidthCssPx':1200.0, 'workAreaWidthCssPx':1600.0}
    assert not backend.moves and not service._tickets


@pytest.mark.parametrize('failure', ['no_copy','duplicate_source','duplicate_copy','saved_copy','source_changed_during_measure'])
def test_measure_rejects_unknown_or_changed_pair(monkeypatch, failure):
    service, backend, _ = setup()
    if failure != 'no_copy':
        backend.add_copy(saved=failure=='saved_copy')
    if failure == 'duplicate_source':
        backend.windows[OTHER] = WindowInfo(OTHER, REQUEST.workbookPathHash, HASH, 'Ledger', False)
    if failure == 'duplicate_copy':
        backend.add_copy(OTHER)
    if failure == 'source_changed_during_measure':
        def measure(key):
            backend.windows[SOURCE] = WindowInfo(SOURCE, path_hash('E:/Other.xlsx'),HASH,'Ledger',False)
            return 1200
        monkeypatch.setattr(backend,'client_width_css_px',measure)
    with pytest.raises(WindowLayoutError):
        service.measure(REQUEST)
    assert not backend.moves and not service._tickets


@pytest.mark.parametrize('dpi', [96,144,192])
def test_native_client_measure_is_dpi_normalized_and_read_only(monkeypatch,dpi):
    import ctypes as ct
    from ctypes import wintypes as wt
    from unittest.mock import Mock
    backend=WindowsExcelWindows.__new__(WindowsExcelWindows)
    backend._is_excel=lambda hwnd:SOURCE
    backend.gui=SimpleNamespace(IsWindowVisible=lambda hwnd:True,IsIconic=lambda hwnd:False)
    def client_rect(hwnd,result):
        assert hwnd == SOURCE.hwnd
        rect=ct.cast(result,ct.POINTER(wt.RECT)).contents
        rect.left,rect.right=0,int(1200*dpi/96)
        return 1
    context=Mock(return_value=123)
    user=SimpleNamespace(SetThreadDpiAwarenessContext=context,GetDpiForWindow=Mock(return_value=dpi),GetClientRect=Mock(side_effect=client_rect))
    monkeypatch.setattr(ct,'WinDLL',lambda *a,**k:user,raising=False)
    assert backend.client_width_css_px(SOURCE)==1200
    assert context.call_count==2


@pytest.mark.parametrize('dpi,work,expected', [
    (96, (-1000, 50, 0, 850), 1000.0),      # secondary monitor left of the primary
    (192, (0, 0, 3840, 2000), 1920.0),      # 200% scaling
    (144, (100, 100, 1700, 900), 1066.6666666666667),
])
def test_native_work_area_is_dpi_normalized_and_offset_tolerant(monkeypatch, dpi, work, expected):
    import ctypes as ct
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    backend._is_excel = lambda hwnd: SOURCE
    backend.api = SimpleNamespace(MonitorFromWindow=Mock(return_value=1), GetMonitorInfo=Mock(return_value={'Work': work}))
    backend.con = SimpleNamespace(MONITOR_DEFAULTTONEAREST=2)
    context = Mock(return_value=123)
    user = SimpleNamespace(SetThreadDpiAwarenessContext=context, GetDpiForWindow=Mock(return_value=dpi))
    monkeypatch.setattr(ct, 'WinDLL', lambda *a, **k: user, raising=False)
    assert backend.work_area_width_css_px(SOURCE) == pytest.approx(expected)
    assert context.call_count == 2  # restores the worker thread's previous DPI context


def test_native_work_area_fails_closed_without_identity_or_size(monkeypatch):
    import ctypes as ct
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    backend._is_excel = lambda hwnd: None
    backend.api = SimpleNamespace(MonitorFromWindow=Mock(return_value=1), GetMonitorInfo=Mock(return_value={'Work': (0, 0, 0, 0)}))
    backend.con = SimpleNamespace(MONITOR_DEFAULTTONEAREST=2)
    user = SimpleNamespace(SetThreadDpiAwarenessContext=Mock(return_value=123), GetDpiForWindow=Mock(return_value=96))
    monkeypatch.setattr(ct, 'WinDLL', lambda *a, **k: user, raising=False)
    with pytest.raises(WindowLayoutError) as error:
        backend.work_area_width_css_px(SOURCE)
    assert error.value.code == 'LAYOUT_WINDOW_CHANGED'


def test_measure_api_projects_only_widths_and_rejects_sensitive_payload(monkeypatch):
    from server.app.models import ExcelReviewPaneMeasurement
    monkeypatch.setattr(main.excel_window_layout,'measure',lambda request:ExcelReviewPaneMeasurement(clientWidthCssPx=1200, workAreaWidthCssPx=1600))
    response=client().post('/api/office/review-layout/measure',json=REQUEST.model_dump())
    assert response.status_code==200 and response.json()=={'clientWidthCssPx':1200.0, 'workAreaWidthCssPx':1600.0}
    for extra in ('plan','values','fullPath','hwnd'):
        rejected=client().post('/api/office/review-layout/measure',json={**REQUEST.model_dump(),extra:'LOCAL_PRIVATE_SENTINEL'})
        assert rejected.status_code==422


def test_measure_api_rejects_non_excel_origin():
    response=client(origin='http://localhost:3889').post('/api/office/review-layout/measure',json=REQUEST.model_dump())
    assert response.status_code==403


def test_arranges_only_original_and_new_copy_in_other_process_and_consumes_ticket():
    service, backend, _ = setup()
    token = service.prepare(REQUEST, ORIGIN)
    backend.add_copy()
    result = service.arrange(token, ORIGIN)
    assert result.model_dump() == {'status':'arranged', 'layoutVersion':3, 'copySide':'left', 'originalSide':'right', 'synchronousScrolling':'different_instance'}
    assert backend.moves == [(SOURCE, COPY)]
    with pytest.raises(WindowLayoutError, match='无法安全'):
        service.arrange(token, ORIGIN)
    assert backend.moves == [(SOURCE, COPY)]
    assert not service._tickets


def test_generic_backend_success_is_not_verified_geometry(monkeypatch):
    service, backend, _ = setup()
    token = service.prepare(REQUEST, ORIGIN)
    backend.add_copy()
    monkeypatch.setattr(backend, 'position_pair', lambda *args: None)
    with pytest.raises(WindowLayoutError) as failure:
        service.arrange(token, ORIGIN)
    assert failure.value.code == 'LAYOUT_POSITION_UNVERIFIED'


@pytest.mark.parametrize('reason', ['same_name_other_path', 'duplicate_original'])
def test_prepare_rejects_wrong_full_path_or_ambiguous_original(reason):
    service, backend, _ = setup()
    if reason == 'same_name_other_path':
        backend.windows[SOURCE] = WindowInfo(SOURCE, path_hash('E:/Another/Original.xlsx'), HASH, 'Ledger', False)
    else:
        backend.windows[OTHER] = WindowInfo(OTHER, REQUEST.workbookPathHash, HASH, 'Ledger', False)
    with pytest.raises(WindowLayoutError):
        service.prepare(REQUEST, ORIGIN)
    assert not backend.moves and not service._tickets


@pytest.mark.parametrize('reason', ['preexisting_copy', 'wrong_nonce', 'saved_copy', 'ambiguous_new', 'source_changed', 'source_missing', 'copy_changed', 'slow_inspection'])
def test_unconfirmed_window_evidence_never_moves(reason):
    service, backend, now = setup()
    if reason == 'preexisting_copy':
        backend.add_copy()
    token = service.prepare(REQUEST, ORIGIN)
    backend.add_copy(name='Unrelated' if reason == 'wrong_nonce' else NONCE, saved=reason == 'saved_copy')
    if reason == 'ambiguous_new':
        backend.add_copy(OTHER)
    if reason == 'source_changed':
        backend.windows[SOURCE] = WindowInfo(SOURCE, path_hash('E:/Another.xlsx'), HASH, 'Ledger', False)
    if reason == 'source_missing':
        del backend.windows[SOURCE]
    calls = [0]
    def inspect(key):
        if key == COPY:
            calls[0] += 1
            if reason == 'copy_changed' and calls[0] == 2:
                backend.windows[COPY] = WindowInfo(COPY, '', '', 'Changed', True)
            if reason == 'slow_inspection':
                now[0] = 1000
    backend.inspect = inspect
    with pytest.raises(WindowLayoutError):
        service.arrange(token, ORIGIN)
    assert backend.moves == []
    assert not service._tickets


@pytest.mark.parametrize('reason', ['other_origin', 'expired', 'unknown'])
def test_ticket_boundaries(reason):
    service, backend, now = setup()
    token = service.prepare(REQUEST, ORIGIN)
    backend.add_copy()
    if reason == 'expired':
        now[0] = 121
    with pytest.raises(WindowLayoutError):
        service.arrange('missing' if reason == 'unknown' else token, 'http://localhost:3000' if reason == 'other_origin' else ORIGIN)
    assert not backend.moves
    assert token in service._tickets  # rejected at entry, not consumed by a later deadline guard


def test_failed_move_cannot_replay_and_ticket_pool_is_bounded():
    service, backend, now = setup()
    token = service.prepare(REQUEST, ORIGIN)
    backend.add_copy()
    backend.error = WindowLayoutError('LAYOUT_POSITION_FAILED')
    with pytest.raises(WindowLayoutError):
        service.arrange(token, ORIGIN)
    assert token not in service._tickets
    for _ in range(20):
        service.prepare(REQUEST, ORIGIN)
    with pytest.raises(WindowLayoutError) as error:
        service.prepare(REQUEST, ORIGIN)
    assert error.value.code == 'LAYOUT_BUSY'
    now[0] = 121
    service.prepare(REQUEST, ORIGIN)
    assert len(service._tickets) == 1


def test_native_metadata_only_reads_identity_whitelist_not_cells_or_values():
    class ReadOnlyMetadata:
        def __init__(self, **attrs):
            object.__setattr__(self, '_attrs', attrs)
        def __getattr__(self, name):
            assert name in self._attrs, f'FORBIDDEN_READ:{name}'
            return self._attrs[name]
        def __setattr__(self, name, value):
            raise AssertionError('FORBIDDEN_WRITE')
    book = ReadOnlyMetadata(Path='E:/', FullName='E:/Original.xlsx')
    sheet = ReadOnlyMetadata(Name='Ledger', Parent=book)
    info = WindowsExcelWindows._window_metadata(ReadOnlyMetadata(ActiveSheet=sheet), SOURCE)
    assert info == WindowInfo(SOURCE, REQUEST.workbookPathHash, HASH, 'Ledger', False)


@pytest.mark.parametrize('failed_metadata', [False, True])
def test_native_accessibility_adapter_uses_64_bit_pointer_and_releases_original_reference(monkeypatch, failed_metadata):
    import ctypes as ct
    from ctypes import wintypes as wt
    import uuid
    from unittest.mock import Mock
    monkeypatch.setattr(ct,'WINFUNCTYPE',getattr(ct,'WINFUNCTYPE',ct.CFUNCTYPE),raising=False)
    releases = []
    release = ct.WINFUNCTYPE(wt.ULONG,ct.c_void_p)(lambda address: releases.append(address) or 0)
    table = (ct.c_void_p * 3)(0,0,ct.cast(release,ct.c_void_p).value)
    dispatch_object = (ct.c_void_p * 1)(ct.addressof(table))
    pointer = ct.addressof(dispatch_object)
    def accessible(child, object_id, iid, result):
        assert child == 77 and object_id == 0xFFFFFFF0
        assert bytes(ct.cast(iid,ct.POINTER(ct.c_ubyte * 16)).contents) == uuid.UUID('00020400-0000-0000-C000-000000000046').bytes_le
        ct.cast(result,ct.POINTER(ct.c_void_p))[0] = pointer
        return 0
    monkeypatch.setattr(ct,'OleDLL',lambda *a:SimpleNamespace(AccessibleObjectFromWindow=accessible),raising=False)
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    backend.com = SimpleNamespace(ObjectFromAddress=Mock(return_value='owned dispatch'),IID_IDispatch='IID')
    book = SimpleNamespace(Path='E:/',FullName='E:/Original.xlsx')
    window = SimpleNamespace(ActiveSheet=None if failed_metadata else SimpleNamespace(Name='Ledger',Parent=book))
    backend.dynamic = SimpleNamespace(Dispatch=Mock(return_value=window))
    info = backend._document_metadata(77,SOURCE)
    assert (info is None) == failed_metadata
    backend.com.ObjectFromAddress.assert_called_once_with(pointer,'IID')
    assert releases == [pointer]


@pytest.mark.parametrize('failure', [None, 'move', 'geometry', 'pid_reused'])
@pytest.mark.parametrize('window_dpi', [96, 192])
@pytest.mark.parametrize('invalid_size', [False, True])
def test_native_positioning_exact_pair_geometry_and_rollback(monkeypatch, failure, window_dpi, invalid_size):
    import ctypes
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    keys = {SOURCE.hwnd:SOURCE, COPY.hwnd:COPY}
    backend._is_excel = lambda hwnd: keys.get(hwnd)
    if failure == 'pid_reused':
        keys[COPY.hwnd] = WindowKey(COPY.hwnd,999)
    placements = {10:'original placement',20:'copy placement'}
    rects, moved, restored = {10:(0,0,800,800),20:(0,0,800,800)}, [], []
    def move(hwnd, _top, x, y, width, height, _flags):
        moved.append(hwnd)
        if failure == 'move' and hwnd == 20:
            raise RuntimeError('native failure')
        rects[hwnd] = (x,y,x+width,y+height)
    backend.api = SimpleNamespace(MonitorFromWindow=lambda *a:1,GetMonitorInfo=lambda *a:{'Work':(-1000,50,-999,850) if invalid_size else (-1000,50,0,850)})
    backend.con = SimpleNamespace(MONITOR_DEFAULTTONEAREST=2,SW_RESTORE=9,HWND_TOP=0,SWP_NOACTIVATE=16,SWP_NOOWNERZORDER=512)
    backend.gui = SimpleNamespace(GetWindowPlacement=lambda hwnd:placements[hwnd],ShowWindow=Mock(),SetWindowPos=move,
                                  GetWindowRect=lambda hwnd:(0,0,1,1) if failure == 'geometry' else rects[hwnd],
                                  SetWindowPlacement=lambda hwnd,p:restored.append((hwnd,p)))
    backend._visible_rect = lambda hwnd: (0,0,1,1) if failure == 'geometry' else rects[hwnd]
    backend.enable_synchronous_comparison = lambda *a: 'different_instance'
    dpi = Mock(return_value=123)
    monkeypatch.setattr(ctypes,'WinDLL',lambda *a,**k:SimpleNamespace(SetThreadDpiAwarenessContext=dpi,GetDpiForWindow=Mock(return_value=window_dpi)),raising=False)
    if failure or invalid_size:
        with pytest.raises(WindowLayoutError):
            backend.position_pair(SOURCE,COPY)
        if failure == 'pid_reused' or invalid_size:
            assert not moved and not restored
        else:
            assert restored == [(10,placements[10]),(20,placements[20])]
    else:
        result = backend.position_pair(SOURCE,COPY)
        assert result.model_dump() == {'status':'arranged', 'layoutVersion':3, 'copySide':'left', 'originalSide':'right', 'synchronousScrolling':'different_instance'}
        assert moved == [10,20] and not restored
        # Work area (-1000,50)..(0,850): width 1000. Copy 35% = 350, original+AI 65% = 650.
        expected = {10:(-650,50,0,850),20:(-1000,50,-650,850)}
        assert rects == expected  # original + AI right; copy left; no overlap/off-screen
        assert rects[10][0] - rects[20][2] == 0
        assert rects[10][2] - rects[10][0] > rects[20][2] - rects[20][0]
        assert rects[10][2] - rects[10][0] == 650 and rects[20][2] - rects[20][0] == 350
    assert dpi.call_count == 2  # restore the worker thread's previous DPI context


def test_unavailable_dpi_context_never_mutates_windows(monkeypatch):
    import ctypes
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    backend.gui = SimpleNamespace(SetWindowPos=Mock(),ShowWindow=Mock())
    monkeypatch.setattr(ctypes,'WinDLL',lambda *a,**k:SimpleNamespace(SetThreadDpiAwarenessContext=Mock(return_value=None),GetDpiForWindow=Mock(return_value=96)),raising=False)
    with pytest.raises(WindowLayoutError) as error:
        backend.position_pair(SOURCE,COPY)
    assert error.value.code == 'LAYOUT_DPI_UNAVAILABLE'
    backend.gui.SetWindowPos.assert_not_called()
    backend.gui.ShowWindow.assert_not_called()


@pytest.mark.parametrize('variant', ['same_panes','different_panes','invisible'])
def test_native_describe_requires_consistent_window_metadata_and_com_cleanup(variant):
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    backend._is_excel = lambda hwnd: SOURCE
    def enumerate_children(hwnd, callback, data):
        assert hwnd == SOURCE.hwnd
        for child in (1,2,3):
            callback(child,data)
    backend.gui = SimpleNamespace(IsWindowVisible=lambda hwnd:variant != 'invisible',EnumChildWindows=enumerate_children,
                                  GetClassName=lambda hwnd:'EXCEL7' if hwnd in (1,2) else 'OtherChild')
    backend.com = SimpleNamespace(CoInitialize=Mock(),CoUninitialize=Mock())
    original = WindowInfo(SOURCE,REQUEST.workbookPathHash,HASH,'Ledger',False)
    backend._document_metadata = lambda child,key: WindowInfo(SOURCE,'other',HASH,'Ledger',False) if variant == 'different_panes' and child == 2 else original
    info = backend.describe(SOURCE)
    assert (info == original) == (variant == 'same_panes')
    assert backend.com.CoInitialize.call_count == backend.com.CoUninitialize.call_count == (0 if variant == 'invisible' else 1)


@pytest.mark.parametrize('path,expected', [
    ('E:\\Case Folder\\temp\\..\\Original.xlsx','e:/case folder/original.xlsx'),
    ('\\\\server\\share\\Original.xlsx','//server/share/original.xlsx'),
])
def test_path_hash_matches_frontend_cases(path, expected):
    assert path_hash(path) == hashlib.sha256(expected.encode()).hexdigest()


def client(*, origin=ORIGIN, client_host='127.0.0.1', base='http://127.0.0.1:8765'):
    headers = {'Origin':origin} if origin else {}
    return TestClient(main.app, base_url=base, client=(client_host,5000), headers=headers)


def test_real_http_body_token_only_response_and_no_persistence(monkeypatch):
    service, backend, _ = setup()
    monkeypatch.setattr(main, 'excel_window_layout', service)
    # Real API routes, not a mocked return value. Ledger access is forbidden here.
    monkeypatch.setattr(main, 'FolderOperationLedger', lambda *a, **kw: pytest.fail('layout must not persist'))
    local_plan = {'actions':[{'type':'writeValues','values':[['LOCAL_PRIVATE_SENTINEL']]}]}
    assert 'LOCAL_PRIVATE_SENTINEL' in str(local_plan)
    response = client().post('/api/office/review-layout/prepare', json=REQUEST.model_dump())
    assert response.status_code == 200
    ticket = response.json()
    assert set(ticket) == {'token', 'layoutVersion'} and ticket['layoutVersion'] == 3 and len(ticket['token']) == 43
    record = service._tickets[ticket['token']]
    assert record.request.model_dump() == REQUEST.model_dump()
    assert 'LOCAL_PRIVATE_SENTINEL' not in str(record) and 'E:/Original.xlsx' not in str(record)
    backend.add_copy()
    arranged = client().post('/api/office/review-layout/arrange', json={'token':ticket['token']})
    assert arranged.status_code == 200 and arranged.json() == {'status':'arranged', 'layoutVersion':3, 'copySide':'left', 'originalSide':'right', 'synchronousScrolling':'different_instance'}
    assert backend.moves == [(SOURCE,COPY)] and not service._tickets


@pytest.mark.parametrize('extra', ['plan','values','fullPath','hwnd','script'])
def test_api_rejects_extra_payload_before_any_native_call(monkeypatch, extra):
    service, backend, _ = setup()
    monkeypatch.setattr(main, 'excel_window_layout', service)
    response = client().post('/api/office/review-layout/prepare', json={**REQUEST.model_dump(), extra:'LOCAL_PRIVATE_SENTINEL'})
    assert response.status_code == 422 and not service._tickets and not backend.moves


@pytest.mark.parametrize('kwargs', [
    {'origin':None}, {'origin':'https://evil.example'}, {'origin':'http://127.0.0.1:3889'},
    {'client_host':'192.168.1.10'}, {'base':'http://evil.example'},
])
def test_api_denies_non_excel_or_non_loopback_callers(monkeypatch, kwargs):
    service, backend, _ = setup()
    monkeypatch.setattr(main, 'excel_window_layout', service)
    response = client(**kwargs).post('/api/office/review-layout/prepare', json=REQUEST.model_dump())
    assert response.status_code == 403 and not service._tickets and not backend.moves


def test_api_failure_never_echoes_native_exception_metadata(monkeypatch):
    monkeypatch.setattr(main, 'excel_window_layout', SimpleNamespace(prepare=lambda *a: (_ for _ in ()).throw(RuntimeError('SECRET_PATH_AND_CELL_VALUE'))))
    response = client().post('/api/office/review-layout/prepare', json=REQUEST.model_dump())
    assert response.status_code == 409
    assert 'SECRET_PATH_AND_CELL_VALUE' not in response.text
    assert response.json()['detail']['code'] == 'LAYOUT_UNAVAILABLE'


@pytest.mark.parametrize('variant', ['enabled', 'rejected', 'ignored', 'third_window', 'wrong_hwnd', 'changed_copy', 'expired'])
def test_com_comparison_only_verified_pair_and_explicit_sync_readback(variant):
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    copy_key = WindowKey(COPY.hwnd, SOURCE.pid)
    backend._is_excel = lambda hwnd: {SOURCE.hwnd: SOURCE, COPY.hwnd: copy_key}.get(hwnd)
    book = SimpleNamespace(Path='E:/', FullName='E:/Original.xlsx')
    original = SimpleNamespace(Hwnd=SOURCE.hwnd, ActiveSheet=SimpleNamespace(Name='Ledger', Parent=book),
                               Activate=Mock(), ScrollRow=37, ScrollColumn=4)
    copy = SimpleNamespace(Hwnd=999 if variant == 'wrong_hwnd' else COPY.hwnd, Caption='Book1',
                           ActiveSheet=SimpleNamespace(Name='wrong' if variant == 'changed_copy' else NONCE,
                                                       Parent=SimpleNamespace(Path='', FullName='Book1')),
                           ScrollRow=1, ScrollColumn=1)
    class Collection:
        Count = 3 if variant == 'third_window' else 2
        def __init__(self):
            self.calls = []
            self.sync = False
        def Item(self, index):
            if variant == 'wrong_hwnd' and index == 2:
                # Independent data source: the collection identifies the correct
                # window, but the accessibility wrapper points elsewhere.
                return SimpleNamespace(Hwnd=COPY.hwnd)
            return [original, copy, SimpleNamespace(Hwnd=30)][index - 1]
        def CompareSideBySideWith(self, caption):
            self.calls.append(caption)
            return variant != 'rejected'
        @property
        def SyncScrollingSideBySide(self):
            return self.sync
        @SyncScrollingSideBySide.setter
        def SyncScrollingSideBySide(self, value):
            if variant != 'ignored': self.sync = value
    windows = Collection()
    original.Application = SimpleNamespace(Windows=windows)
    expected_original = WindowInfo(SOURCE, REQUEST.workbookPathHash, HASH, 'Ledger', False)
    expected_copy = WindowInfo(copy_key, '', hashlib.sha256(NONCE.encode()).hexdigest(), NONCE, True)
    if variant in {'third_window', 'wrong_hwnd', 'changed_copy', 'expired'}:
        with pytest.raises(WindowLayoutError):
            backend._compare_verified_windows(original, copy, expected_original, expected_copy, lambda: variant != 'expired')
        assert not windows.calls and not original.Activate.called
        assert copy.ScrollRow == copy.ScrollColumn == 1
    else:
        result = backend._compare_verified_windows(original, copy, expected_original, expected_copy)
        assert result == ('enabled' if variant == 'enabled' else 'unavailable')
        assert windows.calls == ['Book1']
        if variant == 'enabled':
            assert windows.SyncScrollingSideBySide is True
            assert (copy.ScrollRow, copy.ScrollColumn) == (37, 4)


def test_cross_instance_sync_is_explicitly_unavailable_without_com_access():
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    assert backend.enable_synchronous_comparison(SOURCE, COPY) == 'different_instance'


@pytest.mark.parametrize('variant', ['enabled', 'failed', 'changed'])
def test_sync_accessibility_adapter_receives_original_ticket_identity_and_cleans_up(variant):
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    copy_key = WindowKey(COPY.hwnd, SOURCE.pid)
    original_info = WindowInfo(SOURCE, REQUEST.workbookPathHash, HASH, 'Ledger', False)
    copy_info = WindowInfo(copy_key, '', HASH, NONCE, True)
    backend.describe = lambda key: original_info if key == SOURCE else (
        WindowInfo(copy_key, '', HASH, 'changed', True) if variant == 'changed' else copy_info)
    backend.com = SimpleNamespace(CoInitialize=Mock(), CoUninitialize=Mock())
    backend.gui = SimpleNamespace(EnumChildWindows=lambda hwnd, visit, data: visit(hwnd+1, data),GetClassName=lambda hwnd:'EXCEL7')
    backend._document_window = lambda child, visit: None if variant == 'failed' else visit('window:'+str(child))
    backend._compare_verified_windows = Mock(return_value='enabled')
    result = backend.enable_synchronous_comparison(SOURCE, copy_key, original_info, copy_info)
    assert result == ('enabled' if variant == 'enabled' else 'unavailable')
    assert backend.com.CoInitialize.call_count == backend.com.CoUninitialize.call_count == (0 if variant == 'changed' else 1)
    if variant == 'enabled':
        backend._compare_verified_windows.assert_called_once_with('window:11', 'window:21', original_info, copy_info, None)
    else: backend._compare_verified_windows.assert_not_called()


@pytest.mark.parametrize('window_dpi', [96, 192])
def test_visible_frame_positioning_removes_invisible_resize_border_gap(monkeypatch, window_dpi):
    import ctypes
    from unittest.mock import Mock
    backend = WindowsExcelWindows.__new__(WindowsExcelWindows)
    keys = {SOURCE.hwnd: SOURCE, COPY.hwnd: COPY}
    backend._is_excel = keys.get
    rects = {10:(100,100,900,900),20:(100,100,900,900)}
    inset = 8 * window_dpi // 96
    backend._visible_rect = lambda hwnd: (rects[hwnd][0]+inset, rects[hwnd][1], rects[hwnd][2]-inset, rects[hwnd][3]-inset)
    backend.enable_synchronous_comparison = lambda *a: 'different_instance'
    backend.api = SimpleNamespace(MonitorFromWindow=lambda *a:1, GetMonitorInfo=lambda *a:{'Work':(0,0,2000,1000)})
    backend.con = SimpleNamespace(MONITOR_DEFAULTTONEAREST=2,SW_RESTORE=9,HWND_TOP=0,SWP_NOACTIVATE=16,SWP_NOOWNERZORDER=512)
    def move(hwnd, z, x, y, w, h, flags): rects[hwnd] = (x,y,x+w,y+h)
    backend.gui = SimpleNamespace(GetWindowPlacement=lambda h:'placement',ShowWindow=Mock(),SetWindowPos=move,
                                  GetWindowRect=rects.get,SetWindowPlacement=Mock())
    monkeypatch.setattr(ctypes,'WinDLL',lambda *a,**k:SimpleNamespace(SetThreadDpiAwarenessContext=Mock(return_value=123),GetDpiForWindow=Mock(return_value=window_dpi)),raising=False)
    result = backend.position_pair(SOURCE,COPY)
    # Work area 2000 wide → copy 35% = 700, original + AI 65% = 1300.
    assert backend._visible_rect(20) == (0,0,700,1000)
    assert backend._visible_rect(10) == (700,0,2000,1000)
    assert result.synchronousScrolling == 'different_instance'


def test_dwm_adapter_reads_physical_visible_frame_and_rejects_failed_read(monkeypatch):
    import ctypes as ct
    from ctypes import wintypes as wt
    def read(hwnd, attribute, pointer, size):
        assert hwnd == 10 and attribute == 9 and size == ct.sizeof(wt.RECT)
        rect = ct.cast(pointer, ct.POINTER(wt.RECT)).contents
        rect.left, rect.top, rect.right, rect.bottom = -1000, 50, 0, 850
        return 0
    monkeypatch.setattr(ct, 'WinDLL', lambda *a,**k: SimpleNamespace(DwmGetWindowAttribute=read), raising=False)
    assert WindowsExcelWindows._visible_rect(10) == (-1000,50,0,850)
    def fail(*a): return -1
    monkeypatch.setattr(ct, 'WinDLL', lambda *a,**k: SimpleNamespace(DwmGetWindowAttribute=fail), raising=False)
    with pytest.raises(WindowLayoutError): WindowsExcelWindows._visible_rect(10)


# --- 审阅副本关闭：只动身份已核实的那一份，且结果不冒充原表写回结果 ---


def test_close_targets_only_the_verified_unsaved_copy_and_is_idempotent():
    service, backend, _ = setup()
    backend.add_copy()
    first = service.close_copy(REQUEST)
    assert first.model_dump() == {'status':'closed', 'layoutVersion':4, 'copyClosed':True, 'recoveryCopyPath': "E:/recovery/review-copy.xlsx"}
    assert closed_handles(backend) == [COPY] and not backend.moves
    # Repeat cleanup is idempotent: the copy is gone, and that is a SUCCESS.
    again = service.close_copy(REQUEST)
    assert again.model_dump() == {'status':'already_closed', 'layoutVersion':4, 'copyClosed':True, 'recoveryCopyPath': None}
    assert closed_handles(backend) == [COPY]


def test_close_asks_for_the_copy_identity_never_the_original_identity():
    """关闭请求必须带着**副本**身份发出，且绝不带原表身份。

    这是本轮缺陷的正面锚点：旧实现投递的是一条面向窗口的消息，只重核
    `(class, pid)`，于是"打中哪本工作簿"完全没有证据。现在关闭原语是
    `close_workbook(handle, identity)`，身份由调用方在投递前复核后随请求一起给出 ——
    原表的身份**永远不会**作为参数出现在关闭调用里。
    """
    service, backend, _ = setup()
    backend.add_copy()
    service.close_copy(REQUEST)
    assert [identity for _key, identity in backend.close_posted] == [backend.windows.get(COPY) or backend.close_posted[0][1]]
    # 逐字说清：没有一次关闭是以原表身份发起的。
    original_identity = WindowInfo(SOURCE, REQUEST.workbookPathHash, HASH, 'Ledger', False)
    assert all(identity != original_identity for _key, identity in backend.close_posted)
    assert all(key != SOURCE for key, _identity in backend.close_posted)


def test_close_refuses_when_the_workbook_identity_no_longer_matches():
    """身份不符时**不关任何东西**，并如实报失败（而不是"关成功了"）。"""
    service, backend, _ = setup()
    backend.add_copy()
    # 复核期间副本被换成了别的未保存工作簿：句柄还是那个句柄，身份已经不是了。
    backend.close_identity_override = object()
    result = service.close_copy(REQUEST)
    assert result.copyClosed is False
    assert result.status in {'identity_lost', 'failed'}
    # 原表仍在场：这一条路径绝不能因为副本身份丢失就退而关原表。
    assert SOURCE in backend.windows


def test_close_falls_back_to_manual_when_no_workbook_primitive_exists():
    """平台没有工作簿级原语时降级为手动关闭，绝不回退到窗口消息。"""
    # 用一个**没有** `close_workbook` 的替身：实例上绑定方法删不掉（del 只删实例属性，
    # 类方法仍在），所以这里显式遮蔽成 None —— `hasattr` 见到 None 仍是 True，
    # 因此用一个真正的子类才符合"平台不提供该原语"这件事。
    class NoWorkbookPrimitive(FakeWindows):
        close_workbook = None

        def __getattribute__(self, name):
            if name == 'close_workbook':
                raise AttributeError(name)
            return super().__getattribute__(name)

    now = [0.0]
    backend = NoWorkbookPrimitive()
    backend.add_copy()
    service = ExcelWindowLayout(backend, clock=lambda: now[0],
                                sleep=lambda seconds: now.__setitem__(0, now[0] + seconds))
    result = service.close_copy(REQUEST)
    assert result.model_dump() == {'status': 'unsupported', 'layoutVersion': 4, 'copyClosed': False, 'recoveryCopyPath': None}
    assert backend.close_posted == []


def test_close_never_acts_when_the_review_context_cannot_be_confirmed():
    # A copy of ours is open, but the original is gone or duplicated. Posting WM_CLOSE
    # to what may be the last visible workbook could close the whole Excel instance,
    # so the only safe answer is "I do not know".
    for variant in ('missing_original', 'duplicate_original'):
        service, backend, _ = setup()
        backend.add_copy()
        if variant == 'missing_original':
            del backend.windows[SOURCE]
        else:
            backend.windows[OTHER] = WindowInfo(OTHER, REQUEST.workbookPathHash, HASH, 'Ledger', False)
        assert service.close_copy(REQUEST).status == 'identity_lost'
        assert backend.close_posted == [] and not service._tickets


def test_close_refuses_ambiguous_or_user_saved_copies_without_touching_them():
    # Two windows claim this session's nonce: guessing could close an unrelated workbook.
    service, backend, _ = setup()
    backend.add_copy()
    backend.add_copy(OTHER)
    assert service.close_copy(REQUEST).status == 'ambiguous'
    assert backend.close_posted == []
    # The user saved/replaced our copy: that file is theirs, never auto-discarded.
    service, backend, _ = setup()
    backend.add_copy(saved=True)
    assert service.close_copy(REQUEST).status == 'user_saved'
    assert backend.close_posted == []


def test_close_reports_unconfirmed_close_instead_of_forcing_or_killing():
    service, backend, now = setup()
    backend.add_copy()
    backend.window_stays_open = True   # the window never goes away while we poll
    result = service.close_copy(REQUEST)
    assert result.model_dump() == {'status':'close_unconfirmed', 'layoutVersion':4, 'copyClosed':False, 'recoveryCopyPath': "E:/recovery/review-copy.xlsx"}
    # The receipt has four fields: this bridge has NO dialog evidence, so it must
    # not be able to claim "the host is asking to save" in the first place.
    assert set(result.model_dump()) == {'status', 'layoutVersion', 'copyClosed', 'recoveryCopyPath'}
    assert 'save' not in str(result.model_dump()).lower()
    # Exactly one request: we do not hammer the host on a retry loop inside one call,
    # and we never answer its prompt.
    assert closed_handles(backend) == [COPY]
    assert now[0] >= 3
    # A rejected WM_CLOSE is a close failure, not a fake success.
    backend.close_posted.clear()
    backend.close_accepted = False
    assert service.close_copy(REQUEST).model_dump() == {'status':'failed', 'layoutVersion':4, 'copyClosed':False, 'recoveryCopyPath': None}


def test_close_keeps_polling_a_slow_close_and_only_then_reports_closed():
    # A close that takes a few witness reads is a SUCCESS: the deadline must not be
    # mistaken for "still on screen, so something is blocking it".
    service, backend, now = setup()
    backend.add_copy()
    backend.remove_after_polls = 3
    result = service.close_copy(REQUEST)
    assert result.model_dump() == {'status':'closed', 'layoutVersion':4, 'copyClosed':True, 'recoveryCopyPath': "E:/recovery/review-copy.xlsx"}
    assert closed_handles(backend) == [COPY]
    # It really started before the deadline and finished before it too.
    assert 0 < now[0] < 3


def test_close_cannot_tell_a_save_prompt_from_any_other_blocking_modal():
    # A non-save modal blocks the close, and a save prompt does too. Without dialog
    # evidence both MUST collapse to the same honest answer: "close unconfirmed".
    service, backend, _ = setup()
    backend.add_copy()
    backend.window_stays_open = True
    blocked_by_modal = service.close_copy(REQUEST)
    service, backend, _ = setup()
    backend.add_copy()
    backend.window_stays_open = True
    possibly_save_prompt = service.close_copy(REQUEST)
    assert blocked_by_modal == possibly_save_prompt
    assert blocked_by_modal.status == 'close_unconfirmed'
    # Neither is allowed to masquerade as a completed close or as a business failure.
    assert blocked_by_modal.copyClosed is False
    assert blocked_by_modal.status not in ('closed', 'already_closed', 'failed')


def test_close_degrades_to_unsupported_without_a_verified_native_primitive():
    from server.app.excel_window_layout import ExcelWindowLayout as Layout
    backend = SimpleNamespace(keys=lambda: [SOURCE], describe=lambda key: WindowInfo(key, '', '', NONCE, True))
    assert Layout(backend).close_copy(REQUEST).status == 'unsupported'


def test_close_reverifies_identity_immediately_before_posting():
    # The copy is swapped for a different workbook between enumeration and the action.
    service, backend, _ = setup()
    backend.add_copy()
    seen = []

    def inspect(key):
        if key == COPY:
            seen.append(key)
            if len(seen) == 2:   # the re-verification read, just before WM_CLOSE
                backend.windows[COPY] = WindowInfo(COPY, '', '', 'Unrelated', True)

    backend.inspect = inspect
    assert service.close_copy(REQUEST).status == 'identity_lost'
    assert backend.close_posted == []


def test_close_request_shape_forbids_extra_fields_and_requires_this_session_nonce():
    assert ExcelReviewCopyCloseRequest(**{**REQUEST.model_dump(), "closeVersion": 4}).copySheetName == NONCE
    for bad in ({'copySheetName': 'Workbook5'}, {'copySheetName': 'EB_' + 'G' * 20}, {'copySheetName': 'EB_' + 'a' * 19}):
        with pytest.raises(Exception):
            ExcelReviewCopyCloseRequest(**{**{**REQUEST.model_dump(), "closeVersion": 4}, **bad})
    with pytest.raises(Exception):
        ExcelReviewCopyCloseRequest(**{**{**REQUEST.model_dump(), "closeVersion": 4}, 'hwnd': 12345})


def test_same_named_workbook_or_switched_active_window_never_mis_closes():
    # An unrelated workbook that also has a "Ledger" sheet, in another process, plus a
    # different active window, must not change which handle receives WM_CLOSE.
    service, backend, _ = setup()
    backend.windows[OTHER] = WindowInfo(OTHER, path_hash('E:/SameName.xlsx'), HASH, 'Ledger', False)
    backend.add_copy()
    assert service.close_copy(REQUEST).status == 'closed'
    assert closed_handles(backend) == [COPY]


def test_close_api_projects_only_close_fields_and_rejects_sensitive_payload(monkeypatch):
    service, backend, _ = setup()
    monkeypatch.setattr(main, 'excel_window_layout', service)
    backend.add_copy()
    response = client().post('/api/office/review-layout/close', json={**REQUEST.model_dump(), "closeVersion": 4})
    assert response.status_code == 200
    body = response.json()
    # Nothing about the original's write state may travel in this response.
    assert set(body) == {'status', 'layoutVersion', 'copyClosed', 'recoveryCopyPath'} and body == {'status':'closed', 'layoutVersion':4, 'copyClosed':True, 'recoveryCopyPath': "E:/recovery/review-copy.xlsx"}
    assert 'written' not in response.text and 'E:/Original.xlsx' not in response.text
    for extra in ('plan', 'values', 'fullPath', 'hwnd', 'script'):
        rejected = client().post('/api/office/review-layout/close', json={**{**REQUEST.model_dump(), "closeVersion": 4}, extra:'LOCAL_PRIVATE_SENTINEL'})
        assert rejected.status_code == 422
    assert client(origin='http://localhost:3889').post('/api/office/review-layout/close', json={**REQUEST.model_dump(), "closeVersion": 4}).status_code == 403


def test_close_api_transport_failure_degrades_to_failed_never_to_a_write_result(monkeypatch):
    monkeypatch.setattr(main, 'excel_window_layout', SimpleNamespace(
        close_copy=lambda *a: (_ for _ in ()).throw(RuntimeError('SECRET_PATH_AND_CELL_VALUE'))))
    response = client().post('/api/office/review-layout/close', json={**REQUEST.model_dump(), "closeVersion": 4})
    assert response.status_code == 200
    assert response.json() == {'status':'failed', 'layoutVersion':4, 'copyClosed':False, 'recoveryCopyPath': None}
    assert 'SECRET_PATH_AND_CELL_VALUE' not in response.text

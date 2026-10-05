"""Independent in-memory mutations; never operate native windows or edit source."""
import inspect
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from _pytest.outcomes import Failed
from server.app import excel_window_layout as production
from server.tests import test_excel_window_layout as cases
import pytest

source = inspect.getsource(production.ExcelWindowLayout)
checks = [
    ('verified geometry receipt', 'if not isinstance(result, ExcelReviewLayoutResult):', 'if False:',
     lambda: cases.test_generic_backend_success_is_not_verified_geometry(pytest.MonkeyPatch())),
    ('full path identity', 'info.workbook_hash == request.workbookPathHash', 'True',
     lambda: cases.test_prepare_rejects_wrong_full_path_or_ambiguous_original('same_name_other_path')),
    ('new window only', 'set(native.keys()) - ticket.baseline', 'set(native.keys()) - {ticket.original}',
     lambda: cases.test_unconfirmed_window_evidence_never_moves('preexisting_copy')),
    ('origin-bound ticket', 'ticket.origin != origin', 'False',
     lambda: cases.test_ticket_boundaries('other_origin')),
    ('expired ticket', 'ticket.expires <= self.clock()', 'False',
     lambda: cases.test_ticket_boundaries('expired')),
    ('copy final identity', 'if not info or not info.unsaved or info.sheet_name != ticket.request.copySheetName:', 'if False:',
     lambda: cases.test_unconfirmed_window_evidence_never_moves('copy_changed')),
    ('no late window movement', 'if self.clock() >= deadline:\n                        raise WindowLayoutError("LAYOUT_TICKET_UNAVAILABLE")', 'if False:\n                        raise WindowLayoutError("LAYOUT_TICKET_UNAVAILABLE")',
     lambda: cases.test_unconfirmed_window_evidence_never_moves('slow_inspection')),
]
failed = []
for name, before, after, check in checks:
    assert source.count(before) == 1, name
    namespace = dict(vars(production))
    exec(compile(source.replace(before, after), '<isolated-window-layout-mutation>', 'exec'), namespace)
    original = cases.ExcelWindowLayout
    cases.ExcelWindowLayout = namespace['ExcelWindowLayout']
    try:
        check()
    except (AssertionError, Failed):
        print(f'CAUGHT: {name}')
    else:
        failed.append(name)
        print(f'UNDETECTED: {name}')
    finally:
        cases.ExcelWindowLayout = original
if failed:
    raise SystemExit(1)
print(f'{len(checks)} independent guards caught; production source unchanged.')

# Geometry regressions use the real positioning method with fake Win32 calls.
# No window enumeration, COM metadata or actual SetWindowPos is performed.
position_source = inspect.getsource(production.WindowsExcelWindows)
position_checks = [
    ('reversed copy/AI ordering',
     'targets = {original.hwnd: (right - original_width, y, right, bottom),\n                       copy.hwnd: (x, y, x + copy_width, bottom)}',
     'targets = {original.hwnd: (x, y, x + original_width, bottom),\n                       copy.hwnd: (x + original_width + gap, y, right, bottom)}', 96),
    ('invisible frame compensation removed',
     'target_raw = tuple(v-inset for v, inset in zip((left, top, edge, down), insets))',
     'target_raw = (left, top, edge, down)', 192),
]
for name, before, after, dpi in position_checks:
    assert position_source.count(before) == 1, name
    namespace = dict(vars(production))
    exec(compile(position_source.replace(before,after), '<isolated-layout-geometry-mutation>', 'exec'), namespace)
    original = cases.WindowsExcelWindows
    cases.WindowsExcelWindows = namespace['WindowsExcelWindows']
    try:
        with pytest.MonkeyPatch.context() as monkeypatch:
            if name == 'invisible frame compensation removed':
                cases.test_visible_frame_positioning_removes_invisible_resize_border_gap(monkeypatch, dpi)
            else:
                cases.test_native_positioning_exact_pair_geometry_and_rollback(monkeypatch, None, dpi, False)
    except (AssertionError, Failed, production.WindowLayoutError):
        print(f'CAUGHT: {name}')
    else:
        failed.append(name)
        print(f'UNDETECTED: {name}')
    finally:
        cases.WindowsExcelWindows = original
if failed:
    raise SystemExit(1)
print('2 geometry regression mutations caught; native windows untouched.')

sync_checks = [
    ('comparison must be accepted', 'if not isinstance(accepted, (bool, int)) or not accepted:', 'if False:', 'rejected'),
    ('synchronous property must read back', 'if not bool(collection.SyncScrollingSideBySide):', 'if False:', 'ignored'),
    ('no unrelated window comparison', 'if int(collection.Count) != 2 or', 'if False and', 'third_window'),
    ('exact COM window HWND identity', 'and int(copy.Hwnd) == expected_copy.key.hwnd', 'and True', 'wrong_hwnd'),
]
for name, before, after, variant in sync_checks:
    assert position_source.count(before) == 1, name
    namespace = dict(vars(production))
    exec(compile(position_source.replace(before,after), '<isolated-scroll-guard-mutation>', 'exec'), namespace)
    original = cases.WindowsExcelWindows
    cases.WindowsExcelWindows = namespace['WindowsExcelWindows']
    try:
        cases.test_com_comparison_only_verified_pair_and_explicit_sync_readback(variant)
    except (AssertionError, Failed):
        print(f'CAUGHT: {name}')
    else:
        failed.append(name)
        print(f'UNDETECTED: {name}')
    finally:
        cases.WindowsExcelWindows = original
if failed: raise SystemExit(1)
print('4 independent native scrolling guards caught; real workbooks untouched.')

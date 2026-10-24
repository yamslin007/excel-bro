from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
import json
import pytest
from openpyxl import Workbook, load_workbook
from server.app.models import WorksheetDeleteCheckpointRequest, WorksheetCheckpointBackupConsent
from server.app.worksheet_checkpoint import WorksheetCheckpointService

class Collection:
    def __init__(self, items): self.items = items
    @property
    def Count(self): return len(self.items)
    def Item(self, index): return self.items[index - 1]

class Sheet:
    Visible = -1
    def __init__(self, name, book): self.Name, self.Parent = name, book
    def Delete(self):
        assert (Path(self.Parent.archive).parent / "state.json").is_file()
        state = json.loads((Path(self.Parent.archive).parent / "state.json").read_text(encoding="utf-8"))
        assert state["result"]["status"] == "executing"
        assert state["sha256"]
        self.Parent.calls.append(self.Name)
        self.Parent.items.remove(self)
        if self.Parent.fail_delete: raise RuntimeError("post-mutation exception")

class Book:
    FullName = r"C:\work\example.xlsx"
    Path = r"C:\work"
    ReadOnly = False
    ProtectStructure = False
    FileFormat = 51
    def __init__(self):
        self.items = [Sheet(name, self) for name in ["Sheet1", "Sheet2", "Sheet3"]]
        self.Worksheets = self.Sheets = Collection(self.items)
        self.Application = SimpleNamespace(DisplayAlerts=True, EnableEvents=True)
        self.calls = []; self.fail_delete = False; self.after_backup = lambda: None
    def SaveCopyAs(self, path):
        wb = Workbook(); wb.remove(wb.active)
        for sheet in self.items:
            ws = wb.create_sheet(sheet.Name); ws["A1"] = "original"; ws["B1"] = "=1+1"
            ws.sheet_state = "veryHidden" if sheet.Visible == 2 else "hidden" if sheet.Visible == 0 else "visible"
        wb.save(path); self.archive = path; self.after_backup()

class Native:
    def __init__(self, book): self.book = book; self.opened = []
    @contextmanager
    def workbook(self, request): yield self.book
    def open_copy(self, request, path): self.opened.append(path)

def request(**updates):
    payload = dict(operationId="a"*32, host="wps-et", hostWindowId=11,
        workbookPath=Book.FullName, authorizedSheets=["Sheet1", "Sheet2", "Sheet3"],
        plan={"id":"delete-plan", "title":"retain first", "summary":"delete others", "actions":[
            {"type":"deleteWorksheet", "sheet":"Sheet2"}, {"type":"deleteWorksheet", "sheet":"Sheet3"}]})
    payload.update(updates)
    return WorksheetDeleteCheckpointRequest.model_validate(payload)

@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path))
    native = Native(Book()); service = WorksheetCheckpointService(native)
    return service, native

@pytest.mark.parametrize("host", ["excel", "wps-et"])
def test_delete_readback_and_open_full_before_copy_without_touching_original(env, host):
    service, native = env
    result = service.execute(request(host=host), "https://pane")
    assert result.status == "completed"
    assert result.deletedSheets == ["Sheet2", "Sheet3"]
    assert result.remainingSheets == ["Sheet1"]
    assert native.book.Application.DisplayAlerts and native.book.Application.EnableEvents
    native.book.items[0].extra_edit = "later user edit"
    opened = service.open(request().operationId, "https://pane")
    assert opened["opened"]
    restored = load_workbook(native.opened[0])
    assert restored.sheetnames == ["Sheet1", "Sheet2", "Sheet3"]
    assert restored["Sheet2"]["B1"].value == "=1+1"
    assert native.book.items[0].extra_edit == "later user edit"
    assert [sheet.Name for sheet in native.book.items] == ["Sheet1"]
    assert native.opened[0].name.startswith("recovered-")


def test_repeated_request_does_not_delete_twice(env):
    service, native = env
    first = service.execute(request(), "o")
    assert service.execute(request(), "o") == first
    assert native.book.calls == ["Sheet2", "Sheet3"]


def test_partial_native_mutation_keeps_checkpoint_and_restores_app_flags(env):
    service, native = env; native.book.fail_delete = True
    result = service.execute(request(), "o")
    assert result.status == "partial_or_unknown" and result.recoveryAvailable
    assert native.book.calls == ["Sheet2"]
    assert native.book.Application.DisplayAlerts and native.book.Application.EnableEvents
    service.open(request().operationId, "o")
    assert load_workbook(native.opened[0]).sheetnames == ["Sheet1", "Sheet2", "Sheet3"]


def test_scope_gap_does_not_even_archive_unselected_sheet(env):
    service, native = env
    result = service.execute(request(authorizedSheets=["Sheet2", "Sheet3"]), "o")
    assert result.status == "needs_backup_consent" and not result.recoveryAvailable
    assert result.backupSheets == ["Sheet1"]
    assert not hasattr(native.book, "archive") and not native.book.calls


def test_backup_failure_zero_delete(env):
    service, native = env
    native.book.SaveCopyAs = lambda path: Path(path).write_text("broken zip")
    result = service.execute(request(), "o")
    assert result.status == "failed" and not result.recoveryAvailable
    assert not native.book.calls


def test_identity_changed_during_backup_zero_delete(env):
    service, native = env
    native.book.after_backup = lambda: setattr(native.book, "FullName", r"C:\other.xlsx")
    result = service.execute(request(), "o")
    assert result.status == "failed" and result.recoveryAvailable
    assert not native.book.calls


def test_archive_sheet_mismatch_zero_delete(env):
    service, native = env
    def wrong_archive(path):
        Workbook().save(path)
        native.book.archive = path
    native.book.SaveCopyAs = wrong_archive
    result = service.execute(request(), "o")
    assert result.status == "failed" and not native.book.calls


def test_keep_last_visible_sheet(env):
    service, native = env; native.book.items[0].Visible = 0
    result = service.execute(request(), "o")
    assert result.status == "failed" and not native.book.calls


def test_corrupted_archive_cannot_be_opened(env):
    service, native = env; service.execute(request(), "o")
    Path(native.book.archive).write_bytes(b"broken")
    with pytest.raises(Exception): service.open(request().operationId, "o")
    assert not native.opened


def test_other_origin_and_rebound_operation_rejected(env):
    service, native = env; service.execute(request(), "o")
    with pytest.raises(ValueError): service.open(request().operationId, "another")
    with pytest.raises(ValueError): service.execute(request(host="excel"), "o")
    assert len(native.book.calls) == 2


def test_crash_receipt_is_unknown_not_retryable(env):
    service, native = env; service.execute(request(), "o")
    state_path = Path(native.book.archive).parent / "state.json"
    state = json.loads(state_path.read_text(encoding="utf-8")); state["result"]["status"] = "executing"
    state_path.write_text(json.dumps(state), encoding="utf-8")
    assert service.status(request().operationId, "o").status == "partial_or_unknown"
    assert service.execute(request(), "o").status == "partial_or_unknown"
    assert len(native.book.calls) == 2


def test_mixed_plan_and_unselected_targets_rejected(env):
    service, native = env
    invalid = request().model_dump(); invalid["plan"]["actions"] = [{"type":"setFont", "sheet":"Sheet1", "range":"A1", "bold":True}]
    with pytest.raises(ValueError): service.execute(WorksheetDeleteCheckpointRequest.model_validate(invalid), "o")
    assert not native.book.calls


def test_restoration_copy_is_byte_identical_and_original_archive_immutable(env):
    service, native = env; service.execute(request(), "o")
    archive = Path(native.book.archive); original = archive.read_bytes()
    service.open(request().operationId, "o")
    assert native.opened[0] != archive, "recovery must open a derivative, never the immutable checkpoint"
    assert native.opened[0].read_bytes() == original
    native.opened[0].write_bytes(b"user edit in recovered file")
    service.open(request().operationId, "o")
    assert native.opened[1] != native.opened[0]
    assert native.opened[1].read_bytes() == original == archive.read_bytes()


def test_disabled_checkpoint_cannot_delete(env, monkeypatch):
    from server.app import worksheet_checkpoint as module
    service, native = env
    monkeypatch.setattr(module, "capabilities", lambda: {"worksheetCheckpoint": {"allowedActions": []}})
    with pytest.raises(ValueError, match="未开放"): service.execute(request(), "o")
    assert not native.book.calls


def test_api_requires_local_explicit_origin_and_rejects_wrong_schema(env):
    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from server.app.worksheet_checkpoint_api import create_worksheet_checkpoint_router
    app = FastAPI(); app.include_router(create_worksheet_checkpoint_router(["https://pane"], env[0]))
    client = TestClient(app, base_url="http://127.0.0.1", client=("127.0.0.1", 5000))
    data = request().model_dump(mode="json")
    assert client.post("/api/worksheet-checkpoints/execute", json=data).status_code == 403
    assert client.post("/api/worksheet-checkpoints/execute", json=data, headers={"Origin":"https://evil"}).status_code == 403
    assert not env[1].book.calls
    response = client.post("/api/worksheet-checkpoints/execute", json=data, headers={"Origin":"https://pane"})
    assert response.status_code == 200 and response.json()["status"] == "completed"
    data["command"] = "do something else"
    assert client.post("/api/worksheet-checkpoints/execute", json=data, headers={"Origin":"https://pane"}).status_code == 422


def test_other_service_lock_prevents_archive_and_delete(env):
    from server.app.folder_mutex import folder_write_guard, FolderWriteBusyError
    from server.app.app_data import application_data_root
    service, native = env
    with folder_write_guard(application_data_root() / "worksheet-checkpoints"):
        with pytest.raises(FolderWriteBusyError): service.execute(request(), "o")
    assert not native.book.calls and not hasattr(native.book, "archive")


def test_wps_global_switch_blocks_checkpoint_without_touching_host(env, monkeypatch):
    from copy import deepcopy
    from server.app import worksheet_checkpoint as module
    service, native = env
    config = deepcopy(module.capabilities()); config["wpsWrite"]["enabled"] = False
    monkeypatch.setattr(module, "capabilities", lambda: config)
    with pytest.raises(ValueError, match="WPS 写入能力当前关闭"): service.execute(request(), "o")
    assert not native.book.calls and not hasattr(native.book, "archive")


def consent(names):
    return WorksheetCheckpointBackupConsent(operationId="a"*32, backupSheets=names)


@pytest.mark.parametrize("host", ["excel", "wps-et"])
def test_hidden_sheets_need_explicit_backup_only_consent(env, host):
    service, native = env
    hidden = ["#EB_RULES", "#EB_RULES_BACKUP"]
    for name in hidden:
        sheet = Sheet(name, native.book); sheet.Visible = 2; native.book.items.append(sheet)
    pending = service.execute(request(host=host), "o")
    assert pending.status == "needs_backup_consent" and pending.backupSheets == hidden
    assert not hasattr(native.book, "archive") and not native.book.calls
    assert service.status(request().operationId, "o") == pending
    assert service.execute(request(host=host), "o") == pending
    result = service.confirm_backup(consent(hidden), "o")
    assert result.status == "completed" and result.deletedSheets == ["Sheet2", "Sheet3"]
    assert result.remainingSheets == ["Sheet1", *hidden]
    archived = load_workbook(native.book.archive)
    assert archived.sheetnames == ["Sheet1", "Sheet2", "Sheet3", *hidden]
    assert all(archived[name].sheet_state == "veryHidden" for name in hidden)
    state = json.loads((Path(native.book.archive).parent / "state.json").read_text(encoding="utf-8"))
    assert state["request"]["authorizedSheets"] == ["Sheet1", "Sheet2", "Sheet3"]
    assert state["backupConsent"] == hidden
    assert service.confirm_backup(consent(hidden), "o") == result
    assert native.book.calls == ["Sheet2", "Sheet3"], "consent retry must not replay deletion"


@pytest.mark.parametrize("names", [["Sheet1", "Other"], ["Other"], ["Sheet1", "Sheet1"]])
def test_consent_must_match_exact_missing_sheet_names(env, names):
    service, native = env
    service.execute(request(authorizedSheets=["Sheet2", "Sheet3"]), "o")
    with pytest.raises(ValueError, match="备份许可"):
        service.confirm_backup(consent(names), "o")
    assert not native.book.calls and not hasattr(native.book, "archive")


def test_sheet_catalog_change_while_waiting_for_consent_blocks_archive(env):
    service, native = env
    service.execute(request(authorizedSheets=["Sheet2", "Sheet3"]), "o")
    native.book.items.append(Sheet("AddedLater", native.book))
    result = service.confirm_backup(consent(["Sheet1"]), "o")
    assert result.status == "failed" and "列表发生变化" in result.message
    assert not native.book.calls and not hasattr(native.book, "archive")


def test_backup_permission_never_authorizes_deleting_an_unselected_sheet(env):
    service, native = env
    with pytest.raises(ValueError, match="超出授权"):
        service.execute(request(authorizedSheets=["Sheet1", "Sheet2"]), "o")
    assert not native.book.calls and not hasattr(native.book, "archive")


def test_confirm_backup_api_uses_stored_plan_and_origin(env):
    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from server.app.worksheet_checkpoint_api import create_worksheet_checkpoint_router
    app = FastAPI(); app.include_router(create_worksheet_checkpoint_router(["https://pane"], env[0]))
    client = TestClient(app, base_url="http://127.0.0.1", client=("127.0.0.1", 5000))
    headers = {"Origin": "https://pane"}
    pending = client.post("/api/worksheet-checkpoints/execute", json=request(authorizedSheets=["Sheet2", "Sheet3"]).model_dump(mode="json"), headers=headers)
    assert pending.json()["status"] == "needs_backup_consent"
    payload = consent(["Sheet1"]).model_dump()
    assert client.post("/api/worksheet-checkpoints/confirm-backup", json=payload).status_code == 403
    assert client.post("/api/worksheet-checkpoints/confirm-backup", json={**payload, "authorizedSheets":["Sheet1"]}, headers=headers).status_code == 422
    assert not env[1].book.calls
    assert client.post("/api/worksheet-checkpoints/confirm-backup", json=payload, headers=headers).json()["status"] == "completed"


def _write_state(root, operation_id, *, origin="o", host="wps-et", workbook=Book.FullName,
                 status="completed", created_at=None, conversation=None):
    folder = Path(root) / "worksheet-checkpoints" / operation_id
    folder.mkdir(parents=True)
    request_payload = request(host=host, workbookPath=workbook, conversationId=conversation).model_dump(mode="json")
    record = {
        "origin": origin,
        "request": request_payload,
        "result": {
            "operationId": operation_id, "status": status, "backupSheets": [],
            "deletedSheets": ["Sheet2"], "remainingSheets": ["Sheet1", "Sheet3"],
            "recoveryAvailable": True, "message": "已删除 1 张工作表。"
        },
        "sha256": "0" * 64,
    }
    if created_at is not None:
        record["createdAt"] = created_at
    (folder / "state.json").write_text(json.dumps(record), encoding="utf-8")
    return folder


def test_list_for_workbook_filters_by_origin_host_and_identity(env, tmp_path):
    from server.app.models import WorksheetCheckpointListRequest
    service, _ = env
    keep = "c" * 32
    _write_state(tmp_path, keep, created_at="2026-10-18T00:00:00+00:00")
    _write_state(tmp_path, "d" * 32, workbook=r"C:/work/other.xlsx")
    _write_state(tmp_path, "e" * 32, host="excel")
    _write_state(tmp_path, "f" * 32, origin="another-pane")

    listed = service.list_for_workbook(
        WorksheetCheckpointListRequest(workbookPath=r"c:/WORK/example.XLSX", host="wps-et"), "o")
    assert [item.operationId for item in listed.checkpoints] == [keep]
    assert listed.checkpoints[0].createdAt == "2026-10-18T00:00:00+00:00"
    assert listed.checkpoints[0].planTitle == "retain first"


def test_legacy_record_without_created_at_or_conversation_is_still_recoverable(env, tmp_path):
    from server.app.models import WorksheetCheckpointListRequest
    service, _ = env
    legacy = "1" * 32
    _write_state(tmp_path, legacy)
    listed = service.list_for_workbook(
        WorksheetCheckpointListRequest(workbookPath=Book.FullName, host="wps-et"), "o")
    assert [item.operationId for item in listed.checkpoints] == [legacy]
    assert listed.checkpoints[0].createdAt is None
    assert listed.checkpoints[0].recoveryAvailable is True
    # 旧记录仍能从服务端保存的请求元数据确认归属，但不会凭空多出会话信息。
    stored = json.loads((Path(tmp_path) / "worksheet-checkpoints" / legacy / "state.json").read_text(encoding="utf-8"))
    assert "conversationId" not in stored["request"] or stored["request"]["conversationId"] is None


def test_list_normalises_crashed_receipt_and_skips_damaged_records(env, tmp_path):
    from server.app.models import WorksheetCheckpointListRequest
    service, _ = env
    crashed = "2" * 32
    _write_state(tmp_path, crashed, status="executing")
    broken = Path(tmp_path) / "worksheet-checkpoints" / ("3" * 32)
    broken.mkdir(parents=True)
    (broken / "state.json").write_text("{not json", encoding="utf-8")

    listed = service.list_for_workbook(
        WorksheetCheckpointListRequest(workbookPath=Book.FullName, host="wps-et"), "o")
    assert [item.operationId for item in listed.checkpoints] == [crashed]
    assert listed.checkpoints[0].status == "partial_or_unknown"


def test_list_api_requires_explicit_origin_and_rejects_extra_fields(env, tmp_path):
    from fastapi import FastAPI
    from fastapi.testclient import TestClient
    from server.app.worksheet_checkpoint_api import create_worksheet_checkpoint_router
    service, _ = env
    keep = "4" * 32
    _write_state(tmp_path, keep, origin="https://pane")
    app = FastAPI(); app.include_router(create_worksheet_checkpoint_router(["https://pane"], service))
    client = TestClient(app, base_url="http://127.0.0.1", client=("127.0.0.1", 5000))
    body = {"workbookPath": Book.FullName, "host": "wps-et"}
    assert client.post("/api/worksheet-checkpoints/list", json=body).status_code == 403
    assert client.post("/api/worksheet-checkpoints/list", json={**body, "conversationId": "x"},
                       headers={"Origin": "https://pane"}).status_code == 422
    response = client.post("/api/worksheet-checkpoints/list", json=body, headers={"Origin": "https://pane"})
    assert response.status_code == 200
    assert [item["operationId"] for item in response.json()["checkpoints"]] == [keep]

from pathlib import Path
from unittest.mock import patch

import pytest
from openpyxl import Workbook
from pydantic import ValidationError
from server.app import folder_workbooks as folders
from server.app.folder_admission import check_folder_admission
from server.app.models import AnalysisPlan, FolderAdmissionRequest


def fixture(tmp_path: Path):
    path = tmp_path / "synthetic.xlsx"
    book = Workbook()
    book.active.title = "Selected"
    book.active.append(["PRIVATE_CELL_SENTINEL"])
    book.create_sheet("Other")
    book.save(path)
    book.close()
    catalog = folders.scan_folder(tmp_path)
    request = FolderAdmissionRequest(sessionId=catalog.sessionId, selections=[{"fileId": catalog.files[0].id, "sheets": ["Selected"]}])
    return path, request, folders._sessions[catalog.sessionId]


def test_selection_checks_metadata_without_reading_cells_or_authorizing(tmp_path):
    path, request, session = fixture(tmp_path)
    before = path.read_bytes()
    with patch.object(folders, "load_workbook", side_effect=AssertionError("no cell loader")):
        result = check_folder_admission(request)
    assert result.items[0].reason == "none"
    assert result.fingerprint
    assert "PRIVATE_CELL_SENTINEL" not in result.model_dump_json()
    assert not session.selected_targets and not session.selected_sources
    assert path.read_bytes() == before
    assert not (tmp_path / ".excel-bro").exists()


def test_missing_sheet_is_not_silently_dropped(tmp_path):
    _, request, _ = fixture(tmp_path)
    request.selections[0].sheets.append("Deleted")
    result = check_folder_admission(request)
    assert len(result.items) == 2
    assert result.items[1].reason == "missing"


def test_protected_sheet_is_analysis_only(tmp_path):
    path, request, _ = fixture(tmp_path)
    book = folders.load_workbook(path)
    book["Selected"].protection.sheet = True
    book.save(path)
    book.close()
    result = check_folder_admission(request)
    assert result.items[0].status == "read_only"
    assert result.items[0].reason == "protected"


def test_protected_sheet_cannot_bypass_the_write_guard_after_selection(tmp_path):
    path, request, session = fixture(tmp_path)
    book = folders.load_workbook(path)
    book["Selected"].protection.sheet = True
    book.save(path)
    book.close()
    snapshot = folders.create_folder_snapshot(folders.FolderSnapshotRequest.model_validate(request.model_dump()))
    target = snapshot.worksheets[0].name
    plan = AnalysisPlan.model_validate({"id": "protected-selection", "title": "Synthetic write", "summary": "Synthetic",
        "actions": [{"type": "writeValues", "sheet": target, "range": "A1", "values": [["CHANGED"]]}]})
    before = path.read_bytes()
    with pytest.raises(ValueError, match="受保护"):
        folders.validate_plan_for_folder(plan, session, authorized_sheets=[target])
    assert path.read_bytes() == before
    assert not (tmp_path / ".excel-bro").exists()


def test_replaced_or_unreadable_file_stays_in_selector(tmp_path):
    path, request, _ = fixture(tmp_path)
    path.write_bytes(b"not an xlsx")
    assert check_folder_admission(request).items[0].reason == "unreadable"
    path.unlink()
    assert check_folder_admission(request).items[0].reason == "missing"


def test_unknown_file_id_never_becomes_a_path(tmp_path):
    _, request, _ = fixture(tmp_path)
    request.selections[0].fileId = "C:/PRIVATE_PATH.xlsx"
    result = check_folder_admission(request)
    assert result.items[0].reason == "missing"
    assert result.fingerprint is None


def test_lock_check_is_read_only_and_not_a_grant(tmp_path):
    _, request, session = fixture(tmp_path)
    with patch.object(folders, "_is_locked_for_writing", return_value=True):
        assert check_folder_admission(request).items[0].reason == "locked"
    assert not session.selected_targets


def test_content_change_during_check_is_blocked(tmp_path):
    _, request, _ = fixture(tmp_path)
    with patch.object(folders, "_file_content_fingerprint", side_effect=["before", "after"]):
        assert check_folder_admission(request).items[0].reason == "target_changed"


def test_request_rejects_paths_and_duplicates(tmp_path):
    _, request, _ = fixture(tmp_path)
    data = request.model_dump()
    with pytest.raises(ValidationError):
        FolderAdmissionRequest.model_validate({**data, "path": "PRIVATE_PATH"})
    data["selections"][0]["sheets"] *= 2
    with pytest.raises(ValidationError):
        FolderAdmissionRequest.model_validate(data)


def test_local_admission_endpoint_does_not_authorize_or_echo_private_values(tmp_path):
    from fastapi.testclient import TestClient
    from server.app.main import app
    _, request, session = fixture(tmp_path)
    result = TestClient(app).post("/api/folders/admission", json=request.model_dump())
    assert result.status_code == 200
    assert result.json()["items"][0]["status"] == "ready"
    assert "PRIVATE_CELL_SENTINEL" not in result.text
    assert not session.selected_targets
    result = TestClient(app).post("/api/folders/admission", json={**request.model_dump(), "sessionId": "missing-session"})
    assert result.status_code == 422
    assert result.json()["detail"]["code"] == "FOLDER_SESSION_LOST"

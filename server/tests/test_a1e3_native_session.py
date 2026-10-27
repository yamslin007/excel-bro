"""Real fault/ledger integration for acceptance tooling; not native host evidence."""
import json
import time
from pathlib import Path
from unittest import mock
import pytest
from fastapi.testclient import TestClient
from scripts import a1e3_native as cli
from scripts.a1e3_native_server import FaultAudit, create_app
from server.app import dev_source_changed_fault as fault
from server.app.task_ledger import FolderOperationLedger

TARGET = {"hostId": "excel-office", "workbookName": "fixture.xlsx", "workbookIdentity": "url:C:/fixtures/a.xlsx"}

@pytest.fixture
def audit(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_BRO_SOURCE_CHANGED_FAULT", "1")
    fault.reset()
    instance = FaultAudit({"runId": "this-run"}, tmp_path)
    originals = instance.install(fault)
    yield instance
    for name, original in originals.items():
        setattr(fault, name, original)
    fault.reset()


def arm_and_bind():
    fault.arm(host_id=TARGET["hostId"], workbook_name=TARGET["workbookName"], workbook_identity=TARGET["workbookIdentity"])
    return fault.bind_task(task_id="turn-current", host_id=TARGET["hostId"], workbook_name=TARGET["workbookName"], workbook_identity=TARGET["workbookIdentity"])


def test_real_consumption_keeps_bound_task_after_production_clears_it(audit):
    assert arm_and_bind()
    assert fault.take("turn-current")
    assert fault.status()["taskBound"] is False
    attempt = audit.snapshot()["attempt"]
    assert attempt["boundTaskId"] == attempt["consumedTaskId"] == "turn-current"
    assert fault.take("turn-current") is False


def test_other_task_and_same_name_other_identity_do_not_steal(audit):
    fault.arm(host_id=TARGET["hostId"], workbook_name=TARGET["workbookName"], workbook_identity=TARGET["workbookIdentity"])
    assert not fault.bind_task(task_id="other", host_id=TARGET["hostId"], workbook_name=TARGET["workbookName"], workbook_identity="url:C:/other/a.xlsx")
    assert audit.snapshot()["attempt"]["boundTaskId"] is None
    assert not fault.take("other")
    assert fault.bind_task(task_id="turn-current", host_id=TARGET["hostId"], workbook_name=TARGET["workbookName"], workbook_identity=TARGET["workbookIdentity"])
    assert not fault.take("other")
    assert fault.take("turn-current")


def test_rearming_preserves_previous_attempt_file(audit, tmp_path):
    arm_and_bind()
    first = audit.snapshot()["attempt"]["attemptId"]
    fault.take("turn-current")
    fault.arm(host_id=TARGET["hostId"], workbook_name=TARGET["workbookName"], workbook_identity=TARGET["workbookIdentity"])
    assert audit.snapshot()["attempt"]["attemptId"] != first
    assert len(list((tmp_path / "attempts").glob("*.json"))) == 2
    assert json.loads((tmp_path / "attempts" / (first + ".json")).read_text())["consumedTaskId"] == "turn-current"


def test_audit_io_failure_does_not_change_real_fault_behavior(audit):
    with mock.patch.object(Path, "write_text", side_effect=OSError("disk")):
        assert arm_and_bind()
        assert fault.take("turn-current")
    assert audit.snapshot()["auditError"] == "OSError"


@pytest.fixture
def evidence_context(tmp_path, monkeypatch):
    directory = tmp_path / "config"
    ledger = FolderOperationLedger(lambda: directory / "tasks" / "state.sqlite3")
    request = {"hostId": TARGET["hostId"], "workbook": {"name": TARGET["workbookName"], "workbookIdentity": TARGET["workbookIdentity"]}}
    armed_at = time.time()
    for task_id in ("turn-current", "turn-other-newer"):
        ledger.register_planning_task(task_id, task_id, request, approved_wall_time_seconds=20)
        ledger.finish_planning_task(task_id, status="failed", error={"code": "SOURCE_CHANGED"})
    attempt = {"runId": "this-run", "attemptId": "attempt", "armedAt": armed_at,
               "target": TARGET.copy(), "boundTaskId": "turn-current", "consumedTaskId": "turn-current", "cancelled": False}
    session = {"runId": "this-run", "attempt": attempt, "bridgeTargets": [TARGET], "auditError": None}
    monkeypatch.setattr(cli, "_SESSION", {"runId": "this-run", "base": "http://test", "token": "token"})
    monkeypatch.setattr(cli, "_session", lambda base: session)
    monkeypatch.setattr(cli, "_status", lambda base: {"phase": "consumed", "taskBound": False})
    return directory, session, ledger


def test_evidence_uses_exact_bound_task_not_latest(evidence_context, tmp_path):
    directory, session, ledger = evidence_context
    evidence = cli._collect_evidence("http://test", str(directory), None, tmp_path / "out")
    assert evidence["turn"]["turn_id"] == "turn-current"
    assert evidence["planningStopConfirmed"] is True
    assert evidence["executionEvidence"] == "未进入执行"
    assert evidence["nativeUiAndWorkbook"] == "pending"
    assert "request_json" not in json.dumps(evidence)
    again = cli._collect_evidence("http://test", str(directory), None, tmp_path / "out")
    assert evidence["evidencePath"] != again["evidencePath"]
    assert Path(evidence["evidencePath"]).exists()


@pytest.mark.parametrize("change", ["not_consumed", "cancelled", "wrong_host", "missing_task", "before_arm", "wrong_target"])
def test_insufficient_evidence_never_becomes_unentered(evidence_context, tmp_path, change):
    directory, session, ledger = evidence_context
    host = None
    if change == "not_consumed": session["attempt"]["consumedTaskId"] = None
    if change == "cancelled": session["attempt"]["cancelled"] = True
    if change == "wrong_host": host = "wps-et"
    if change == "missing_task":
        session["attempt"].update(boundTaskId="missing", consumedTaskId="missing")
    if change == "before_arm": session["attempt"]["armedAt"] += 100
    if change == "wrong_target": session["attempt"]["target"]["workbookIdentity"] = "url:another"
    with pytest.raises(cli.ServiceError):
        cli._collect_evidence("http://test", str(directory), host, tmp_path / "out")


def test_service_errors_from_host_observation_are_not_swallowed(evidence_context, tmp_path, monkeypatch):
    directory, session, ledger = evidence_context
    monkeypatch.setattr(cli, "_readonly_rows", lambda db, sql, params: [
        {"turn_id": "turn-current", "status": "failed", "last_sequence": 1,
         "request_json": json.dumps({"hostId": TARGET["hostId"], "workbook": {"name": TARGET["workbookName"], "workbookIdentity": TARGET["workbookIdentity"]}}),
         "error_json": '{"code":"SOURCE_CHANGED"}', "created_at": time.time(), "updated_at": time.time()}
    ] if "planning_tasks" in sql else [{"operation_id":"op-current"}])
    with mock.patch.object(cli, "_request", side_effect=cli.ServiceError("HTTP 403")) as send:
        with pytest.raises(cli.ServiceError, match="403"):
            cli._collect_evidence("http://test", str(directory), None, tmp_path / "out")
    assert send.call_args.args[1] == "/api/office/executions/op-current/host-writes"
    assert send.call_count == 1


def test_bare_arm_noninteractive_does_not_report_success(monkeypatch):
    monkeypatch.setattr(cli, "_health", lambda base: True)
    monkeypatch.setattr(cli, "_status", lambda base: {"seenTargets": [TARGET]})
    with mock.patch.object(cli.sys.stdin, "isatty", return_value=False), mock.patch.object(cli, "_request") as send:
        assert cli.cmd_arm("http://test", None, None, None, 120) == 3
    send.assert_not_called()


def test_wrapper_is_explicit_and_token_guarded(tmp_path, monkeypatch):
    config = tmp_path / "server-config"
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps({"runId":"test-run", "token":"a-test-token"}), encoding="utf-8")
    monkeypatch.setenv("EB_NATIVE_MANIFEST", str(manifest))
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(config))
    monkeypatch.setenv("EXCEL_BRO_SOURCE_CHANGED_FAULT", "1")
    from server.app import main
    monkeypatch.setattr(main, "_folder_operation_ledger", FolderOperationLedger(lambda: config / "tasks" / "state.sqlite3"))
    originals = {name:getattr(fault,name) for name in ("arm","bind_task","take","cancel")}
    try:
        wrapper = create_app()
        with TestClient(wrapper, base_url="http://127.0.0.1", client=("127.0.0.1",43210)) as client:
            assert client.get("/api/dev/native-acceptance/session").status_code == 403
            response = client.get("/api/dev/native-acceptance/session", headers={"X-Native-Acceptance-Token":"a-test-token"})
            assert response.status_code == 200
            assert response.json()["runId"] == "test-run"
            assert client.get("/api/capabilities").status_code == 200
        assert all(getattr(route,"path",None) != "/api/dev/native-acceptance/session" for route in main.app.routes)
    finally:
        for name, original in originals.items(): setattr(fault,name,original)
        fault.reset()


def test_wrapper_rejects_sdk_state_inside_repository(tmp_path, monkeypatch):
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps({"runId":"bad-sdk", "token":"a-test-token"}), encoding="utf-8")
    monkeypatch.setenv("EB_NATIVE_MANIFEST", str(manifest))
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "server-config"))
    monkeypatch.setenv("EXCEL_BRO_SOURCE_CHANGED_FAULT", "1")
    monkeypatch.setenv("CLAUDE_AGENT_SDK_STATE_DIR", str(Path(__file__).resolve().parents[2] / "outputs" / "invalid-sdk-dir"))
    with pytest.raises(RuntimeError, match="outside the repository"):
        create_app()


def test_fixture_refuses_to_overwrite_open_or_previous_evidence(tmp_path):
    from scripts.make_a1e3_native_fixture import build
    from openpyxl import load_workbook
    target = tmp_path / "fixture.xlsx"
    build(target)
    before = target.read_bytes()
    with pytest.raises(FileExistsError): build(target)
    assert target.read_bytes() == before
    book = load_workbook(target, read_only=True)
    try:
        assert [book["来源"].cell(row,2).value for row in (2,3,4)] == [111,250,333]
        assert [book["目标"].cell(row,2).value for row in (2,3,4)] == [-101,-202,-303]
    finally: book.close()


@pytest.mark.skipif(__import__("os").name != "nt", reason="Windows process ownership test")
def test_stop_rejects_live_unrelated_pid_from_another_cwd(tmp_path):
    import os
    import subprocess
    script = Path(__file__).resolve().parents[2] / "scripts" / "a1e3_native.ps1"
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps({"base":"http://127.0.0.1:8877", "port":8877,
        "process":{"pid":os.getpid(), "created":"wrong", "executable":"wrong", "command":"wrong"}}), encoding="utf-8")
    result = subprocess.run(["powershell.exe","-NoProfile","-File",str(script),"stop","-Manifest",str(manifest)], cwd=tmp_path, capture_output=True, timeout=20)
    assert result.returncode != 0
    assert b"Process identity mismatch" in result.stderr


@pytest.mark.skipif(__import__("os").name != "nt", reason="Windows launcher test")
def test_prepare_rejects_inconsistent_base_before_creating_a_run(tmp_path):
    import subprocess
    script = Path(__file__).resolve().parents[2] / "scripts" / "a1e3_native.ps1"
    result = subprocess.run(["powershell.exe","-NoProfile","-File",str(script),"prepare","-Port","8877","-Base","http://127.0.0.1:8888"], cwd=tmp_path, capture_output=True, timeout=20)
    assert result.returncode != 0
    assert b"Base must agree with Port" in result.stderr

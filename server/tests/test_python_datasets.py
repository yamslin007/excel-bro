import json
import math

import pytest
from fastapi.testclient import TestClient

from server.app.models import PythonDataScope, PythonDatasetReceipt
from server.app.python_datasets import PythonDatasetRegistry, DatasetRejected
from server.app import main, python_data_api, python_analysis

ORIGIN = "https://localhost:3000"


def scope(**updates):
    data = dict(version=1, taskId="synthetic-task", hostId="excel-office",
                workbookIdentityHash="a"*64, authorizationVersion="b"*64,
                sources=[dict(sheetId="sheet-1", sheetName="Sheet1", range="A1:B2")])
    data.update(updates)
    return PythonDataScope.model_validate(data)


def prepared(**kwargs):
    registry = PythonDatasetRegistry(**kwargs)
    target = scope()
    token = registry.create(target, ORIGIN)
    receipt = registry.register(token, target, ORIGIN, 0, [[1, "PRIVATE_DATA"], [2, None]])
    return registry, target, token, receipt


def test_snapshot_is_immutable_and_locally_resolved():
    registry, target, token, receipt = prepared()
    data = registry.resolve(token, target, ORIGIN, [receipt])
    data["datasets"][0]["values"][0][0] = 99
    assert registry.resolve(token, target, ORIGIN, [receipt])["datasets"][0]["values"][0][0] == 1
    assert "PRIVATE_DATA" not in receipt.model_dump_json() + repr(registry._grants)


@pytest.mark.parametrize("change", ["task", "file", "authorization", "scope", "origin", "token", "fingerprint", "handle"])
def test_exact_binding_required(change):
    registry, target, token, receipt = prepared()
    origin = ORIGIN
    if change == "task": target = scope(taskId="another-task")
    if change == "file": target = scope(workbookIdentityHash="c"*64)
    if change == "authorization": target = scope(authorizationVersion="c"*64)
    if change == "scope": target = scope(sources=[dict(sheetId="2", sheetName="Sheet2", range="A1:B2")])
    if change == "origin": origin = "http://localhost:3889"
    if change == "token": token = "z"*43
    if change == "fingerprint": receipt = receipt.model_copy(update={"contentFingerprint": "f"*64})
    if change == "handle": receipt = PythonDatasetReceipt(datasetHandle="dataset://"+"z"*43, contentFingerprint="f"*64)
    with pytest.raises(DatasetRejected): registry.resolve(token, target, origin, [receipt])


def test_replacement_revocation_and_expiry():
    now = [0]
    registry, target, token, receipt = prepared(clock=lambda: now[0])
    registry.create(scope(authorizationVersion="c"*64), ORIGIN)
    assert not registry.valid(token, target, ORIGIN, [receipt])
    token = registry.create(target, ORIGIN)
    receipt = registry.register(token, target, ORIGIN, 0, [[1, 2], [3, 4]])
    registry.revoke(token, target, ORIGIN)
    assert not registry.valid(token, target, ORIGIN, [receipt])
    token = registry.create(target, ORIGIN)
    now[0] = 901
    with pytest.raises(DatasetRejected): registry.register(token, target, ORIGIN, 0, [[1, 2], [3, 4]])


@pytest.mark.parametrize("values", [[[1]], [[1, 2]], [[{}, 2], [3, 4]], [[math.nan, 2], [3, 4]], [[1, 2, 3], [4, 5, 6]]])
def test_invalid_values_rejected_without_echo(values):
    registry, target, token, _ = prepared()
    with pytest.raises(DatasetRejected) as error: registry.register(token, target, ORIGIN, 0, values)
    assert "nan" not in str(error.value).lower()


def test_handles_replace_not_accumulate_and_call_limit():
    registry, target, token, receipt = prepared()
    new = registry.register(token, target, ORIGIN, 0, [[2, 3], [4, 5]])
    assert not registry.valid(token, target, ORIGIN, [receipt])
    for _ in range(12): registry.resolve(token, target, ORIGIN, [new])
    with pytest.raises(DatasetRejected): registry.resolve(token, target, ORIGIN, [new])
    with pytest.raises(DatasetRejected): registry.resolve(token, target, ORIGIN, [new, new])


def client(origin=ORIGIN, host="127.0.0.1", base="http://127.0.0.1:8765"):
    return TestClient(main.app, base_url=base, client=(host, 5000), headers={"Origin": origin})


@pytest.mark.parametrize("origin,host,base", [("", "127.0.0.1", "http://localhost"),
    ("https://evil.example", "127.0.0.1", "http://localhost"),
    (ORIGIN, "10.0.0.1", "http://localhost"), (ORIGIN, "127.0.0.1", "http://remote.example")])
def test_api_rejects_remote_or_missing_origin(monkeypatch, origin, host, base):
    monkeypatch.setattr(python_data_api, "python_execution_enabled", lambda: True)
    response = client(origin, host, base).post("/api/python/data/scope", json={"scope": scope().model_dump()})
    assert response.status_code == 403


def test_api_default_off_and_validation_never_echoes_rows(monkeypatch):
    response = client().post("/api/python/data/scope", json={"scope": scope().model_dump()})
    assert response.status_code == 409 and "未启用" in response.text
    monkeypatch.setattr(python_data_api, "python_execution_enabled", lambda: True)
    for body in ({"scope": {}, "values": "PRIVATE_ROW_SENTINEL"}, {"scope": scope().model_dump(), "path": "PRIVATE_ROW_SENTINEL"}):
        response = client().post("/api/python/data/scope", json=body)
        assert response.status_code == 409 and "PRIVATE_ROW_SENTINEL" not in response.text


@pytest.mark.parametrize("host_id,origin", [("excel-office", ORIGIN), ("wps-et", "http://localhost:3889")])
def test_local_api_both_hosts_roundtrip(monkeypatch, host_id, origin):
    monkeypatch.setattr(python_data_api, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(python_data_api, "registry", PythonDatasetRegistry())
    target = scope(hostId=host_id).model_dump()
    local = client(origin)
    response = local.post("/api/python/data/scope", json={"scope": target})
    assert response.status_code == 200 and response.headers["cache-control"] == "no-store"
    token = response.json()["scopeToken"]
    snapshot = local.post("/api/python/data/snapshot", json=dict(scope=target, scopeToken=token, sourceIndex=0, values=[[1, "PRIVATE_ROW"], [2, None]]))
    assert snapshot.status_code == 200 and "PRIVATE_ROW" not in snapshot.text
    revoke = local.post("/api/python/data/revoke", json=dict(scope=target, scopeToken=token))
    assert revoke.status_code == 200
    assert local.post("/api/python/data/snapshot", json=dict(scope=target, scopeToken=token, sourceIndex=0, values=[[1, 2], [3, 4]])).status_code == 409


def test_late_result_discarded_when_scope_revoked(monkeypatch):
    registry, target, token, receipt = prepared()
    def fake_runner(code, values, *, cancel_check, on_runner_start=None):
        assert values["datasets"][0]["values"][0][1] == "PRIVATE_DATA"
        if on_runner_start:
            on_runner_start()
        registry.revoke(token, target, ORIGIN)
        assert cancel_check()
        return {"status": "completed", "result": "PRIVATE_LATE_RESULT"}
    monkeypatch.setattr(python_analysis, "run_isolated_analysis_internal", fake_runner)
    result = python_analysis.run_registered_analysis_internal(data_registry=registry,
        scope_token=token, scope=target, origin=ORIGIN, receipts=[receipt], code="result=1")
    assert result == {"status": "cancelled", "reasonCategory": "authorization_changed"}


def test_invalid_grant_never_reaches_runner(monkeypatch):
    registry, target, token, receipt = prepared()
    monkeypatch.setattr(python_analysis, "run_isolated_analysis_internal", lambda *args, **kwargs: pytest.fail("runner reached"))
    result = python_analysis.run_registered_analysis_internal(data_registry=registry,
        scope_token=token, scope=scope(taskId="other"), origin=ORIGIN, receipts=[receipt], code="result=1")
    assert result["status"] == "rejected"


def test_request_and_storage_limits_fail_closed(monkeypatch):
    from server.app import python_datasets
    real_limit = python_datasets.capability_int
    monkeypatch.setattr(python_datasets, "capability_int", lambda section, key:
        10 if key in ("maxDatasetBytes", "maxRegistryBytes") else real_limit(section, key))
    registry = PythonDatasetRegistry()
    target = scope()
    token = registry.create(target, ORIGIN)
    with pytest.raises(DatasetRejected): registry.register(token, target, ORIGIN, 0, [[1, "PRIVATE_OVERSIZED"], [2, 3]])
    monkeypatch.setattr(python_data_api, "python_execution_enabled", lambda: True)
    real_api_limit = python_data_api.capability_int
    monkeypatch.setattr(python_data_api, "capability_int", lambda section, key: 100 if key == "maxRequestBytes" else real_api_limit(section, key))
    response = client().post("/api/python/data/snapshot", content=json.dumps({"PRIVATE_BODY": "x"*200}), headers={"Content-Type": "application/json"})
    assert response.status_code == 409 and "PRIVATE_BODY" not in response.text


def test_host_origin_mismatch_and_no_plaintext_paths(monkeypatch):
    monkeypatch.setattr(python_data_api, "python_execution_enabled", lambda: True)
    response = client().post("/api/python/data/scope", json={"scope": scope(hostId="wps-et").model_dump()})
    assert response.status_code == 409
    raw = scope().model_dump()
    raw["workbookIdentityHash"] = "C:/PRIVATE/file.xlsx"
    response = client().post("/api/python/data/scope", json={"scope": raw})
    assert response.status_code == 409 and "PRIVATE" not in response.text


def test_ready_runtime_does_not_implicitly_enable_product(monkeypatch):
    monkeypatch.setattr(python_analysis, "python_execution_enabled", lambda: True)
    monkeypatch.setattr(python_analysis, "_runner_status", lambda: "ready")
    monkeypatch.setattr(python_analysis, "run_isolated_analysis_internal", lambda *args, **kwargs: pytest.fail("product executed"))
    result = python_analysis.execute_python_analysis(None, {"code": "result=1", "datasetHandles": []})
    assert result["reasonCategory"] == "product_not_ready"
    assert python_analysis.project_python_analysis_receipt(result) == result


def test_reauthorization_or_revocation_cannot_reset_task_budget_or_ttl():
    now = [0]
    registry, target, token, receipt = prepared(clock=lambda: now[0])
    for _ in range(12): registry.resolve(token, target, ORIGIN, [receipt])
    registry.revoke(token, target, ORIGIN)
    now[0] = 899
    new_target = scope(authorizationVersion="c"*64)
    new_token = registry.create(new_target, ORIGIN)
    new_receipt = registry.register(new_token, new_target, ORIGIN, 0, [[1, 2], [3, 4]])
    with pytest.raises(DatasetRejected): registry.resolve(new_token, new_target, ORIGIN, [new_receipt])
    now[0] = 901
    assert not registry.valid(new_token, new_target, ORIGIN, [new_receipt])

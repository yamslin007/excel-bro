import hashlib
import math

import pytest

from server.app import python_analysis, python_data_api
from server.app.python_datasets import DatasetRejected, PythonDatasetRegistry
from server.tests.test_python_datasets import prepared, scope, client, ORIGIN


def published(**kwargs):
    registry, target, token, data = prepared(**kwargs)
    receipt = registry.publish_result(token, target, ORIGIN, [data],
        {"kind": "matrix", "values": [[3, "PRIVATE_RESULT"]]})
    return registry, target, token, data, receipt


def test_immutable_result_and_metadata_only_receipt():
    registry, target, token, data, receipt = published()
    result = registry.read_result(token, target, ORIGIN, [data], receipt)
    assert result == {"kind": "matrix", "values": [[3, "PRIVATE_RESULT"]]}
    result["values"][0][0] = 99
    assert registry.read_result(token, target, ORIGIN, [data], receipt)["values"][0][0] == 3
    assert "PRIVATE_RESULT" not in receipt.model_dump_json() + repr(registry._grants)


@pytest.mark.parametrize("change", ["task", "file", "authorization", "origin", "token", "source", "result", "fingerprint", "shape"])
def test_results_require_every_binding(change):
    registry, target, token, data, receipt = published()
    origin = ORIGIN
    if change == "task": target = scope(taskId="other")
    if change == "file": target = scope(workbookIdentityHash="c"*64)
    if change == "authorization": target = scope(authorizationVersion="c"*64)
    if change == "origin": origin = "http://localhost:3889"
    if change == "token": token = "x"*43
    if change == "source": data = data.model_copy(update={"datasetHandle":"dataset://"+"x"*43})
    if change == "result": receipt = receipt.model_copy(update={"resultHandle":"result://"+"x"*43})
    if change == "fingerprint": receipt = receipt.model_copy(update={"contentFingerprint":"c"*64})
    if change == "shape": receipt = receipt.model_copy(update={"rows":2})
    with pytest.raises(DatasetRejected): registry.read_result(token, target, origin, [data], receipt)


@pytest.mark.parametrize("operation", ["revoke", "replace", "expiry", "discard"])
def test_invalidated_results_cannot_be_read(operation):
    now = [0]
    registry, target, token, data, receipt = published(clock=lambda: now[0])
    if operation == "revoke": registry.revoke(token, target, ORIGIN)
    if operation == "replace": registry.register(token, target, ORIGIN, 0, [[2, 3], [4, 5]])
    if operation == "expiry": now[0] = 901
    if operation == "discard": registry.discard_result(token, target, ORIGIN, [data], receipt)
    with pytest.raises(DatasetRejected): registry.read_result(token, target, ORIGIN, [data], receipt)


@pytest.mark.parametrize("result", [None, {"kind":"matrix","values":[]},
    {"kind":"matrix","values":[[1],[2,3]]}, {"kind":"matrix","values":[[{}]]},
    {"kind":"matrix","values":[[math.nan]]}, {"kind":"matrix","values":[[math.inf]]},
    {"kind":"matrix","values":[[2**53]]}, {"kind":"matrix","values":[[float(2**53)]]}, {"kind":"matrix","values":[["x"*32768]]},
    {"kind":"matrix","values":[["\ud800"]]}, {"kind":"matrix","values":[[1]],"actions":[]},
    {"kind":"formula","values":[["=1"]]}])
def test_untrusted_output_has_strict_matrix_contract(result):
    registry, target, token, data = prepared()
    with pytest.raises(DatasetRejected): registry.publish_result(token, target, ORIGIN, [data], result)


def test_internal_publication_returns_no_values_or_worker_fields(monkeypatch):
    registry, target, token, data = prepared()
    monkeypatch.setattr(python_analysis, "run_isolated_analysis_internal", lambda *args, **kwargs:
        {"status":"completed", "result":{"kind":"matrix","values":[[3,"PRIVATE_RESULT"]]}})
    response = python_analysis.run_registered_result_internal(data_registry=registry,
        scope_token=token, scope=target, origin=ORIGIN, receipts=[data], code="result=1")
    assert response["status"] == "completed"
    assert set(response) == {"status", "receipt"}
    assert "PRIVATE_RESULT" not in str(response)


def test_failed_or_cancelled_runner_publishes_nothing(monkeypatch):
    registry, target, token, data = prepared()
    def fake(*args, **kwargs):
        registry.revoke(token, target, ORIGIN)
        return {"status":"completed", "result":{"kind":"matrix","values":[["PRIVATE_LATE"]]}}
    monkeypatch.setattr(python_analysis, "run_isolated_analysis_internal", fake)
    response = python_analysis.run_registered_result_internal(data_registry=registry,
        scope_token=token, scope=target, origin=ORIGIN, receipts=[data], code="result=1")
    assert response == {"status":"cancelled", "reasonCategory":"authorization_changed"}


@pytest.mark.parametrize("host_id,origin", [("excel-office", ORIGIN), ("wps-et","http://localhost:3889")])
def test_local_only_fetch_and_discard_both_hosts(monkeypatch, host_id, origin):
    registry = PythonDatasetRegistry()
    target = scope(hostId=host_id)
    token = registry.create(target, origin)
    data = registry.register(token, target, origin, 0, [[1,2],[3,4]])
    receipt = registry.publish_result(token, target, origin, [data], {"kind":"matrix", "values":[["PRIVATE_RESULT"]]})
    monkeypatch.setattr(python_data_api, "registry", registry)
    monkeypatch.setattr(python_data_api, "python_execution_enabled", lambda: True)
    body = dict(scope=target.model_dump(), scopeToken=token, datasets=[data.model_dump()], result=receipt.model_dump())
    local = client(origin)
    read = local.post("/api/python/data/result", json=body)
    assert read.status_code == 200 and read.headers["cache-control"] == "no-store"
    assert read.json()["values"] == [["PRIVATE_RESULT"]]
    remote = client(origin, host="10.0.0.1").post("/api/python/data/result", json=body)
    assert remote.status_code == 403 and "PRIVATE_RESULT" not in remote.text
    invalid = local.post("/api/python/data/result", json={**body,"stdout":"PRIVATE_ERROR"})
    assert invalid.status_code == 409 and "PRIVATE_ERROR" not in invalid.text
    assert local.post("/api/python/data/result/discard", json=body).status_code == 200
    assert local.post("/api/python/data/result", json=body).status_code == 409


def test_http_bytes_match_receipt_including_python_float_encoding(monkeypatch):
    registry, target, token, data = prepared()
    values = [[1.0, 1e-7, -0.0, "中文😀", None, True]]
    receipt = registry.publish_result(token,target,ORIGIN,[data],{"kind":"matrix","values":values})
    monkeypatch.setattr(python_data_api,"registry",registry)
    monkeypatch.setattr(python_data_api,"python_execution_enabled",lambda:True)
    body = dict(scope=target.model_dump(),scopeToken=token,datasets=[data.model_dump()],result=receipt.model_dump())
    response = client().post("/api/python/data/result",json=body)
    assert response.status_code == 200
    assert hashlib.sha256(response.content).hexdigest() == receipt.contentFingerprint
    assert response.json()["values"] == values


@pytest.mark.parametrize("limit,value,result", [
    ("maxRows",1,{"kind":"matrix","values":[[1],[2]]}),
    ("maxColumns",1,{"kind":"matrix","values":[[1,2]]}),
    ("maxCells",1,{"kind":"matrix","values":[[1,2]]}),
    ("maxCellChars",1,{"kind":"matrix","values":[["😀"]]}),
    ("maxBytes",10,{"kind":"matrix","values":[[1]]}),
    ("maxRegistryBytes",10,{"kind":"matrix","values":[[1]]}),
    ("maxResultsPerScope",1,{"kind":"matrix","values":[[2]]}),
])
def test_result_limits_are_enforced_not_only_declared(monkeypatch,limit,value,result):
    from server.app.capabilities import capability_int
    from server.app import python_results, python_datasets
    registry,target,token,data,_ = published()
    def limited(section,key):
        return value if section == "pythonResult" and key == limit else capability_int(section,key)
    monkeypatch.setattr(python_results,"capability_int",limited)
    monkeypatch.setattr(python_datasets,"capability_int",limited)
    with pytest.raises(DatasetRejected): registry.publish_result(token,target,ORIGIN,[data],result)


def test_exact_order_of_inputs_and_duplicates_are_not_interchangeable():
    target = scope(sources=[dict(sheetId="1",sheetName="Sheet1",range="A1"),dict(sheetId="1",sheetName="Sheet1",range="B1")])
    registry = PythonDatasetRegistry(); token = registry.create(target,ORIGIN)
    first = registry.register(token,target,ORIGIN,0,[[1]])
    second = registry.register(token,target,ORIGIN,1,[[2]])
    receipt = registry.publish_result(token,target,ORIGIN,[first,second],{"kind":"matrix","values":[[1]]})
    for inputs in ([second,first],[first],[first,first]):
        with pytest.raises(DatasetRejected): registry.read_result(token,target,ORIGIN,inputs,receipt)


def test_new_authorization_restart_and_expiry_during_encoding_invalidate(monkeypatch):
    from server.app import python_datasets
    registry,target,token,data,receipt = published()
    registry.create(target,ORIGIN)
    with pytest.raises(DatasetRejected): registry.read_result(token,target,ORIGIN,[data],receipt)
    with pytest.raises(DatasetRejected): PythonDatasetRegistry().read_result(token,target,ORIGIN,[data],receipt)
    now = [0]; registry,target,token,data = prepared(clock=lambda:now[0])
    original = python_datasets.encode_matrix_result
    def expire(result):
        encoded = original(result); now[0] = 901; return encoded
    monkeypatch.setattr(python_datasets,"encode_matrix_result",expire)
    with pytest.raises(DatasetRejected): registry.publish_result(token,target,ORIGIN,[data],{"kind":"matrix","values":[[1]]})
    assert not registry._grants


def test_cancel_after_publication_discards_late_result(monkeypatch):
    registry,target,token,data = prepared()
    cancelled = [False]
    original = registry.publish_result
    def publish(*args,**kwargs):
        receipt = original(*args,**kwargs); cancelled[0] = True; return receipt
    monkeypatch.setattr(registry,"publish_result",publish)
    monkeypatch.setattr(python_analysis,"run_isolated_analysis_internal",lambda *args,**kwargs:
        {"status":"completed","result":{"kind":"matrix","values":[["PRIVATE_LATE"]]}})
    response = python_analysis.run_registered_result_internal(data_registry=registry,scope_token=token,
        scope=target,origin=ORIGIN,receipts=[data],code="result=1",cancel_check=lambda:cancelled[0])
    assert response == {"status":"cancelled","reasonCategory":"authorization_changed"}
    assert not registry._grants[token].results
    assert "PRIVATE_LATE" not in str(response)


def test_product_tool_and_api_remain_closed_even_with_local_result(monkeypatch):
    _,target,token,data,receipt = published()
    body = dict(scope=target.model_dump(),scopeToken=token,datasets=[data.model_dump()],result=receipt.model_dump())
    assert client().post("/api/python/data/result",json=body).status_code == 409
    monkeypatch.setattr(python_analysis,"python_execution_enabled",lambda:True)
    monkeypatch.setattr(python_analysis,"_runner_status",lambda:"verified")
    assert python_analysis.execute_python_analysis(None,{"code":"result=1","datasetHandles":[data.datasetHandle]})["reasonCategory"] == "product_not_ready"
    with pytest.raises(ValueError): python_analysis.project_python_analysis_receipt({"status":"completed","receipt":receipt.model_dump()})

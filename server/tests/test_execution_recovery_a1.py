"""A1: real HTTP routes + SQLite; synthetic host evidence, not native-host proof."""
from concurrent.futures import ThreadPoolExecutor
from threading import Barrier

import pytest
from fastapi.testclient import TestClient
import server.app.main as main
from server.app.task_ledger import FolderOperationLedger


@pytest.fixture(params=["office", "wps"])
def execution(request, monkeypatch, tmp_path):
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "config"))
    monkeypatch.setattr(main, "_folder_operation_ledger", FolderOperationLedger(lambda: tmp_path / "ledger.sqlite3"))
    host = request.param
    client = TestClient(main.app)
    op = f"{host}-a1-recovery"
    base = f"/api/{host}/executions/{op}"
    actions = [dict(type="writeValues", sheet="Sheet1", range=f"A{i+1}:A{i+1}", values=[[i+10]]) for i in range(3)]
    payload = dict(operationId=op, confirmed=True, authorizedSheetNames=["Sheet1"],
                   workbookIdentity=("url:" if host == "office" else "wps-path:") + "c:/a1.xlsx")
    if host == "office":
        payload.update(plan=dict(id="a1-plan", title="A1", summary="three actions", assumptions=[], warnings=[], actions=actions), allowedExternalFiles=[])
    else:
        payload.update(taskId="a1-task", descriptor=dict(actions=[dict(actionIndex=i, type=a["type"], sheet=a["sheet"], range=a["range"], rowCount=1, columnCount=1, cellCount=1, previewFingerprint=f"v2:before-{i}", expectedFingerprint=f"v2:after-{i}") for i,a in enumerate(actions)]))
    response = client.post(f"/api/{host}/executions/register", json=payload)
    assert response.status_code == 200, response.text

    class Execution:
        state = response.json()

        def evidence(self, index, after=False):
            if host == "wps":
                value = dict(sheetName="Sheet1", range=actions[index]["range"], type="writeValues", beforeFingerprint=f"v2:before-{index}")
                if after:
                    value["afterFingerprint"] = f"v2:after-{index}"
                return value
            before = dict(actionIndex=index, sheet="Sheet1", range=actions[index]["range"], formulas=[[None]], numberFormat=[["General"]])
            if not after:
                return before
            return dict(planId="a1-plan", executionId=op, workbookIdentity=payload["workbookIdentity"], completedActionCount=index+1,
                        ranges=[dict(actionIndex=i, sheet="Sheet1", range=actions[i]["range"], formulas=[[None]], postFingerprint=f"fnv1a32:{i:08x}") for i in range(index+1)], coverage="full", postStateCoverage="complete")

        def post(self, route, **fields):
            return client.post(base + "/" + route, json=dict(expectedRevision=self.state["revision"], **fields))

        def accept(self, route, **fields):
            response = self.post(route, **fields)
            assert response.status_code == 200, response.text
            self.state = response.json()
            return self.state

        def start(self, index):
            return self.accept("actions/start", ownerToken=self.state["ownerToken"], actionIndex=index, beforeSnapshot=self.evidence(index))

        def complete(self, index):
            return self.accept("actions/complete", ownerToken=self.state["ownerToken"], actionIndex=index, checkpoint=self.evidence(index, True))

        def suspend(self):
            return self.accept("suspend", ownerToken=self.state["ownerToken"], code="CONNECTION_LOST", message="test interruption")

        def status(self):
            response = client.get(base)
            assert response.status_code == 200, response.text
            return response.json()

    value = Execution()
    value.client, value.base, value.host = client, base, host
    yield value
    client.close()


def test_a1_inflight_cannot_claim_even_after_owner_stops(execution):
    e = execution
    e.start(0)
    e.suspend()
    before = e.status()
    assert before["inFlightActionIndex"] == 0
    assert e.post("resume", confirmed=True).status_code == 409
    assert e.status() == before


@pytest.mark.parametrize("outcome,next_index", [("applied", 2), ("not_applied", 1)])
def test_a1_three_actions_lost_receipt_resumes_exact_remaining_suffix(execution, outcome, next_index):
    e = execution
    e.start(0)
    e.complete(0)
    e.start(1)
    # A host write may have happened, but its completion receipt never arrived.
    e.suspend()
    observed = e.status()
    assert (observed["nextActionIndex"], observed["inFlightActionIndex"]) == (1, 1)
    assert observed["status"] == "needs_reconciliation"
    fields = dict(actionIndex=1, outcome=outcome, confirmed=True)
    if outcome == "applied":
        fields["checkpoint"] = e.evidence(1, True)
    e.accept("actions/reconcile", **fields)
    assert e.state["nextActionIndex"] == next_index
    assert e.state["inFlightActionIndex"] is None
    remaining = []
    while e.state["nextActionIndex"] < 3:
        index = e.state["nextActionIndex"]
        remaining.append(index)
        e.start(index)
        e.complete(index)
    assert remaining == ([2] if outcome == "applied" else [1, 2])
    result = dict(actionCount=3) if e.host == "wps" else dict(actionResults=[], verification={"status": "verified"})
    e.accept("complete", ownerToken=e.state["ownerToken"], result=result)
    # Reopening is read-only and cannot re-acquire completed execution.
    first = e.status()
    assert first == e.status()
    assert first["status"] == "completed"
    assert first["ownerToken"] is None
    assert e.post("resume", confirmed=True).status_code == 409


def test_a1_unknown_host_state_cannot_be_marked_complete(execution):
    e = execution
    e.start(0)
    e.suspend()
    before = e.status()
    assert e.post("actions/reconcile", actionIndex=0, outcome="unknown", confirmed=True).status_code == 422
    assert e.status() == before


def test_a1_concurrent_resume_has_one_owner_and_rejects_stale_receipt(execution):
    e = execution
    original_owner = e.state["ownerToken"]
    e.suspend()
    revision = e.state["revision"]
    barrier = Barrier(2)

    def claim():
        with TestClient(main.app) as contender:
            barrier.wait(timeout=10)
            return contender.post(e.base + "/resume", json=dict(expectedRevision=revision, confirmed=True))

    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = [pool.submit(claim) for _ in range(2)]
        responses = [future.result(timeout=20) for future in futures]
    assert sorted(r.status_code for r in responses) == [200, 409]
    winner = next(r.json() for r in responses if r.status_code == 200)
    assert winner["ownerToken"] != original_owner
    assert winner["revision"] == revision + 1
    e.state = winner
    e.start(0)
    # Supply current revision so rejection proves owner-token protection, not stale revision.
    before = e.status()
    rejected = e.post("actions/complete", ownerToken=original_owner, actionIndex=0, checkpoint=e.evidence(0, True))
    assert rejected.status_code == 409
    assert e.status() == before
    e.complete(0)
    assert e.state["nextActionIndex"] == 1

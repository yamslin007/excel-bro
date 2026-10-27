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
    value.db_path = tmp_path / "ledger.sqlite3"
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


def test_a1_killed_process_keeps_inflight_and_new_process_reads_same_state(execution, tmp_path, monkeypatch):
    """Real process termination; no suspend. Host application itself is not killed."""
    import json
    import subprocess
    import sys
    import time
    from pathlib import Path
    import server.app.task_ledger as ledger_module
    e = execution
    e.start(0)
    e.complete(0)
    request = dict(method="POST", path=e.base + "/actions/start", body=dict(
        expectedRevision=e.state["revision"], ownerToken=e.state["ownerToken"],
        actionIndex=1, beforeSnapshot=e.evidence(1)))
    ready = tmp_path / "committed.json"
    worker = Path(__file__).with_name("execution_process_worker.py")
    process = subprocess.Popen([sys.executable, str(worker), str(e.db_path), str(ready)],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    try:
        process.stdin.write(json.dumps(request) + "\n")
        process.stdin.flush()
        deadline = time.monotonic() + 30
        while not ready.exists():
            assert process.poll() is None, "worker exited before committing"
            assert time.monotonic() < deadline, "worker did not commit in time"
            time.sleep(.02)
        committed = json.loads(ready.read_text())
        assert committed["status"] == 200, committed
        assert committed["body"]["inFlightActionIndex"] == 1
        process.kill()
        process.wait(timeout=10)
        assert process.returncode != 0
    finally:
        if process.poll() is None:
            process.kill()
        process.communicate(timeout=10)
    # A fresh interpreter and ledger, not the killed process's in-memory state.
    reopened = subprocess.run([sys.executable, str(worker), str(e.db_path)],
        input=json.dumps(dict(method="GET", path=e.base))+"\n", capture_output=True,
        text=True, timeout=30, check=True)
    state = json.loads(reopened.stdout)["body"]
    assert (state["nextActionIndex"], state["inFlightActionIndex"]) == (1, 1)
    assert state["revision"] == committed["body"]["revision"]
    e.state = state
    # WPS owner leases survive process death. Simulate expiry explicitly, without
    # altering the persisted record or pretending a suspend occurred.
    if e.host == "wps":
        assert state["ownerActive"] is True
        assert e.post("actions/reconcile", actionIndex=1, outcome="applied", confirmed=True,
                      checkpoint=e.evidence(1, True)).status_code == 409
        now = time.time()
        lease = ledger_module.capability_int("taskRecovery", "wpsOwnerLeaseSeconds")
        monkeypatch.setattr(ledger_module.time, "time", lambda: now + lease + 1)
    e.accept("actions/reconcile", actionIndex=1, outcome="applied", confirmed=True,
             checkpoint=e.evidence(1, True))
    assert e.state["nextActionIndex"] == 2
    e.start(2)
    e.complete(2)
    e.accept("complete", ownerToken=e.state["ownerToken"],
             result=dict(actionCount=3) if e.host == "wps" else dict(actionResults=[]))
    assert e.status()["status"] == "completed"


def test_a1_separate_processes_compete_for_same_revision(execution):
    import json
    import subprocess
    import sys
    from pathlib import Path
    e = execution
    e.suspend()
    barrier = Barrier(2)
    worker = Path(__file__).with_name("execution_process_worker.py")
    request = json.dumps(dict(method="POST", path=e.base + "/resume",
        body=dict(expectedRevision=e.state["revision"], confirmed=True))) + "\n"
    def claim():
        barrier.wait(timeout=10)
        result = subprocess.run([sys.executable, str(worker), str(e.db_path)],
            input=request, capture_output=True, text=True, timeout=40, check=True)
        return json.loads(result.stdout)
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = [pool.submit(claim) for _ in range(2)]
        results = [f.result(timeout=50) for f in futures]
    assert sorted(r["status"] for r in results) == [200, 409]
    winner = next(r["body"] for r in results if r["status"] == 200)
    assert winner["revision"] == e.state["revision"] + 1
    assert e.status()["revision"] == winner["revision"]

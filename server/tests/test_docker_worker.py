"""Host tests never execute generated analysis; real execution is Docker-only."""
import io
import json
import subprocess
import sys
from pathlib import Path

from server.app import docker_worker as worker

WORKER = str(Path(worker.__file__))


def _run(payload):
    return subprocess.run([sys.executable, WORKER], input=json.dumps(payload),
                          capture_output=True, text=True, timeout=15)


def test_host_worker_fails_closed_without_executing_code():
    proc = _run({"code": "raise RuntimeError('MUST_NOT_EXECUTE')", "input": {}})
    assert json.loads(proc.stdout)["reasonCategory"] == "isolation_not_applied"
    assert "MUST_NOT_EXECUTE" not in proc.stdout + proc.stderr


def test_worker_requires_successful_kernel_gate(monkeypatch):
    monkeypatch.setattr(worker, "_read_request", lambda: {"code": "unused"})
    def fail():
        raise RuntimeError("PRIVATE_GATE_ERROR")
    monkeypatch.setattr(worker, "_lock_down", fail)
    def forbidden(*args):
        raise AssertionError("execution reached")
    monkeypatch.setattr(worker, "_run_code", forbidden)
    output = io.StringIO()
    monkeypatch.setattr(sys, "stdout", output)
    worker.main()
    assert json.loads(output.getvalue())["reasonCategory"] == "isolation_not_applied"
    assert "PRIVATE_GATE_ERROR" not in output.getvalue()


def test_worker_protocol_with_trusted_stub(monkeypatch):
    monkeypatch.setattr(worker, "_read_request", lambda: {"code": "unused"})
    monkeypatch.setattr(worker, "_lock_down", lambda: None)
    monkeypatch.setattr(worker, "_run_code", lambda *args: {"sum": 60})
    output = io.StringIO()
    monkeypatch.setattr(sys, "stdout", output)
    worker.main()
    assert json.loads(output.getvalue()) == {"status": "completed", "result": {"sum": 60}}


def test_worker_rejects_non_object_and_invalid_code():
    for payload in ([], {"notcode": 1}, {"code": 2}):
        assert json.loads(_run(payload).stdout)["reasonCategory"] == "invalid_request"


def test_worker_enforces_input_size_limit():
    assert json.loads(_run({"code": "unused", "input": "x" * (3 * 1024 * 1024)}).stdout)["reasonCategory"] == "invalid_request"

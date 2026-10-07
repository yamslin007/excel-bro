"""Real kernel enforcement, including process replacement (PID limits cannot do it)."""
import pytest

from server.tests.test_python_docker_real import DOCKER_AVAILABLE, _run_probe

pytestmark = pytest.mark.skipif(not DOCKER_AVAILABLE, reason="real local Docker image is unavailable")


def test_exec_replacement_is_denied_before_external_program_runs(monkeypatch):
    result, _ = _run_probe(monkeypatch,
        "import os\ntry: os.execv('/bin/sh',['sh','-c','printf EXTERNAL_PROGRAM_RAN'])\n"
        "except OSError as error: result={'errno':error.errno}\n", pids_limit=64)
    assert result.is_success(), result
    assert result.result_payload == {"status":"completed","result":{"errno":1}}


def test_fork_denied_even_with_spare_process_budget(monkeypatch):
    result, _ = _run_probe(monkeypatch,
        "import os\ntry:\n child=os.fork()\n"
        " if child==0: os._exit(0)\n"
        " os.waitpid(child,0);result={'forked':True}\n"
        "except OSError as error: result={'errno':error.errno}\n", pids_limit=64)
    assert result.is_success(), result
    assert result.result_payload["result"] == {"errno":1}


def test_native_execveat_and_x32_syscall_are_denied(monkeypatch):
    result, _ = _run_probe(monkeypatch,
        "import ctypes,errno\nlibc=ctypes.CDLL(None,use_errno=True)\nerrors=[]\n"
        "for number in (322,59|0x40000000,435):\n"
        " ctypes.set_errno(0);ret=libc.syscall(number,-1,None,None,None,0)\n"
        " errors.append([ret,ctypes.get_errno()])\n"
        "result=errors\n", pids_limit=64)
    assert result.is_success(), result
    assert result.result_payload["result"] == [[-1,1]] * 3


def test_network_socket_and_loopback_syscalls_are_denied(monkeypatch):
    result, _ = _run_probe(monkeypatch,
        "import socket\nerrors=[]\n"
        "for family in (socket.AF_INET,socket.AF_INET6,socket.AF_UNIX):\n"
        " try: socket.socket(family);errors.append('opened')\n"
        " except OSError as error: errors.append(error.errno)\n"
        "result=errors\n", pids_limit=64)
    assert result.is_success(), result
    assert result.result_payload["result"] == [1, 1, 1]


def test_registered_dataset_reaches_real_container_only(monkeypatch):
    from server.tests.test_python_datasets import prepared, ORIGIN
    from server.app.python_analysis import run_registered_analysis_internal
    registry, scope, token, receipt = prepared()
    result = run_registered_analysis_internal(data_registry=registry, scope_token=token,
        scope=scope, origin=ORIGIN, receipts=[receipt],
        code="rows=input['datasets'][0]['values'];result={'sum':sum(row[0] for row in rows),'count':len(rows)}")
    assert result == {"status": "completed", "result": {"sum": 3, "count": 2}}


def test_real_matrix_publication_returns_only_registered_result_receipt():
    from server.tests.test_python_datasets import prepared, ORIGIN
    from server.app.python_analysis import run_registered_result_internal
    from server.app.models import PythonResultReceipt
    registry,scope,token,data = prepared()
    response = run_registered_result_internal(data_registry=registry,scope_token=token,
        scope=scope,origin=ORIGIN,receipts=[data],
        code="rows=input['datasets'][0]['values'];result={'kind':'matrix','values':[[sum(row[0] for row in rows),len(rows)]]}")
    assert response["status"] == "completed", response
    assert set(response) == {"status","receipt"}
    receipt = PythonResultReceipt.model_validate(response["receipt"])
    assert registry.read_result(token,scope,ORIGIN,[data],receipt) == {"kind":"matrix","values":[[3,2]]}

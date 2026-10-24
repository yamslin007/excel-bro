"""Docker 隔离 Python 运行时适配器测试（mock 层，不依赖 Docker 真实运行）。

覆盖：三层分离、固定启动链参数、IPC 校验、限额、失败关闭、异常映射、
内部接口。真实 Docker 实测在 ``test_python_docker_real.py`` 单独标注（环境可用时跑）。

所有 mock 只替换 docker CLI 调用与子进程，不 mock 掉安全守卫本身——守卫仍真实执行。

第八批（Request I）接口更新：
- ``build_run_arguments`` 新增 ``cpu_seconds``（累计 CPU 时间）。
- ``verify_engine``/``image_present``/``_run_docker_cli`` 新增 ``endpoint``。
- ``image_ref()`` 在缺 digest 时返回不可被解析的哨兵（失败关闭），不再静默用裸 tag。
- ``run_python_isolated`` 用 ``Popen`` + reader 线程（不再 ``subprocess.run``）。
"""
from __future__ import annotations

import io
import json
import os
import subprocess
import threading

import pytest

from server.app import python_docker_runner as runner
from server.app import python_analysis


# 预先核实的本机端点替身（跳过端点核实本身，聚焦被测行为）。
_EP = runner.VerifiedEndpoint(scheme="npipe", host="npipe:////./pipe/docker_engine")


def _valid_config(**over):
    base = dict(image="img", image_digest="sha256:" + "0" * 64, non_root_user="65534:65534")
    base.update(over)
    return runner.DockerRuntimeConfig(**base)


def _popen_returning(stdout: str = "", stderr: str = "", returncode: int = 0,
                     hang: bool = False):
    """构造一个假 Popen：把 stdout/stderr 当作**有界管道**暴露。

    hang=True 时 poll() 持续返回 None（模拟长时间运行），供取消/超时路径使用。
    """
    class FakeProc:
        def __init__(self, *a, **k):
            self.returncode = None if hang else returncode
            self.stdin = io.BytesIO()
            self.stdout = io.BytesIO(stdout.encode("utf-8"))
            self.stderr = io.BytesIO(stderr.encode("utf-8"))

        def kill(self):
            self.returncode = -9

        def wait(self, timeout=None):
            if self.returncode is None:
                self.returncode = 0
            return self.returncode

        def poll(self):
            return self.returncode

    return FakeProc


def _fake_cli(name_object: str = "No such object"):
    """假 docker CLI：inspect 返回「不存在」，其余成功。"""
    def fake(args, timeout=30.0, **kwargs):
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 1, "", name_object)
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")
    return fake


# ---------------------------------------------------------------------------
# 1. 固定启动链参数（模型不能影响任何一项）
# ---------------------------------------------------------------------------


def test_build_run_arguments_is_fixed_and_isolated():
    config = runner.DockerRuntimeConfig(
        image="excel-bro-python-sandbox",
        image_digest="sha256:deadbeef",
    )
    args, env = runner.build_run_arguments(
        config,
        cpu_quota=1.0,
        cpu_seconds=60,
        memory_mb=1024,
        pids_limit=2,
        scratch_mb=256,
        container_name="excel-bro-py-testname",
    )
    joined = " ".join(args)
    assert "--network none" in joined
    assert "--read-only" in joined
    assert "--tmpfs" in joined and "size=256m" in joined
    assert "--user 65534:65534" in joined
    assert "--cap-drop ALL" in joined
    assert "--security-opt no-new-privileges" in joined
    assert "--log-driver none" in joined
    assert "--name excel-bro-py-testname" in joined
    assert "--memory 1024m" in joined
    assert "--memory-swap 1024m" in joined
    assert "--cpus 1.0" in joined
    assert "--pids-limit 2" in joined
    # 累计 CPU 时间限额（RLIMIT_CPU）已接线
    assert "--ulimit cpu=60:60" in joined
    assert "excel-bro-python-sandbox@sha256:deadbeef" in joined
    assert args[-2] == "python3"
    assert args[-1] == "/app/worker.py"
    assert "--mount" not in joined
    assert "-v" not in joined and "--volume" not in joined
    assert "/var/run/docker.sock" not in joined
    assert "API_KEY" not in env and "OPENAI" not in env and "ANTHROPIC" not in env


def test_image_ref_uses_digest_over_tag():
    config = runner.DockerRuntimeConfig(image="img", image_digest="sha256:abc")
    assert config.image_ref() == "img@sha256:abc"


def test_container_threads_are_fixed_not_inherited_from_host(monkeypatch):
    names = ("OPENBLAS_NUM_THREADS", "OMP_NUM_THREADS", "MKL_NUM_THREADS",
             "NUMEXPR_MAX_THREADS", "NUMEXPR_NUM_THREADS", "PYTHONUNBUFFERED",
             "PYTHONDONTWRITEBYTECODE")
    for name in names:
        monkeypatch.setenv(name, "999")
    args, _ = runner.build_run_arguments(
        _valid_config(), cpu_quota=1.0, cpu_seconds=60, memory_mb=1024,
        pids_limit=1, scratch_mb=256, container_name="excel-bro-py-env-test",
    )
    actual = [args[index + 1] for index, arg in enumerate(args) if arg == "--env"]
    assert set(actual) == {name + "=1" for name in names}
    assert args[args.index("--pids-limit") + 1] == "1"


def test_pump_delivers_short_output_before_eof():
    read_fd, write_fd = os.pipe()
    reader = io.BufferedReader(os.fdopen(read_fd, "rb", buffering=0))
    writer = os.fdopen(write_fd, "wb", buffering=0)
    collector = runner._OutputCollector(4096, 4096)
    delivered = threading.Event()
    feed = collector.feed_stderr

    def observe(chunk):
        feed(chunk)
        delivered.set()

    collector.feed_stderr = observe
    worker = threading.Thread(target=runner._pump, args=(reader, collector, False), daemon=True)
    try:
        writer.write(b"STARTED")
        worker.start()
        before_eof = delivered.wait(1)
    finally:
        writer.close()
        worker.join(2)
        reader.close()
    assert before_eof, "short output was buffered until EOF or a full 4096-byte read"
    assert collector.stderr_bytes() == b"STARTED"


def test_image_ref_without_digest_fails_closed():
    """缺 digest 不得静默降级为可变 tag：返回不可解析哨兵，并由校验拒绝。"""
    config = runner.DockerRuntimeConfig(image="img")
    assert config.image_ref() == "img@sha256:"
    assert config.is_immutable() is False
    ok, reason = runner.validate_isolation_config(config)
    assert ok is False and reason


# ---------------------------------------------------------------------------
# 2. 引擎验证与失败关闭
# ---------------------------------------------------------------------------


def test_verify_engine_rejects_non_linux(monkeypatch):
    def fake_cli(args, timeout=30.0, **kwargs):
        return subprocess.CompletedProcess(args, 0, "windows\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_cli)
    assert runner.verify_engine(_EP) is False


def test_verify_engine_accepts_linux(monkeypatch):
    def fake_cli(args, timeout=30.0, **kwargs):
        return subprocess.CompletedProcess(args, 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_cli)
    assert runner.verify_engine(_EP) is True


def test_verify_engine_returns_false_on_cli_error(monkeypatch):
    def fake_cli(args, timeout=30.0, **kwargs):
        return subprocess.CompletedProcess(args, 1, "", "error")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_cli)
    assert runner.verify_engine(_EP) is False


def test_verify_engine_returns_false_on_oserror(monkeypatch):
    def fake_cli(args, timeout=30.0, **kwargs):
        raise OSError("docker not found")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_cli)
    assert runner.verify_engine(_EP) is False


# ---------------------------------------------------------------------------
# 3. 执行路径：异常映射为固定有类型回执（不回显原始文本）
# ---------------------------------------------------------------------------


def _prepare_ready(monkeypatch):
    monkeypatch.setattr(runner, "verify_engine", lambda *a, **k: True)
    monkeypatch.setattr(runner, "image_present", lambda *a, **k: True)
    monkeypatch.setattr(runner, "_run_docker_cli", _fake_cli())


def test_run_python_isolated_fails_closed_when_engine_missing(monkeypatch):
    monkeypatch.setattr(runner, "verify_engine", lambda *a, **k: False)
    with pytest.raises(runner.DockerRuntimeError) as exc:
        runner.run_python_isolated(
            _valid_config(),
            code="print(1)", input_payload={},
            cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=10, endpoint=_EP,
        )
    assert exc.value.reason_category == "runner_missing"
    assert "print(1)" not in str(exc.value)


def test_run_python_isolated_fails_closed_when_image_missing(monkeypatch):
    monkeypatch.setattr(runner, "verify_engine", lambda *a, **k: True)
    monkeypatch.setattr(runner, "image_present", lambda *a, **k: False)
    with pytest.raises(runner.DockerRuntimeError) as exc:
        runner.run_python_isolated(
            _valid_config(),
            code="print(1)", input_payload={},
            cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=10, endpoint=_EP,
        )
    assert exc.value.reason_category == "image_missing"


def test_run_python_isolated_maps_timeout(monkeypatch):
    """超时 → 固定 timeout 状态，不回显任何内容。"""
    _prepare_ready(monkeypatch)
    monkeypatch.setattr(runner.subprocess, "Popen",
                        _popen_returning(hang=True))
    # 让循环快速判定超时
    monkeypatch.setattr(runner, "_new_container_name", lambda: "excel-bro-py-timeout")
    result = runner.run_python_isolated(
        _valid_config(),
        code="print(1)", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=0.1, endpoint=_EP,
    )
    assert result.status == "timeout"
    assert result.reason_category == "timeout"


def test_run_python_isolated_maps_worker_error_without_echoing(monkeypatch):
    """工作程序异常（非零退出码）→ 固定 error，stderr/traceback 不回显。"""
    _prepare_ready(monkeypatch)
    sentinel = "TRACEBACK_SENTINEL_secret"
    monkeypatch.setattr(runner.subprocess, "Popen",
                        _popen_returning(stderr=sentinel, returncode=1))
    result = runner.run_python_isolated(
        _valid_config(),
        code="print(1)", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10, endpoint=_EP,
    )
    assert result.status == "error"
    assert result.reason_category == "execution_error"
    assert sentinel not in json.dumps(result.result_payload or {})


def test_run_python_isolated_enforces_output_limit(monkeypatch):
    """输出超限 → limit_exceeded，不发布结果。"""
    _prepare_ready(monkeypatch)
    big = json.dumps({"status": "completed", "result": "x" * 5000})
    monkeypatch.setattr(runner.subprocess, "Popen", _popen_returning(stdout=big, returncode=0))
    result = runner.run_python_isolated(
        _valid_config(),
        code="print(1)", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=100, wall_timeout_seconds=10, endpoint=_EP,
    )
    assert result.status == "limit_exceeded"
    assert result.result_payload is None


def test_run_python_isolated_returns_typed_result(monkeypatch):
    """成功 → completed + 有类型结果。"""
    _prepare_ready(monkeypatch)
    result_json = json.dumps({"status": "completed", "result": {"sum": 4950}})
    monkeypatch.setattr(runner.subprocess, "Popen",
                        _popen_returning(stdout=result_json, returncode=0))
    result = runner.run_python_isolated(
        _valid_config(),
        code="result = sum(input)", input_payload={"input": []},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10, endpoint=_EP,
    )
    assert result.is_success()
    assert result.result_payload == {"status": "completed", "result": {"sum": 4950}}


# ---------------------------------------------------------------------------
# 4. 内部接口：失败关闭，不回退宿主执行
# ---------------------------------------------------------------------------


def test_internal_entry_returns_runner_missing_without_docker(monkeypatch):
    """隔离配置不完整时内部入口失败关闭，宿主执行 Python 次数为 0。

    显式注入缺少 ``imageDigest`` 的配置，不依赖开发机器是否已经构建过镜像。
    配置校验在镜像不可变身份这一项失败关闭——这是**正确**行为：
    没有经核实的不可变镜像就不执行。此测试证明：失败关闭发生时没有启动任何子进程。
    """
    calls = []
    original_run = runner.subprocess.run
    monkeypatch.setattr(runner, "_docker_config", lambda: runner.DockerRuntimeConfig(image="img"))

    def spy_run(args, **kwargs):
        calls.append(args)
        return original_run(args, **kwargs)

    monkeypatch.setattr(runner.subprocess, "run", spy_run)
    result = python_analysis.run_isolated_analysis_internal("print(1)", {})
    assert result["status"] == "unavailable"
    # 缺不可变镜像身份 → 失败关闭（不是 runner_missing，因为还没走到引擎检查）
    assert result["reasonCategory"] in ("image_not_immutable", "isolation_not_applied")
    assert calls == []


def test_internal_entry_maps_docker_error_to_fixed_receipt(monkeypatch):
    """协调器异常映射为固定有类型回执，不回显异常文本。

    用一份**完整合法**的配置注入（含不可变摘要）直达引擎/镜像检查：
    镜像缺失 → 固定 image_missing。
    """
    monkeypatch.setattr(runner, "_docker_config", lambda: _valid_config())
    monkeypatch.setattr(runner, "verified_endpoint", lambda: _EP)
    monkeypatch.setattr(runner, "verify_engine", lambda *a, **k: True)
    monkeypatch.setattr(runner, "image_present", lambda *a, **k: False)
    result = python_analysis.run_isolated_analysis_internal("print(1)", {})
    assert result["status"] == "unavailable"
    assert result["reasonCategory"] == "image_missing"


# ---------------------------------------------------------------------------
# 5. 产品入口仍保持 enabled=false（不因内部接口存在而开放）
# ---------------------------------------------------------------------------


def test_product_entry_still_disabled_by_default():
    from server.app.capabilities import capabilities
    capabilities.cache_clear()
    assert python_analysis.python_execution_enabled() is False
    receipt = python_analysis.execute_python_analysis(
        None, {"code": "print(1)", "datasetHandles": []}
    )
    assert receipt["status"] == "unavailable"
    assert receipt["reasonCategory"] == "disabled"


def test_execution_ack_requires_complete_first_frame():
    collector = runner._OutputCollector(4096, 4096)
    marker = runner.EXECUTION_STARTED_MARKER
    collector.feed_stderr(marker[:9])
    assert not collector.execution_started()
    collector.feed_stderr(marker[9:])
    assert collector.execution_started()
    other = runner._OutputCollector(4096, 4096)
    other.feed_stderr(b"error message: " + marker)
    assert not other.execution_started()

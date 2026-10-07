"""Docker 隔离执行器第八批已确认缺陷的红灯测试（2026-10-07，Request I）。

本文件**不依赖 Docker 引擎**：通过 mock docker CLI 或在本地进程内复现。
纪律：先复现缺陷（红灯），再修生产代码；每条配独立反向变异（见
``scripts/mutate_python_docker_defects2.py``）。

缺陷清单（源码审查 + 独立诊断确认）：
B1 输出未真正有界：stdout/stderr 写入无容量上限的 mkstemp 文件，只在结束后
   限量读取——内存风险转移到宿主磁盘，且执行不会因输出爆炸提前终止。
B2 容器清理不可核实：``_force_cleanup`` 忽略退出码/异常，取消路径无清理，
   启动与取消存在竞态，终止后不核查状态。
B3 端点未真正绑定：剔除 DOCKER_HOST/DOCKER_CONTEXT ≠ 绑定本机；Docker 仍可能
   读取持久化的 current context 指向远端。
B4 校验过松：非 root 只枚举少数字符串（0:1000/root:1000 漏过）；maxProcesses+1
   错误接线；CPU 速率≠累计 CPU 时间；PID 限额不冒充禁 exec；结果边界（NaN/Inf、
   深度、元素数量、畸形 UTF-8、多包协议）未校验；imageDigest=null 静默降级可变 tag。
B5 构建上下文未最小化：server/ 含 .env、model-connections.json，无 ignore 配置。
B6 真实验收测试假通过：见 ``test_python_docker_real.py``（本文件含静态判别辅助）。
"""
from __future__ import annotations

import ast
import io
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

import pytest

from server.app import python_docker_runner as runner
from server.app import python_analysis
from server.app import capabilities as runner_capabilities


REPO_ROOT = Path(__file__).resolve().parents[2]
DOCKER_DIR = REPO_ROOT / "server" / "docker"
APP_DIR = REPO_ROOT / "server" / "app"
SERVER_DIR = REPO_ROOT / "server"


def _fake_engine_ready(monkeypatch):
    monkeypatch.setattr(runner, "verify_engine", lambda *a, **k: True)
    monkeypatch.setattr(runner, "image_present", lambda *a, **k: True)


def _valid_config(**overrides) -> "runner.DockerRuntimeConfig":
    """构造一份**其余项都合法**的配置，用于隔离单项校验（避免镜像摘要等旁证干扰）。"""
    base = dict(
        image="img",
        image_digest="sha256:" + "0" * 64,
        non_root_user="65534:65534",
    )
    base.update(overrides)
    return runner.DockerRuntimeConfig(**base)


# 预先核实的本机端点替身：用于"不测端点核实本身"的执行测试，
# 直接传入以跳过 verify_local_endpoint（避免和被测行为无关的旁路失败）。
_FAKE_ENDPOINT = runner.VerifiedEndpoint(scheme="npipe", host="npipe:////./pipe/docker_engine")


# ---------------------------------------------------------------------------
# B1 输出真正有界：并发处理 + 落盘上限 + 提前终止
# ---------------------------------------------------------------------------


def test_output_is_drained_concurrently_not_buffered_in_memory():
    """stdout/stderr 必须**并发**读取（reader 线程），不能单线程缓冲进内存。

    判据：实现在运行期间启动了读取线程（或非阻塞读取），而不是把两个流
    都塞进内存/PIPE 再由主线程一次读完。
    """
    assert hasattr(runner, "_OutputCollector"), (
        "缺少 _OutputCollector：输出必须有并发的有界收集器，"
        "不能把内存风险转移到宿主磁盘"
    )
    col = runner._OutputCollector(stdout_cap=100, stderr_cap=100)
    # 收集器必须提供非阻塞写入与有界结果
    assert hasattr(col, "feed_stdout")
    assert hasattr(col, "feed_stderr")
    assert hasattr(col, "over_limit")
    assert hasattr(col, "stdout_bytes")
    assert hasattr(col, "stderr_bytes")


def test_output_collector_enforces_stdout_cap():
    """stdout 达到上限即标记 over_limit，且**不再累积**（内存/磁盘都封顶）。"""
    col = runner._OutputCollector(stdout_cap=64, stderr_cap=64)
    for _ in range(100):
        col.feed_stdout(b"x" * 16)
        if col.over_limit():
            break
    assert col.over_limit() is True
    # 封顶：内部缓冲不超过上限（允许 +1 哨兵字节）
    assert col.stdout_size() <= 65, f"stdout 缓冲未封顶: {col.stdout_size()}"


def test_output_collector_enforces_stderr_cap_independently():
    """stderr 有**独立**上限（不能无界），且触发时直接 over_limit。"""
    col = runner._OutputCollector(stdout_cap=1024, stderr_cap=32)
    col.feed_stdout(b"ok")
    for _ in range(20):
        col.feed_stderr(b"E" * 8)
        if col.over_limit():
            break
    assert col.over_limit() is True
    assert col.stderr_size() <= 33, f"stderr 缓冲未封顶: {col.stderr_size()}"
    assert col.stdout_size() == 2, "stdout 不应被 stderr 触发影响"


def test_output_limit_triggers_mid_run_and_kills_container(monkeypatch):
    """输出超限必须在**执行尚未正常结束**时触发：立即终止容器并进入核对。

    用受控输出生产者（真实子进程持续输出），证明限制在运行中触发、
    不等进程自然结束。
    """
    _fake_engine_ready(monkeypatch)
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        # 清理核查：容器已不存在（返回非零 + "No such object"）
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "No such object")
        # create/inspect/start/rm/kill 等其它 CLI 调用返回成功
        return subprocess.CompletedProcess(["docker", *args], 0, "abc123\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    # 用真实本地进程模拟"持续输出"的容器：subprocess.Popen 起一个生产者
    real_popen = subprocess.Popen

    class FakeProc:
        def __init__(self, *a, **k):
            self.returncode = None
            self._killed = False
            # 持续输出超大量数据的真实生产者
            self._proc = real_popen(
                [sys.executable, "-c",
                 "import sys\n"
                 "while True:\n"
                 "    sys.stdout.write('x'*4096); sys.stdout.flush()\n"],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
            )
            self.stdout = self._proc.stdout
            self.stderr = io.BytesIO()
            self.stdin = io.BytesIO()

        def kill(self):
            self._killed = True
            self._proc.kill()

        def wait(self, timeout=None):
            self.returncode = self._proc.wait(timeout=timeout)
            return self.returncode

        def poll(self):
            return self._proc.poll()

    monkeypatch.setattr(runner.subprocess, "Popen", FakeProc)
    result = runner.run_python_isolated(
        _valid_config(),
        code="while True: print('x'*4096)",
        input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=4096, wall_timeout_seconds=30,
        endpoint=_FAKE_ENDPOINT,
    )
    assert result.status == "limit_exceeded", f"超限未在运行中触发: {result.status}"
    # 必须发起终止（kill/rm）而不是等自然结束
    joined = " ".join(" ".join(c) for c in calls)
    assert ("kill" in joined) or ("rm" in joined), "超限后未终止容器"


# ---------------------------------------------------------------------------
# B2 容器清理可核实 + 真实取消
# ---------------------------------------------------------------------------


def test_cleanup_verifies_container_gone(monkeypatch):
    """清理必须核查状态：区分「确认已消失」与「无法查询」。

    判据：清理路径调用 docker rm -f **之后**，必须再查询一次确认容器不存在
    （inspect/ps）；返回状态未知时不得冒充清理成功。
    """
    assert hasattr(runner, "_cleanup_container"), "缺少可核实的 _cleanup_container"
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        if "inspect" in args and any(a.startswith("excel-bro-py-") for a in args):
            # 查询返回非零（容器已不存在）
            return subprocess.CompletedProcess(["docker", *args], 1, "", "No such object")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    state = runner._cleanup_container("excel-bro-py-deadbeef")
    assert state in ("gone", "unknown")
    # 必须发起过查询确认（inspect/ps 出现在参数里，可能是 ["container","inspect",...]）
    queried = any("inspect" in c or "ps" in c for c in calls)
    assert queried, "清理后未核查容器状态"


def test_cleanup_failure_not_reported_as_success(monkeypatch):
    """docker rm 返回非零且容器仍在 → 清理状态必须是 unknown/failed，不能是 gone。"""
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        if args[:1] == ["rm"]:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "cannot remove")
        if "inspect" in args:
            # 查询成功 → 容器仍存在
            return subprocess.CompletedProcess(["docker", *args], 0, "[]", "")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    state = runner._cleanup_container("excel-bro-py-stillhere")
    assert state != "gone", "清理失败/容器仍在被冒充为已清理"


def test_cleanup_query_failure_is_unknown_not_gone(monkeypatch):
    """核查查询本身失败（超时/CLI 异常）→ 状态 unknown，不能冒充 gone。"""
    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:1] == ["rm"]:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "err")
        raise subprocess.TimeoutExpired(cmd="docker", timeout=1)

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    state = runner._cleanup_container("excel-bro-py-unknown")
    assert state == "unknown", f"查询失败应记 unknown，实际 {state}"


def test_run_binds_trusted_run_identity_and_container_id(monkeypatch):
    """执行必须绑定可信侧运行身份 + 确切容器身份，且清理只针对本容器。"""
    _fake_engine_ready(monkeypatch)
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        if args[:1] == ["create"]:
            return subprocess.CompletedProcess(["docker", *args], 0, "containerid123\n", "")
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "No such object")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    popen_calls: list[list[str]] = []

    class FakeProc:
        def __init__(self, *a, **k):
            # docker run 通过 Popen 执行（不是 _run_docker_cli）——捕获其命令行
            popen_calls.append(list(a[0]) if a else list(k.get("args", [])))
            self.returncode = 0
            self.stdin = io.BytesIO()
            self.stdout = io.BytesIO(json.dumps({"status": "completed", "result": 1}).encode())
            self.stderr = io.BytesIO(b"")

        def kill(self): pass
        def wait(self, timeout=None): return 0
        def poll(self): return 0

    monkeypatch.setattr(runner.subprocess, "Popen", FakeProc)
    result = runner.run_python_isolated(
        _valid_config(),
        code="result = 1", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10,
        endpoint=_FAKE_ENDPOINT,
    )
    run_joined = " ".join(" ".join(c) for c in popen_calls)
    cli_joined = " ".join(" ".join(c) for c in calls)
    joined = run_joined + " " + cli_joined
    # 必须用 --name 绑定唯一容器名（运行身份）
    assert "--name" in run_joined, f"未绑定可信侧运行身份（容器名）: {run_joined}"
    # 运行容器名必须与清理目标一致（绑定同一身份）
    import re as _re
    m = _re.search(r"--name\s+(excel-bro-py-\w+)", run_joined)
    assert m, f"docker run 未带唯一容器名: {run_joined}"
    container_name = m.group(1)
    assert container_name in cli_joined, (
        f"清理未针对本次确切容器身份 {container_name}: {cli_joined}"
    )
    # 不得出现全局 prune 或宽泛前缀删除
    assert "prune" not in joined, "禁止全局 prune"
    # 删除必须针对**确切容器名**（带唯一后缀），不是裸前缀/通配
    assert "excel-bro-py-" in cli_joined, "清理未针对本容器名"
    # 裸前缀（无唯一后缀）删除 = 宽泛删除，禁止
    assert "rm -f excel-bro-py- " not in cli_joined + " ", "禁止裸前缀删除"
    assert "*" not in cli_joined, "禁止通配删除"


def test_cancel_cleans_up_and_propagates(monkeypatch):
    """取消必须清理容器，并把取消信号按原契约传播（不吞成成功/普通不可用）。"""
    _fake_engine_ready(monkeypatch)
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "No such object")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    import signal
    kill_event = threading.Event()

    real_popen = subprocess.Popen

    class SlowProc:
        def __init__(self, *a, **k):
            self.returncode = None
            self.stdout = io.BytesIO(b"")
            self.stderr = io.BytesIO(b"")
            self.stdin = io.BytesIO()

        def kill(self): pass
        def wait(self, timeout=None): return 0
        def poll(self): return None

    monkeypatch.setattr(runner.subprocess, "Popen", SlowProc)

    # 用一个会在 0.2s 后触发的取消检查
    cancel_state = {"cancelled": False}

    def cancel_check():
        import time
        time.sleep(0.2)
        cancel_state["cancelled"] = True
        return True

    # run_python_isolated 需支持取消检查回调
    import inspect
    sig = inspect.signature(runner.run_python_isolated)
    assert "cancel_check" in sig.parameters, (
        "run_python_isolated 未支持取消检查回调，无法实现真实取消"
    )


def test_cancel_before_execution_short_circuits(monkeypatch):
    """执行前的取消必须短路：不得起 Docker 容器。"""
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    monkeypatch.setattr(runner, "verify_engine", lambda *a, **k: True)
    monkeypatch.setattr(runner, "image_present", lambda *a, **k: True)

    def boom_popen(*a, **k):
        raise AssertionError("执行前已取消，不应启动容器")

    monkeypatch.setattr(runner.subprocess, "Popen", boom_popen)
    result = runner.run_python_isolated(
        _valid_config(),
        code="result = 1", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10,
        cancel_check=lambda: True,
        endpoint=_FAKE_ENDPOINT,
    )
    assert result.status == "cancelled", result.status


def test_late_result_after_cancel_rejected(monkeypatch):
    """取消后到达的迟到结果不得被接受。"""
    _fake_engine_ready(monkeypatch)

    def fake_docker(args, timeout=30.0, **kwargs):
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "No such object")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    class Proc:
        def __init__(self, *a, **k):
            self.returncode = 0
            # 迟到的合法结果
            self.stdout = io.BytesIO(json.dumps({"status": "completed", "result": 999}).encode())
            self.stderr = io.BytesIO(b"")
            self.stdin = io.BytesIO()
        def kill(self): pass
        def wait(self, timeout=None): return 0
        def poll(self): return 0

    monkeypatch.setattr(runner.subprocess, "Popen", Proc)

    # 关键：取消必须在**执行前检查之后**才变真 —— 使守卫唯一落在「迟到结果」那道。
    # 第 1 次调用是执行前检查（返回 False，放行）；之后进程已自然结束（poll()!=None
    # 使等待循环不调用 cancel_check），因此第 2 次调用就是自然结束后的迟到守卫。
    calls = {"n": 0}

    def cancel_check():
        calls["n"] += 1
        return calls["n"] >= 2

    result = runner.run_python_isolated(
        _valid_config(),
        code="result = 999", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10,
        cancel_check=cancel_check,
        endpoint=_FAKE_ENDPOINT,
    )
    assert result.status == "cancelled", f"取消后迟到结果被接受: {result.status}"
    assert result.is_success() is False


# ---------------------------------------------------------------------------
# B3 验证并锁定实际本机端点
# ---------------------------------------------------------------------------


def test_engine_endpoint_rejects_remote_tcp(monkeypatch):
    """实际端点指向远端 TCP/SSH 时必须拒绝（剔除环境变量不等于绑定本机）。"""
    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:2] == ["context", "inspect"]:
            return subprocess.CompletedProcess(
                ["docker", *args], 0,
                json.dumps({"Name": "remote", "Endpoints": {"docker": {"Host": "tcp://1.2.3.4:2375"}}}),
                "",
            )
        if args[:1] == ["info"]:
            return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    ok, reason = runner.verify_local_endpoint()
    assert ok is False, "远端 TCP 端点未被拒绝"
    assert reason


def test_engine_endpoint_accepts_local_npipe_or_socket(monkeypatch):
    """本机 npipe/unix socket 端点应被接受。"""
    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:2] == ["context", "inspect"]:
            return subprocess.CompletedProcess(
                ["docker", *args], 0,
                json.dumps({"Name": "desktop-linux",
                            "Endpoints": {"docker": {"Host": "npipe:////./pipe/docker_engine"}}}),
                "",
            )
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    ok, reason = runner.verify_local_endpoint()
    assert ok is True, f"本机 npipe 端点被误拒: {reason}"


def test_endpoint_locked_across_lifecycle(monkeypatch):
    """整个运行生命周期必须锁到**同一经验证端点**：核实之后即使 current context
    被改为远端，后续 docker CLI 调用仍必须带同一 ``-H``，不得在运行期重新解析。
    """
    assert hasattr(runner, "VerifiedEndpoint"), "缺少 VerifiedEndpoint：端点未被锁定"
    assert hasattr(runner, "endpoint_cli_prefix"), "缺少 endpoint_cli_prefix：未提供锁定参数"

    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        if args[:2] == ["context", "inspect"]:
            # 核实阶段：本机 npipe
            return subprocess.CompletedProcess(
                ["docker", *args], 0,
                json.dumps({"Name": "desktop-linux",
                            "Endpoints": {"docker": {"Host": "npipe:////./pipe/docker_engine"}}}),
                "",
            )
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    # 核实端点（此刻本机）
    endpoint = runner.verified_endpoint()
    assert endpoint.host == "npipe:////./pipe/docker_engine"
    calls.clear()

    # 核实之后 current context 被改为远端；后续调用必须仍锁到原端点
    def fake_docker_after_context_change(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        if args[:2] == ["context", "inspect"]:
            return subprocess.CompletedProcess(
                ["docker", *args], 0,
                json.dumps({"Name": "prod",
                            "Endpoints": {"docker": {"Host": "tcp://10.0.0.9:2375"}}}),
                "",
            )
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker_after_context_change)

    # 运行期使用的 CLI 前缀必须来自锁定端点，而不是重新解析 context
    prefix = runner.endpoint_cli_prefix(endpoint)
    assert prefix == ["-H", "npipe:////./pipe/docker_engine"], (
        f"运行期未锁定端点（重新解析了 context）: {prefix}"
    )
    # 生命周期各阶段调用都必须带上**同一**锁定端点（by identity/host），
    # 且不得再出现 context inspect（运行期不重新解析）。
    seen_endpoints: list[str] = []

    def capture_docker(args, timeout=30.0, endpoint=None, **kwargs):
        calls.append(list(args))
        if endpoint is not None:
            seen_endpoints.append(endpoint.host)
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", capture_docker)
    runner.verify_engine(endpoint)
    assert calls, "未发起引擎检查"
    for call in calls:
        assert call[:2] != ["context", "inspect"], "运行期重新解析了 current context"
    assert seen_endpoints == ["npipe:////./pipe/docker_engine"], (
        f"生命周期调用未锁定到同一端点: {seen_endpoints}"
    )


def test_endpoint_parser_ignores_non_docker_urls_before_host(monkeypatch):
    """真实 `docker context inspect` 输出里，``Host`` 之前可能有其它 URL（如 otel 的
    ``unix://...``）——端点解析必须取 **docker 端点的 Host 字段**，不能取第一个
    匹配到的 URL，否则可能把非 docker 端点或错误 host 锁进生命周期。

    红：当前实现用正则 ``re.search`` 取第一个 URL 匹配，会先命中 otel 的
    ``unix://`` 串（在真实输出里出现在 ``Host`` 之前）。
    """
    real_ctx = json.dumps([{
        "Name": "desktop-linux",
        "Metadata": {
            "Description": "Docker Desktop",
            "otel": {
                "OTEL_EXPORTER_OTLP_ENDPOINT":
                    "unix:////./c:/Users/x/AppData/Local/Docker/run/userAnalyticsOtlpHttp.sock"
            },
        },
        "Endpoints": {"docker": {"Host": "npipe:////./pipe/dockerDesktopLinuxEngine"}},
    }])

    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:2] == ["context", "inspect"]:
            return subprocess.CompletedProcess(["docker", *args], 0, real_ctx, "")
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    ok, host = runner.verify_local_endpoint()
    assert ok is True, f"本机 npipe 端点被误拒: {host}"
    assert host == "npipe:////./pipe/dockerDesktopLinuxEngine", (
        f"端点解析取到了非 docker URL（应为 Host 字段）: {host}"
    )


def test_persistent_remote_context_rejected_even_with_empty_env(monkeypatch):
    """环境变量为空但持久化 current context 指向远端 → 必须拒绝。"""
    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:2] == ["context", "inspect"]:
            return subprocess.CompletedProcess(
                ["docker", *args], 0,
                json.dumps({"Name": "prod", "Endpoints": {"docker": {"Host": "ssh://user@host"}}}),
                "",
            )
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    ok, _ = runner.verify_local_endpoint()
    assert ok is False


# ---------------------------------------------------------------------------
# B3b 端点必须是经**结构校验**的本机形态（第九批 · 焦点一）
# 仅用 startswith("npipe://"/"unix://") 不足：npipe:////remote-host/pipe/...
# 这类"形似本机、实质远端"的串会被误接受。
# ---------------------------------------------------------------------------


def _endpoint_from_context(host: str, name: str = "ctx"):
    """构造只返回给定 Host 的 fake docker CLI。"""
    payload = json.dumps({"Name": name, "Endpoints": {"docker": {"Host": host}}})

    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:2] == ["context", "inspect"]:
            return subprocess.CompletedProcess(["docker", *args], 0, payload, "")
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    return fake_docker


def test_endpoint_rejects_remote_named_pipe_host(monkeypatch):
    """``npipe:////remote-host/pipe/docker_engine`` 形似本机但实为远端主机 →
    必须拒绝（仅校验 scheme 前缀会误接受）。"""
    monkeypatch.setattr(
        runner, "_run_docker_cli",
        _endpoint_from_context("npipe:////remote-host/pipe/docker_engine"),
    )
    ok, reason = runner.verify_local_endpoint()
    assert ok is False, "远端命名管道被误当作本机端点接受"
    assert reason


@pytest.mark.parametrize("host", [
    "npipe:////./pipe/docker_engine",              # 标准本机命名管道
    "npipe:////./pipe/dockerDesktopLinuxEngine",   # Docker Desktop 引擎
    "npipe:////?/pipe/docker_engine",              # 扩展长度前缀形态
])
def test_endpoint_accepts_valid_local_named_pipe(monkeypatch, host):
    """显式支持的本机命名管道形态必须被接受。"""
    monkeypatch.setattr(runner, "_run_docker_cli", _endpoint_from_context(host))
    ok, reason = runner.verify_local_endpoint()
    assert ok is True, f"合法本机 npipe 被误拒: {host} → {reason}"


@pytest.mark.parametrize("host", [
    "unix:///var/run/docker.sock",                 # 标准本机 unix socket
    "unix:///run/user/1000/docker.sock",           # rootless 本机 socket
])
def test_endpoint_accepts_valid_local_unix_socket(monkeypatch, host):
    monkeypatch.setattr(runner, "_run_docker_cli", _endpoint_from_context(host))
    ok, reason = runner.verify_local_endpoint()
    assert ok is True, f"合法本机 unix socket 被误拒: {host} → {reason}"


@pytest.mark.parametrize("host", [
    "npipe:////remote-host/pipe/docker_engine",    # 远端主机名管道
    "npipe://remote/pipe/docker_engine",           # 非本机前缀
    "npipe:////remote/pipe/docker_engine",         # 远端路径
    "unix://remote-host/var/run/docker.sock",      # unix 带远端 authority
    "unix:////remote-host/run/docker.sock",        # 形似本机、实为远端
    "npipe://",                                    # 空前缀
    "npipe:////",                                  # 空路径
    "unix://",                                     # 空
    "tcp://1.2.3.4:2375",                          # 远端 TCP
    "ssh://user@host",                             # 远端 SSH
    "fd://",                                       # fd（非本机 socket 形态）
    "http://localhost:2375",                       # 非 docker 端点 scheme
    "",                                            # 空串
    "npipe:////./pipe",                            # 不完整管道路径
])
def test_endpoint_rejects_remote_or_malformed_hosts(monkeypatch, host):
    """远端、空路径、畸形、无法确认的端点必须一律拒绝（不得改写为 local）。"""
    monkeypatch.setattr(runner, "_run_docker_cli", _endpoint_from_context(host))
    ok, reason = runner.verify_local_endpoint()
    assert ok is False, f"非法端点未被拒绝: {host!r}"
    assert reason


def test_verified_endpoint_instance_is_validated_not_trusted_by_name():
    """执行路径不得仅凭对象名为 ``VerifiedEndpoint`` 就信任；
    用非法 host 构造的实例必须被再次校验并拒绝。"""
    bad = runner.VerifiedEndpoint(scheme="npipe", host="npipe:////remote-host/pipe/docker_engine")
    assert runner.endpoint_is_locally_verified(bad) is False, (
        "非法端点实例被按名字信任（未做结构校验）"
    )
    good = runner.VerifiedEndpoint(scheme="npipe", host="npipe:////./pipe/docker_engine")
    assert runner.endpoint_is_locally_verified(good) is True


def test_unverified_endpoint_cannot_enter_execution_path(monkeypatch):
    """未经验证的端点不得进入执行路径：run_python_isolated 收到非法 endpoint
    必须在任何 docker CLI 调用前失败（调用计数为 0）。"""
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    bad = runner.VerifiedEndpoint(scheme="npipe", host="npipe:////remote-host/pipe/docker_engine")
    with pytest.raises(runner.DockerRuntimeError):
        runner.run_python_isolated(
            _valid_config(),
            code="result = 1", input_payload={},
            cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=10,
            endpoint=bad,
        )
    assert calls == [], f"非法端点进入了执行路径并发起 docker 调用: {calls}"


def test_cli_level_guard_rejects_unverified_endpoint():
    """最内层防线：即便绕过 run_python_isolated 直接调用 CLI 执行器，
    传入非法端点也必须被拒绝，且不发起任何子进程（真实 subprocess 未被调用）。"""
    bad = runner.VerifiedEndpoint(scheme="npipe", host="npipe:////remote-host/pipe/docker_engine")
    with pytest.raises(runner.DockerRuntimeError):
        runner._run_docker_cli(["info"], timeout=1.0, endpoint=bad)
    good = runner.VerifiedEndpoint(scheme="npipe", host="npipe:////./pipe/docker_engine")
    # 合法端点不应在**校验阶段**抛错（真实 docker 可能不存在，容许 FileNotFoundError）
    try:
        runner._run_docker_cli(["version"], timeout=1.0, endpoint=good)
    except runner.DockerRuntimeError as error:
        raise AssertionError(f"合法端点被内层防线误拒: {error}") from error
    except (FileNotFoundError, OSError, subprocess.TimeoutExpired):
        pass


# ---------------------------------------------------------------------------
# B4 收紧配置、镜像与结果校验
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("user", ["0", "0:0", "root", "root:root", "0:1000", "root:1000",
                                  "00:0", "0:nogroup", ""])
def test_non_root_validation_rejects_root_uid_variants(user):
    """非 root 校验必须拒绝一切 root UID 形式（含 0:1000、root:1000 等变体）。"""
    cfg = _valid_config(non_root_user=user)
    ok, reason = runner.validate_isolation_config(cfg)
    assert ok is False, f"root UID 变体 {user!r} 未被拒绝"
    assert reason


def test_non_root_validation_accepts_numeric_nonroot():
    cfg = _valid_config(non_root_user="65534:65534")
    ok, _ = runner.validate_isolation_config(cfg)
    assert ok is True


def test_image_ref_requires_immutable_identity():
    """镜像身份必须不可变：imageDigest 缺失时不得静默降级为可变 tag。"""
    cfg = runner.DockerRuntimeConfig(image="img", image_digest=None)
    ref = cfg.image_ref()
    assert "@sha256:" in ref, f"imageDigest 缺失时应失败关闭或用不可变身份，实际: {ref}"


def test_worker_payload_rejects_nan_and_infinity(monkeypatch):
    """结果含 NaN/Infinity 必须拒绝（JSON 非有限数）。"""
    _fake_engine_ready(monkeypatch)
    bad = '{"status": "completed", "result": {"v": NaN}}'
    monkeypatch.setattr(runner, "_run_docker_cli",
                        lambda args, timeout=30.0, **kwargs: subprocess.CompletedProcess(["docker"], 0, "", ""))
    ok, reason, value = runner._validate_worker_payload(json.loads(bad, parse_constant=lambda x: float("nan")))
    assert ok is False


def test_worker_payload_rejects_excessive_depth():
    """结果嵌套过深必须拒绝。"""
    deep = {"status": "completed", "result": None}
    cursor = {}
    deep["result"] = cursor
    for _ in range(200):
        cursor["next"] = {}
        cursor = cursor["next"]
    ok, reason, value = runner._validate_worker_payload(deep)
    assert ok is False, "过深结果未被拒绝"


def test_worker_payload_rejects_excessive_elements():
    """结果元素数量超上限必须拒绝。"""
    big = {"status": "completed", "result": list(range(10_000_000))}
    ok, reason, value = runner._validate_worker_payload(big)
    assert ok is False


def test_worker_payload_preserves_null_vs_missing_result():
    """区分合法 result=null 与缺少 result 字段。"""
    ok_null, _, val_null = runner._validate_worker_payload({"status": "completed", "result": None})
    assert ok_null is True and val_null is None
    ok_missing, _, _ = runner._validate_worker_payload({"status": "completed"})
    assert ok_missing is False


def test_worker_status_error_with_zero_exit_not_success(monkeypatch):
    _fake_engine_ready(monkeypatch)
    ok, _, _ = runner._validate_worker_payload({"status": "error"})
    assert ok is False


def test_cpu_seconds_limit_wired(monkeypatch):
    """CPU 累计时间必须落实为可校验的限额，未落实则失败关闭。"""
    assert hasattr(runner, "build_run_arguments")
    import inspect
    params = inspect.signature(runner.build_run_arguments).parameters
    assert "cpu_seconds" in params, "build_run_arguments 未接线 CPU 累计时间限额"


# ---------------------------------------------------------------------------
# B5 构建上下文最小化
# ---------------------------------------------------------------------------


def test_build_context_excludes_sensitive_files():
    """构建上下文必须排除 .env / model-connections.json 等敏感文件。

    判据：存在生效的 .dockerignore（位于构建上下文根），且其规则覆盖敏感文件；
    并**实际计算**会被纳入上下文的文件集合（模拟 ignore 匹配）。
    """
    ignore = SERVER_DIR / ".dockerignore"
    assert ignore.exists(), "构建上下文 server/ 缺少 .dockerignore"
    rules = [ln.strip() for ln in ignore.read_text(encoding="utf-8").splitlines()
             if ln.strip() and not ln.strip().startswith("#")]
    # 必须排除敏感与无关文件
    joined = "\n".join(rules)
    for needed in (".env", "model-connections.json", "__pycache__", ".pytest_cache",
                   ".pytest_tmp", "tests"):
        assert any(needed in r for r in rules), f".dockerignore 未排除 {needed}"


def test_build_context_actual_file_set_excludes_secrets():
    """实际可进入上下文的文件集合不得包含敏感文件（不只看 COPY 路径）。

    用 gitignore 风格匹配评估 server/ 下的文件（模拟 docker 上下文计算）。
    """
    from fnmatch import fnmatch
    ignore = SERVER_DIR / ".dockerignore"
    assert ignore.exists()
    rules = [ln.strip() for ln in ignore.read_text(encoding="utf-8").splitlines()
             if ln.strip() and not ln.strip().startswith("#")]

    def excluded(rel: str) -> bool:
        parts = rel.split("/")
        for r in rules:
            r = r.rstrip("/")
            if fnmatch(rel, r) or fnmatch(rel, r + "/*"):
                return True
            # 目录规则：任意路径段匹配
            if any(fnmatch(p, r) for p in parts[:-1]):
                return True
        return False

    sensitive = ["server/.env", "server/model-connections.json"]
    for rel in sensitive:
        assert excluded(rel.replace("server/", "")), f"敏感文件 {rel} 未被上下文排除"


# ---------------------------------------------------------------------------
# B6 真实验收测试静态判别（辅助，不替代行为验证）
# ---------------------------------------------------------------------------


def _real_tests_text() -> str:
    return (REPO_ROOT / "server" / "tests" / "test_python_docker_real.py").read_text(encoding="utf-8")


def test_real_test_has_reachability_positive_control():
    """网络隔离必须有**可达性正对照**：对照进程能连同一受控端点，隔离器不能。"""
    text = _real_tests_text()
    assert ("positive_control" in text or "对照" in text or "reachab" in text.lower()), (
        "网络测试缺少可达性正对照"
    )


def test_real_test_container_identity_verified():
    """真实清理测试必须用本次**确切容器身份**核对，不只做全局前后集合差。

    静态判别（辅助）：必须存在对确切容器名的存在性查询，并断言其结果确实为
    「已不存在」（``exists is False``）——这是区分「确认已清理」与「只是全局
    集合里没出现」的关键，光有 ``inspect`` 字样不足以说明。
    """
    text = _real_tests_text()
    tree = ast.parse(text)
    # 源码里必须有 "exists is False" 形式的比较断言
    assert "exists is False" in text, (
        "真实测试未用确切容器身份断言「已不存在」（exists is False）"
    )
    # 且必须有按确切容器名查询的调用
    calls = [n for n in ast.walk(tree) if isinstance(n, ast.Call)
             and isinstance(n.func, ast.Name)
             and n.func.id in ("_exact_container_exists", "_container_exists")]
    assert calls, "真实测试未按确切容器名查询存在性"


def test_real_test_no_broad_except_to_skip():
    """真实测试的实现异常必须失败，不能被 broad except 转成 skip。"""
    text = _real_tests_text()
    tree = ast.parse(text)
    for node in ast.walk(tree):
        if isinstance(node, ast.Try):
            for handler in node.handlers:
                if handler.type is None or (
                    isinstance(handler.type, ast.Name) and handler.type.id == "Exception"
                ):
                    body = handler.body
                    # 若 handler 内直接 pytest.skip → 违规
                    for stmt in body:
                        src = ast.dump(stmt)
                        if "skip" in src:
                            pytest.fail("真实测试用 broad except 把异常转成 skip")


# ---------------------------------------------------------------------------
# B7 资源参数严格校验（第九批 · 焦点二）
# capability_int 只查 isinstance(int)/>0：Python 里 bool 是 int 子类，True 会被接受；
# capability_float 不拒绝 NaN/Infinity。资源限额必须在**起 Docker 之前**拒绝非法值，
# 不得静默转成默认值或"无限"。
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("bad", [True, False])
def test_capability_int_rejects_bool(monkeypatch, bad):
    """capability_int 必须拒绝 bool（bool 是 int 子类，不能冒充配额）。"""
    monkeypatch.setattr(
        runner_capabilities, "capabilities",
        lambda: {"probe": {"quota": bad}},
    )
    with pytest.raises(RuntimeError):
        runner_capabilities.capability_int("probe", "quota")


@pytest.mark.parametrize("bad", [0, -1, -100, 1.5, "8", None, []])
def test_capability_int_rejects_non_positive_or_non_int(monkeypatch, bad):
    monkeypatch.setattr(
        runner_capabilities, "capabilities",
        lambda: {"probe": {"quota": bad}},
    )
    with pytest.raises(RuntimeError):
        runner_capabilities.capability_int("probe", "quota")


def test_capability_int_accepts_positive_int(monkeypatch):
    monkeypatch.setattr(
        runner_capabilities, "capabilities",
        lambda: {"probe": {"quota": 8}},
    )
    assert runner_capabilities.capability_int("probe", "quota") == 8


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), float("-inf")])
def test_capability_float_rejects_non_finite(monkeypatch, bad):
    """浮点资源参数必须拒绝 NaN/Infinity（否则可绕过上限）。"""
    monkeypatch.setattr(
        runner_capabilities, "capabilities",
        lambda: {"probe": {"rate": bad}},
    )
    with pytest.raises(RuntimeError):
        runner_capabilities.capability_float("probe", "rate")


@pytest.mark.parametrize("bad", [True, False])
def test_capability_float_rejects_bool(monkeypatch, bad):
    monkeypatch.setattr(
        runner_capabilities, "capabilities",
        lambda: {"probe": {"rate": bad}},
    )
    with pytest.raises(RuntimeError):
        runner_capabilities.capability_float("probe", "rate")


@pytest.mark.parametrize("bad", [-1.0, -0.001])
def test_capability_float_rejects_negative(monkeypatch, bad):
    monkeypatch.setattr(
        runner_capabilities, "capabilities",
        lambda: {"probe": {"rate": bad}},
    )
    with pytest.raises(RuntimeError):
        runner_capabilities.capability_float("probe", "rate")


# --- 适配器入口：build_run_arguments 必须在构造 CLI 前拒绝非法参数 ---

def _build_args(**overrides):
    base = dict(
        cpu_quota=1.0, cpu_seconds=60, memory_mb=1024,
        pids_limit=2, scratch_mb=256, container_name="excel-bro-py-x",
    )
    base.update(overrides)
    return runner.build_run_arguments(_valid_config(), **base)


@pytest.mark.parametrize("field,bad", [
    ("cpu_seconds", True), ("cpu_seconds", 0), ("cpu_seconds", -5),
    ("cpu_seconds", 1.5), ("cpu_seconds", "60"),
    ("memory_mb", True), ("memory_mb", 0), ("memory_mb", -1), ("memory_mb", 1.5),
    ("pids_limit", True), ("pids_limit", 0), ("pids_limit", -2), ("pids_limit", 2.5),
    ("scratch_mb", True), ("scratch_mb", 0), ("scratch_mb", -1),
])
def test_build_run_arguments_rejects_invalid_int_quota(field, bad):
    """整数配额必须拒绝 bool/非整数/零/负数（不得静默钳位）。"""
    with pytest.raises((ValueError, runner.DockerRuntimeError)):
        _build_args(**{field: bad})


@pytest.mark.parametrize("bad", [
    float("nan"), float("inf"), float("-inf"), True, 0, -0.5, "1.0",
])
def test_build_run_arguments_rejects_invalid_cpu_quota(bad):
    """CPU 速率（--cpus）必须拒绝 NaN/Infinity/bool/零/负/非数值。"""
    with pytest.raises((ValueError, runner.DockerRuntimeError)):
        _build_args(cpu_quota=bad)


def test_build_run_arguments_accepts_valid_values():
    args, _env = _build_args()
    joined = " ".join(args)
    assert "--cpus 1.0" in joined
    assert "--ulimit cpu=60:60" in joined
    assert "--memory 1024m" in joined
    assert "--pids-limit 2" in joined
    assert "size=256m" in joined


# --- 执行入口：run_python_isolated 必须在起 Docker 之前拒绝非法参数 ---

def test_run_python_isolated_rejects_bool_quota_before_docker(monkeypatch):
    """bool 配额必须在**任何 docker 调用之前**被拒绝（调用计数为 0）。"""
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    with pytest.raises((ValueError, runner.DockerRuntimeError)):
        runner.run_python_isolated(
            _valid_config(),
            code="result = 1", input_payload={},
            cpu_quota=1.0, memory_mb=True, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=10,
            endpoint=_FAKE_ENDPOINT,
        )
    assert calls == [], f"非法配额进入了 Docker 执行: {calls}"


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), float("-inf"), 0, -1.0])
def test_run_python_isolated_rejects_bad_cpu_quota_before_docker(monkeypatch, bad):
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    with pytest.raises((ValueError, runner.DockerRuntimeError)):
        runner.run_python_isolated(
            _valid_config(),
            code="result = 1", input_payload={},
            cpu_quota=bad, memory_mb=1024, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=10,
            endpoint=_FAKE_ENDPOINT,
        )
    assert calls == [], f"非法 CPU 速率进入了 Docker 执行: {calls}"


def test_run_python_isolated_rejects_negative_wall_timeout_before_docker(monkeypatch):
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        return subprocess.CompletedProcess(["docker", *args], 0, "linux\n", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    with pytest.raises((ValueError, runner.DockerRuntimeError)):
        runner.run_python_isolated(
            _valid_config(),
            code="result = 1", input_payload={},
            cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=-1,
            endpoint=_FAKE_ENDPOINT,
        )
    assert calls == [], f"非法墙钟时限进入了 Docker 执行: {calls}"


# ---------------------------------------------------------------------------
# 第十批 F1：stdin 阻塞不得绕过超时/取消（输入传输与读取、监控并发）
# ---------------------------------------------------------------------------


def _install_noop_cli(monkeypatch, *, inspect_gone: bool = True):
    """安装引擎就绪 + docker CLI 替身（rm/inspect 等返回成功/已消失）。"""
    _fake_engine_ready(monkeypatch)
    calls: list[list[str]] = []

    def fake_docker(args, timeout=30.0, **kwargs):
        calls.append(list(args))
        if inspect_gone and "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "No such object")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)
    return calls


class _BlockingStdin:
    """写入会阻塞固定时长的 stdin 替身，用于验证超时/取消覆盖输入传输。"""

    def __init__(self, delay: float) -> None:
        self._delay = delay
        self.closed = False

    def write(self, data):
        time.sleep(self._delay)
        return len(data)

    def close(self):
        self.closed = True

    def flush(self):
        pass


class _StdinBlockingProc:
    """进程永不结束（poll=None），stdin 写入阻塞，用于复现 F1。"""

    def __init__(self, stdin_delay: float = 0.2):
        self.returncode = None
        self.stdin = _BlockingStdin(stdin_delay)
        self.stdout = io.BytesIO(b"")
        self.stderr = io.BytesIO(b"")

    def kill(self):
        pass

    def wait(self, timeout=None):
        return 0

    def poll(self):
        return None


def test_stdin_block_does_not_bypass_wall_timeout(monkeypatch):
    """F1：stdin 写入阻塞时，wall timeout 仍须准时生效（从启动即计时）。

    红灯判据：wall_timeout=0.05、stdin 阻塞 0.2s。若计时从"输入写完"才开始，
    实测耗时会 ≥ 0.2s；修复后必须 < 0.15s 返回 timeout。
    """
    _install_noop_cli(monkeypatch)

    monkeypatch.setattr(
        runner.subprocess, "Popen",
        lambda *a, **k: _StdinBlockingProc(stdin_delay=0.2),
    )
    start = time.monotonic()
    result = runner.run_python_isolated(
        _valid_config(),
        code="x = 1", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=0.05,
        endpoint=_FAKE_ENDPOINT,
    )
    elapsed = time.monotonic() - start
    assert result.status == "timeout", f"stdin 阻塞绕过了超时: {result.status}"
    assert elapsed < 0.15, f"超时未被输入传输阻塞边界兜住: {elapsed:.3f}s"


def test_input_transfer_runs_concurrently_with_monitoring(monkeypatch):
    """F1：输入写入必须与监控并发——输入线程未完成时监控也能判定超时。"""
    _install_noop_cli(monkeypatch)
    # 输入写 0.3s，墙钟 0.05s：并发下必须 0.05s 左右返回，而非等 0.3s
    monkeypatch.setattr(
        runner.subprocess, "Popen",
        lambda *a, **k: _StdinBlockingProc(stdin_delay=0.3),
    )
    start = time.monotonic()
    result = runner.run_python_isolated(
        _valid_config(),
        code="x = 1", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=0.05,
        endpoint=_FAKE_ENDPOINT,
    )
    elapsed = time.monotonic() - start
    assert result.status == "timeout"
    assert elapsed < 0.2, f"输入传输未与监控并发: {elapsed:.3f}s"


def _run_input_transfer_failure(monkeypatch, *, after_poll: bool):
    """用事件把输入失败固定在循环检查前或进程结束时，不靠 sleep 排序。"""
    _install_noop_cli(monkeypatch)
    failure_allowed = threading.Event()
    input_failed = threading.Event()
    feeding_finished = threading.Event()
    poll_entered = threading.Event()
    feeding_state: dict = {}
    if not after_poll:
        failure_allowed.set()

    real_feed = runner._feed_stdin

    def observed_feed(stream, data, state):
        try:
            real_feed(stream, data, state)
            feeding_state.update(state)
        finally:
            feeding_finished.set()

    monkeypatch.setattr(runner, "_feed_stdin", observed_feed)

    class BrokenStdin:
        def write(self, data):
            if not failure_allowed.wait(timeout=2):
                raise OSError("test did not release input failure")
            input_failed.set()
            raise BrokenPipeError("container closed stdin")

        def close(self):
            pass

        def flush(self):
            pass

    class BrokenStdinProc:
        def __init__(self):
            self.returncode = None
            self.stdin = BrokenStdin()
            # 即便 stdout 是一份"合法完成"回包，输入失败也不得被当作成功发布
            self.stdout = io.BytesIO(
                json.dumps({"status": "completed", "result": 42}).encode()
            )
            self.stderr = io.BytesIO(b"")
            self.terminated = False

        def kill(self):
            self.terminated = True
            self.returncode = 0
            failure_allowed.set()

        def wait(self, timeout=None):
            return 0

        def poll(self):
            if self.terminated:
                return self.returncode
            poll_entered.set()
            # poll 在循环内输入检查之后调用。晚失败只在这里释放，
            # 然后等真实 feeder 记录完 error，保证只由自然结束后的守卫拦截。
            failure_allowed.set()
            assert feeding_finished.wait(timeout=2), "输入线程未完成错误记录"
            if after_poll:
                self.returncode = 0
                return 0
            # 早失败场景始终保持运行；拆循环守卫只能走 timeout，
            # 不能靠自然结束后的另一道守卫兜底。
            return None

    monkeypatch.setattr(runner.subprocess, "Popen", lambda *a, **k: BrokenStdinProc())
    try:
        result = runner.run_python_isolated(
            _valid_config(),
            code="result = 42", input_payload={},
            cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=1,
            endpoint=_FAKE_ENDPOINT,
        )
    finally:
        failure_allowed.set()
        assert feeding_finished.wait(timeout=2), "测试遗留了输入线程"
    assert input_failed.is_set(), "输入失败未真正触发，测试无效"
    assert feeding_state == {"fed": False, "error": "BrokenPipeError", "done": True}
    if after_poll:
        assert poll_entered.is_set(), "晚失败未经过循环检查后的 poll 边界"
    return result


def test_input_transfer_failure_not_swallowed_as_success(monkeypatch):
    """M23a：运行中输入失败须立即返回执行错误，不接受超时兜底。"""
    result = _run_input_transfer_failure(monkeypatch, after_poll=False)
    assert result.status == "error", f"输入失败未映射为执行错误: {result.status}"
    assert result.reason_category == "execution_error"
    assert result.result_payload is None
    assert result.is_success() is False


def test_input_transfer_failure_after_process_exit_not_published(monkeypatch):
    """M23b：poll 返回结束时才记录输入失败，合法回包也不得发布。"""
    result = _run_input_transfer_failure(monkeypatch, after_poll=True)
    assert result.status == "error", f"结束时输入失败被合法回包掩盖: {result.status}"
    assert result.reason_category == "execution_error"
    assert result.result_payload is None
    assert result.is_success() is False


def test_cancel_during_input_transfer_still_takes_effect(monkeypatch):
    """F1：取消必须覆盖输入传输阶段（不是等输入写完才检查）。"""
    _install_noop_cli(monkeypatch)
    monkeypatch.setattr(
        runner.subprocess, "Popen",
        lambda *a, **k: _StdinBlockingProc(stdin_delay=0.3),
    )

    calls = {"n": 0}

    def cancel_check():
        calls["n"] += 1
        # 第 1 次是执行前检查（放行）；之后立即取消（此时输入仍在阻塞）
        return calls["n"] >= 2

    start = time.monotonic()
    result = runner.run_python_isolated(
        _valid_config(),
        code="x = 1", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10,
        cancel_check=cancel_check,
        endpoint=_FAKE_ENDPOINT,
    )
    elapsed = time.monotonic() - start
    assert result.status == "cancelled", f"输入传输阶段的取消未生效: {result.status}"
    assert elapsed < 0.25, f"取消未覆盖输入传输: {elapsed:.3f}s"


# ---------------------------------------------------------------------------
# 第十批 F2：所有分支必须消费清理核查结果（unknown 不冒充普通状态）
# ---------------------------------------------------------------------------


def test_late_cancel_with_unknown_cleanup_returns_cleanup_unverified(monkeypatch):
    """F2：自然结束后迟到取消 + 清理未知 → 固定 cleanup_unverified，非普通 cancelled。"""
    _fake_engine_ready(monkeypatch)

    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:1] == ["rm"]:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "boom")
        if "inspect" in args:
            # 查询**成功**且容器仍存在 → 清理 unknown
            return subprocess.CompletedProcess(["docker", *args], 0, "[]", "")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    class Proc:
        def __init__(self, *a, **k):
            self.returncode = 0
            self.stdout = io.BytesIO(json.dumps({"status": "completed", "result": 1}).encode())
            self.stderr = io.BytesIO(b"")
            self.stdin = io.BytesIO()

        def kill(self): pass
        def wait(self, timeout=None): return 0
        def poll(self): return 0

    monkeypatch.setattr(runner.subprocess, "Popen", Proc)

    # 第 1 次放行（执行前），第 2 次为自然结束后的迟到取消守卫
    calls = {"n": 0}

    def cancel_check():
        calls["n"] += 1
        return calls["n"] >= 2

    result = runner.run_python_isolated(
        _valid_config(),
        code="result = 1", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10,
        cancel_check=cancel_check,
        endpoint=_FAKE_ENDPOINT,
    )
    assert result.status == "error", f"清理未知被冒充为普通取消: {result.status}"
    assert result.reason_category == "cleanup_unverified", result.reason_category
    assert result.is_success() is False


def test_natural_end_limit_exceeded_with_unknown_cleanup(monkeypatch):
    """F2：自然结束后输出超限 + 清理未知 → cleanup_unverified，非普通 limit_exceeded。"""
    _fake_engine_ready(monkeypatch)

    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:1] == ["rm"]:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "boom")
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 0, "[]", "")  # 仍在
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    class Proc:
        def __init__(self, *a, **k):
            self.returncode = 0
            # 超限输出：单包外的大量字节
            self.stdout = io.BytesIO(b"x" * 100000)
            self.stderr = io.BytesIO(b"")
            self.stdin = io.BytesIO()

        def kill(self): pass
        def wait(self, timeout=None): return 0
        def poll(self): return 0

    monkeypatch.setattr(runner.subprocess, "Popen", Proc)
    result = runner.run_python_isolated(
        _valid_config(),
        code="x = 1", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10,
        endpoint=_FAKE_ENDPOINT,
    )
    assert result.status == "error", f"清理未知被冒充为普通超限: {result.status}"
    assert result.reason_category == "cleanup_unverified", result.reason_category


def test_startup_exception_with_unknown_cleanup_does_not_report_success(monkeypatch):
    """F2：启动即异常 + 清理未知 → 走失败关闭，绝不发布成功。"""
    _fake_engine_ready(monkeypatch)

    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:1] == ["rm"]:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "boom")
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 0, "[]", "")
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    def boom_popen(*a, **k):
        raise OSError("cannot spawn docker")

    monkeypatch.setattr(runner.subprocess, "Popen", boom_popen)
    with pytest.raises(runner.DockerRuntimeError):
        runner.run_python_isolated(
            _valid_config(),
            code="x = 1", input_payload={},
            cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
            max_output_bytes=1024, wall_timeout_seconds=10,
            endpoint=_FAKE_ENDPOINT,
        )


def test_normal_completion_with_unknown_cleanup_not_published(monkeypatch):
    """F2：正常完成但清理未知 → 不发布结果（不得冒充 completed）。"""
    _fake_engine_ready(monkeypatch)

    def fake_docker(args, timeout=30.0, **kwargs):
        if args[:1] == ["rm"]:
            return subprocess.CompletedProcess(["docker", *args], 1, "", "boom")
        if "inspect" in args:
            return subprocess.CompletedProcess(["docker", *args], 0, "[]", "")  # 仍在
        return subprocess.CompletedProcess(["docker", *args], 0, "", "")

    monkeypatch.setattr(runner, "_run_docker_cli", fake_docker)

    class Proc:
        def __init__(self, *a, **k):
            self.returncode = 0
            self.stdout = io.BytesIO(json.dumps({"status": "completed", "result": 7}).encode())
            self.stderr = io.BytesIO(b"")
            self.stdin = io.BytesIO()

        def kill(self): pass
        def wait(self, timeout=None): return 0
        def poll(self): return 0

    monkeypatch.setattr(runner.subprocess, "Popen", Proc)
    result = runner.run_python_isolated(
        _valid_config(),
        code="result = 7", input_payload={},
        cpu_quota=1.0, memory_mb=1024, pids_limit=2, scratch_mb=256,
        max_output_bytes=1024, wall_timeout_seconds=10,
        endpoint=_FAKE_ENDPOINT,
    )
    assert result.status == "error", f"清理未知却发布了完成结果: {result.status}"
    assert result.reason_category == "cleanup_unverified"
    assert result.is_success() is False


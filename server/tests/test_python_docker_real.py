"""真实 Docker 隔离实测（环境可用时运行；不可用时 skip 并记录环境阻断）。

这些测试必须在真实 Docker Linux 引擎上运行，不能 mock。每项先证明测试条件
**真实存在**或代码确实到达访问尝试，避免恒真的假通过；实现异常一律失败，
**不得**被 broad except 转成 skip。

关键纪律（第八批缺陷 B6 修复）：
- 网络隔离必须有**可达性正对照**：受控端点同时被对照进程探到（证明端点真的在监听），
  隔离容器探不到；不能只拿"Python 能跑"当对照，也不能用可能已关闭的端口证明隔离。
- 宿主文件哨兵必须**真实存在**，并区分"容器自身回环"与"宿主服务"；
  Windows 路径在 Linux 里天然不存在，单凭读失败**不能**证明隔离。
- 宿主凭据：先注入**合成哨兵**，再证明未继承（不读真实密钥）。
- 内存/CPU/输出/临时盘/进程限额都必须**证明探测真的启动过**并能识别限额证据；
  不可用/配置错误/镜像缺失**不算**隔离成功。
- 停止/清理必须用**本次确切容器身份**（container_id/inspect）核对，
  不能只做全局前后集合差；容器查询失败不得返回空列表冒充"无残留"。
- 静态搜索或 AST 检查只能辅助，不能替代真实行为验证。

注意：本文件在 Docker 引擎不可用时整体 skip（不是"通过"），由交付报告单独
记录环境阻断，不与 mock 测试混淆。
"""
from __future__ import annotations

import json
import socket
import subprocess
import sys
import threading
import time

import pytest

from server.app import python_analysis
from server.app import python_docker_runner as runner


def _docker_available() -> bool:
    """判断真实 Docker 验证是否可执行。

    只把**明确的环境前置缺失**当作不可用（引擎不可达 / 镜像未就绪）；
    任何**实现异常**必须显式暴露为失败，不得被宽泛 except 吞成 skip。
    """
    try:
        endpoint = runner.verified_endpoint()
    except runner.DockerRuntimeError as error:
        # 端点核实失败（含引擎不可达）→ 环境前置缺失
        if error.reason_category in ("endpoint_not_local", "runner_missing"):
            return False
        raise
    if not runner.verify_engine(endpoint):
        return False  # 引擎不可用（明确环境前置）
    try:
        return runner.image_present(runner._docker_config(), endpoint)
    except runner.DockerRuntimeError:
        return False  # 镜像未就绪（明确环境前置）


DOCKER_AVAILABLE = _docker_available()

pytestmark = pytest.mark.skipif(
    not DOCKER_AVAILABLE,
    reason="本机无可用的 Docker Linux 容器引擎或隔离镜像（环境阻断，非测试通过）",
)


# ---------------------------------------------------------------------------
# 受控端点与精确核对辅助
# ---------------------------------------------------------------------------


def _locked_endpoint() -> "runner.VerifiedEndpoint":
    """返回本次执行锁定的端点（清理核查必须用同一端点，不能回默认 context）。"""
    return runner.verified_endpoint()


def _start_control_listener() -> tuple[socket.socket, int]:
    """在宿主启动一个受控 TCP 监听器，绑定**所有接口**并返回 (socket, port)。

    绑定 ``0.0.0.0`` 而非 ``127.0.0.1``：容器内的 ``127.0.0.1`` 是**容器自身**
    的回环，不是宿主监听器。要让"同一受控端点"对容器而言可寻址，必须绑到
    宿主在容器网络命名空间里可见的地址（由同镜像 bridge 正对照核实）。
    """
    srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    srv.bind(("0.0.0.0", 0))
    srv.listen(4)
    return srv, srv.getsockname()[1]


def _control_can_reach(addr: str, port: int, timeout: float = 3.0) -> bool:
    """证明受控端点在宿主侧**确实可达**（正对照），否则网络测试无意义。"""
    try:
        with socket.create_connection((addr, port), timeout=timeout):
            return True
    except OSError:
        return False


def _trusted_control(code: str, *, network="none", pids_limit=1):
    """Fixed synthetic test control, NOT an alternate product execution path.

    Bypass the worker only to distinguish Docker resource limits from the
    worker's stricter process/socket denial. No generated/customer code.
    """
    endpoint, name = _locked_endpoint(), runner._new_container_name()
    args, _ = runner.build_run_arguments(runner._docker_config(), cpu_quota=1.0,
        cpu_seconds=10, memory_mb=128, pids_limit=pids_limit, scratch_mb=1, container_name=name)
    args[args.index("--network") + 1] = network
    args[-2:] = ["python3", "-c", code]
    try:
        proc = subprocess.run(["docker", *endpoint.cli_prefix(), *args[1:]],
            capture_output=True, timeout=15, shell=False, env=runner._cli_env())
        assert proc.returncode == 0, "fixed control did not run"
        return json.loads(proc.stdout)
    finally:
        assert runner._cleanup_container(name, endpoint=endpoint) == runner.CLEANUP_GONE


def _bridge_control_can_reach(addr: str, port: int) -> str | None:
    """Synthetic positive control in the same image, with bridge networking only.

    This test-only control never changes the runner's validated network=none
    configuration; no volumes, credentials or ports are added.
    """
    payload = _trusted_control(
        f"import socket,json\nhost=socket.gethostbyname({addr!r})\ns=socket.socket();s.settimeout(3);s.connect((host,{port}));s.close();print(json.dumps(host))",
        network="bridge")
    import ipaddress
    return str(ipaddress.IPv4Address(payload))


def _exact_container_exists(container_name: str) -> bool | None:
    """用**确切容器名**查询容器是否存在；查询失败返回 None（不冒充不存在）。

    必须使用**本次执行锁定的端点**，不得回落到默认 context（否则可能查到别处）。
    """
    return runner._container_exists(container_name, endpoint=_locked_endpoint())


def _container_inspect(container_name: str) -> str | None:
    proc = runner._run_docker_cli(
        ["container", "inspect", container_name], timeout=15.0, endpoint=_locked_endpoint()
    )
    if proc.returncode != 0:
        return None
    return proc.stdout


def _run_probe(monkeypatch, code: str, **overrides):
    """Observe only bounded synthetic evidence; all execution uses the real runner."""
    collectors, names = [], []
    real_collector, real_name = runner._OutputCollector, runner._new_container_name

    def collector_factory(*args, **kwargs):
        collector = real_collector(*args, **kwargs)
        collectors.append(collector)
        return collector

    def new_name():
        name = real_name()
        names.append(name)
        return name

    monkeypatch.setattr(runner, "_OutputCollector", collector_factory)
    monkeypatch.setattr(runner, "_new_container_name", new_name)
    endpoint = _locked_endpoint()
    options = dict(cpu_quota=1.0, cpu_seconds=60, memory_mb=128, pids_limit=1,
                   scratch_mb=4, max_output_bytes=65536, max_stderr_bytes=65536,
                   wall_timeout_seconds=15, endpoint=endpoint)
    options.update(overrides)
    result = runner.run_python_isolated(runner._docker_config(), code=code,
                                        input_payload={}, **options)
    assert names and collectors, "probe never reached the real container runner"
    for name in names:
        exists = runner._container_exists(name, endpoint=endpoint)
        assert exists is False, f"exact probe container not confirmed gone: {exists}"
    return result, collectors[0].stderr_bytes()


def test_real_analysis_produces_correct_result():
    """真实 Docker 内运行结构化分析，得到正确结果。"""
    result = python_analysis.run_isolated_analysis_internal(
        code=(
            "data = input['rows']\n"
            "values = [r['v'] for r in data]\n"
            "result = {'count': len(values), 'sum': sum(values), 'mean': sum(values)/len(values)}\n"
        ),
        input_payload={"rows": [{"v": 1}, {"v": 2}, {"v": 3}, {"v": 4}]},
    )
    assert result["status"] == "completed"
    assert result["result"] == {"count": 4, "sum": 10, "mean": 2.5}


def test_preinstalled_analysis_libraries_work_with_one_process():
    """Library imports must work without weakening the one-process limit."""
    result = python_analysis.run_isolated_analysis_internal(
        code="import numpy as np\nimport pandas as pd\nresult={'sum':int(np.array([1,2,3]).sum()),'rows':len(pd.DataFrame({'x':[1,2,3]}))}\n",
        input_payload={},
    )
    assert result == {"status": "completed", "result": {"sum": 6, "rows": 3}}


def test_no_network_access_in_container():
    """容器无外网/局域网/宿主服务——带**同一受控端点**的可达性正对照。

    正对照要求：受控端点必须在**非隔离条件下**、经**与容器相同的寻址路径**
    可达。容器内的 ``127.0.0.1`` 是容器自身回环，不能代表宿主监听器；因此
    正对照用**同镜像 bridge 容器实际连通的宿主 IP** 验证端点，
    再证明隔离容器连同一地址失败。不能用关闭端口/DNS 失败/同名不同命名空间冒充。
    """
    srv, port = _start_control_listener()
    try:
        assert _control_can_reach("127.0.0.1", port)
        # Docker Desktop's bridge gateway belongs to its VM, not the Windows
        # host. Resolve the host alias IN the bridge control and
        # feed that exact reachable numeric address to the isolated probe.
        gateway = _bridge_control_can_reach("host.docker.internal", port)
        assert gateway, "same-image bridge control could not reach the host endpoint"

        # 附加宿主对照：普通 Python 可连通同一监听器；容器寻址由 bridge 对照核实。
        control = subprocess.run(
            [sys.executable, "-c",
             "import socket,sys;"
             "s=socket.socket();s.settimeout(3);"
             f"s.connect(('127.0.0.1',{port}));print('REACHED')"],
            capture_output=True, text=True, timeout=15,
        )
        assert "REACHED" in control.stdout, f"对照进程未连上受控端点: {control}"

        # 隔离执行器：探同一受控端点（宿主网关）+ 外网，均应被拒
        net_result = python_analysis.run_isolated_analysis_internal(
            code=(
                "import socket\n"
                "out = []\n"
                f"for name, host, port in [('control','{gateway}',{port}), ('internet','1.1.1.1',53)]:\n"
                "    try:\n"
                "        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)\n"
                "        s.settimeout(3)\n"
                "        s.connect((host, port))\n"
                "        out.append(name + '=CONNECTED')\n"
                "    except Exception as e:\n"
                "        out.append(name + '=BLOCKED:' + str(e.errno))\n"
                "result = '|'.join(out)\n"
            ),
            input_payload={},
        )
    finally:
        srv.close()

    assert net_result["status"] == "completed", f"网络探测未完成: {net_result}"
    text = net_result["result"]
    assert "control=BLOCKED:1" in text, f"socket syscall must receive EPERM, not DNS failure or timeout: {text}"
    assert "internet=BLOCKED:1" in text, f"外网未隔离: {text}"


def test_loopback_is_container_own_not_host_service():
    """区分容器自身回环与宿主服务：经宿主网关可达的宿主监听端口，容器也连不到。

    与上一条的区别：这里显式以**宿主网关 IP**（非容器 127.0.0.1）作为宿主服务
    地址，证明容器即使在自身回环之外也触达不到宿主服务。
    """
    srv, port = _start_control_listener()
    try:
        gateway = _bridge_control_can_reach("host.docker.internal", port)
        assert gateway, "无法取得宿主网关 IP，宿主服务寻址不可验证"
        assert _control_can_reach("127.0.0.1", port)
        result = python_analysis.run_isolated_analysis_internal(
            code=(
                "import socket\n"
                f"gateway={gateway!r}\n"
                f"port={port}\n"
                "try:\n"
                "    s=socket.socket();s.settimeout(3);s.connect((gateway,port))\n"
                "    result='HOST_SERVICE_REACHED'\n"
                "except Exception as e:\n"
                "    result='HOST_SERVICE_BLOCKED:'+type(e).__name__\n"
            ),
            input_payload={},
        )
    finally:
        srv.close()
    assert result["status"] == "completed", f"探测未完成: {result}"
    assert result["result"].startswith("HOST_SERVICE_BLOCKED"), (
        f"容器经网关连到了宿主服务（网络未隔离）: {result['result']}"
    )


def test_host_file_not_mountable_or_readable(tmp_path):
    """宿主合成哨兵文件不可读（未挂载用户目录/仓库）。

    先在宿主**创建**哨兵并断言存在，再证明容器读不到；并额外证明探测逻辑
    本身可达（读一个容器内存在的文件应成功），否则"读失败"可能只是路径写错。
    """
    sentinel = tmp_path / "excelbro_host_sentinel.txt"
    sentinel.write_text("HOST_SENTINEL_SECRET_VALUE", encoding="utf-8")
    assert sentinel.exists(), "宿主哨兵未创建，测试无效"

    result = python_analysis.run_isolated_analysis_internal(
        code=(
            # 先证明容器内读取逻辑正常：读一个确定存在的文件
            "try:\n"
            "    ok = open('/etc/hostname').read()\n"
            "    probe = 'PROBE_OK'\n"
            "except Exception as e:\n"
            "    probe = 'PROBE_FAIL:' + type(e).__name__\n"
            f"try:\n"
            f"    data = open({str(sentinel)!r}).read()\n"
            f"    result = probe + '|READ_OK:' + data\n"
            f"except Exception as e:\n"
            f"    result = probe + '|READ_BLOCKED:' + type(e).__name__\n"
        ),
        input_payload={},
    )
    assert result["status"] == "completed", f"探测未完成: {result}"
    # 探测逻辑本身有效（能读到容器内文件）
    assert "PROBE_OK" in result["result"], f"容器内读取探测无效: {result['result']}"
    assert "READ_BLOCKED" in result["result"], f"宿主文件被读到: {result['result']}"
    assert "HOST_SENTINEL_SECRET_VALUE" not in str(result["result"])


def test_no_host_env_credentials(monkeypatch):
    """宿主环境凭据未继承（注入合成哨兵，不读真实密钥）。"""
    synthetic_keys = [
        ("EXCELBRO_SYNTHETIC_API_KEY", "SYNTHETIC_SECRET_ALPHA"),
        ("OPENAI_API_KEY", "SYNTHETIC_SECRET_BETA"),
    ]
    for key, value in synthetic_keys:
        monkeypatch.setenv(key, value)
    import os
    assert os.environ.get("EXCELBRO_SYNTHETIC_API_KEY") == "SYNTHETIC_SECRET_ALPHA"

    result = python_analysis.run_isolated_analysis_internal(
        code=(
            "import os\n"
            "result = {k: 'PRESENT' for k in "
            "['EXCELBRO_SYNTHETIC_API_KEY','OPENAI_API_KEY','ANTHROPIC_API_KEY'] "
            "if os.environ.get(k)}\n"
        ),
        input_payload={},
    )
    assert result["status"] == "completed", f"探测未完成: {result}"
    assert result["result"] == {}, f"宿主凭据被继承: {result['result']}"


def test_memory_limit_enforced(monkeypatch):
    """内存超限被容器强制终止——必须证明是**内存限额**触发，非启动失败/协议错误。

    先分配少量成功（证明解释器与协议正常、探测确实启动过），再分配超限内存；
    结果必须是明确的 limit/timeout/内存相关终止，且不得是"启动失败/镜像缺失/
    协议错误"这类旁路状态。
    """
    result, evidence = _run_probe(monkeypatch,
        code=(
            "import sys\n"
            "small = bytearray(16 * 1024 * 1024)\n"  # 16MB 先成功，证明解释器正常
            "limit = open('/sys/fs/cgroup/memory.max').read().strip()\n"
            "sys.stderr.write('MEMORY_STARTED:' + limit + '\\n');sys.stderr.flush()\n"
            "x = bytearray(256 * 1024 * 1024)\n"  # 256MB exceeds the actual 128MB cgroup
            "result = 'NOT_KILLED'\n"
        ),
    )
    assert b"MEMORY_STARTED:134217728" in evidence
    assert result.status == "error" and result.exit_code == 137, result
    assert result.result_payload is None


def test_stderr_is_bounded():
    """stderr 有独立上限：持续写 stderr 会被限流并提前终止（不无界堆积）。

    必须由**输出限额**（stderr 通道）触发，不能被泛化的 error/timeout 代替。
    """
    result = python_analysis.run_isolated_analysis_internal(
        code=(
            "import sys\n"
            "for _ in range(10_000_000):\n"
            "    sys.stderr.write('E' * 4096)\n"
        ),
        input_payload={},
    )
    # stderr 上限小于合理墙钟内可写量 → 必须是输出限额终止
    assert result["status"] == "limit_exceeded", (
        f"stderr 未被输出限额兜住（状态旁路）: {result}"
    )


def test_stdout_is_bounded_without_publishing_partial_result(monkeypatch):
    result, _ = _run_probe(monkeypatch,
        "import sys\n"
        "while True: sys.stdout.write('O'*4096);sys.stdout.flush()\n",
        max_output_bytes=4096,
    )
    assert result.status == "limit_exceeded"
    assert result.result_payload is None


def test_readonly_root_and_scratch_is_new_for_each_run():
    code = (
        "import os\nroot_errno=None\n"
        "try: open('/app/write-probe','w').write('synthetic')\n"
        "except OSError as error: root_errno=error.errno\n"
        "before=os.path.exists('/scratch/fresh-probe')\n"
        "open('/scratch/fresh-probe','w').write('synthetic')\n"
        "result={'uid':os.getuid(),'root_errno':root_errno,'before':before,"
        "'scratch':open('/scratch/fresh-probe').read()}\n"
    )
    for _ in range(2):
        result = python_analysis.run_isolated_analysis_internal(code, {})
        assert result == {"status":"completed","result":{
            "uid":65534,"root_errno":30,"before":False,"scratch":"synthetic"}}


def test_cpu_time_limit_enforced(monkeypatch):
    """累计 CPU 时间限额：纯计算死循环被 **CPU ulimit** 终止（区别于 wall timeout）。

    判据：把墙钟时限设得远大于 CPU 时限，若仍被终止且墙钟未到，则是 CPU 限额生效，
    而不是被普通 wall timeout 兜住。
    """
    started = time.monotonic()
    result, evidence = _run_probe(monkeypatch,
        code="import resource,sys\nsys.stderr.write('CPU_STARTED:'+str(resource.getrlimit(resource.RLIMIT_CPU))+'\\n');sys.stderr.flush()\nx=0\nwhile True: x+=1\n",
        cpu_seconds=2, wall_timeout_seconds=30,
    )
    elapsed = time.monotonic() - started
    assert b"CPU_STARTED:(2, 2)" in evidence
    assert result.status == "error" and result.exit_code == 137, result
    assert elapsed < 20


def test_timeout_terminates_and_container_cleaned_with_exact_identity(monkeypatch):
    """无限循环被超时终止，且**本次确切容器**已被清理（非仅全局集合差）。

    用确切容器名/inspect 核对：本运行创建的容器在结束后必须不存在；
    查询失败不得冒充"已清理"。
    """
    created: list[str] = []
    real_new_name = runner._new_container_name

    def spy_new_name() -> str:
        name = real_new_name()
        created.append(name)
        return name

    monkeypatch.setattr(runner, "_new_container_name", spy_new_name)

    limits = dict(python_analysis.python_execution_limits())
    limits["maxWallTimeSeconds"] = 2
    limits["maxCpuTimeSeconds"] = 60
    monkeypatch.setattr(python_analysis, "python_execution_limits", lambda: limits)
    result = python_analysis.run_isolated_analysis_internal(
        code="import time\nwhile True: time.sleep(0.05)",
        input_payload={},
    )
    assert result["status"] == "timeout", f"墙钟超时未终止: {result}"
    assert created, "未捕获到本次容器身份，无法精确核对清理"

    # 用确切容器名核对：必须确认不存在（查询失败 → None → 明确失败，不冒充）
    for name in created:
        exists = _exact_container_exists(name)
        assert exists is False, f"确切容器 {name} 未被清理（exists={exists}）"
        assert _container_inspect(name) is None, f"确切容器 {name} 仍可 inspect"


def test_pids_limit_enforced_distinct_from_startup_failure(monkeypatch):
    """进程/线程数限额被触发，且区别于启动失败/普通 error。

    注意：--pids-limit **只限制进程/线程数**，不表示禁止 exec；本测试只证明
    线程/进程爆炸被限额兜住，不声称禁止外部程序执行（那未落实）。
    """
    result = _trusted_control(
        (
            "import json\n"
            "import threading, time\n"
            "cap = int(open('/sys/fs/cgroup/pids.max').read())\n"
            "threads = []\n"
            "try:\n"
            "    for _ in range(1000):\n"
            "        t = threading.Thread(target=lambda: time.sleep(30), daemon=True)\n"
            "        t.start();threads.append(t)\n"
            "    result = {'limited':False}\n"
            "except RuntimeError:\n"
            "    result = {'limited':True,'started':len(threads),'cap':cap}\n"
            "print(json.dumps(result))\n"
        ),
        pids_limit=8,
    )
    assert result["limited"] is True
    assert result["cap"] == 8
    assert 0 < result["started"] < 8


def test_scratch_limit_enforced(monkeypatch):
    """scratch 临时盘限额被触发（读只根 + 有配额 tmpfs），区别于启动失败。"""
    from server.app import python_analysis as pa

    limits = dict(pa.python_execution_limits())
    limits["maxScratchMb"] = 1  # 1MB 便于触发
    monkeypatch.setattr(pa, "python_execution_limits", lambda: limits)

    result = pa.run_isolated_analysis_internal(
        code=(
            "import os\n"
            "open('/scratch/small.bin','wb').write(b'ok')\n"
            "fs=os.statvfs('/scratch');size=fs.f_blocks*fs.f_frsize\n"
            "data = b'x' * (64 * 1024 * 1024)\n"  # 64MB 超 1MB scratch
            "try:\n"
            "    open('/scratch/big.bin','wb').write(data)\n"
            "    result = {'limited':False}\n"
            "except OSError as e:\n"
            "    result = {'limited':True,'errno':e.errno,'size':size,'small':open('/scratch/small.bin','rb').read()==b'ok'}\n"
        ),
        input_payload={},
    )
    assert result == {"status":"completed","result":{"limited":True,"errno":28,"size":1048576,"small":True}}


def test_real_cancel_terminates_and_cleans_up(monkeypatch):
    """真实取消：运行中取消必须终止容器并按契约传播，且**确切容器**已清理。"""
    created: list[str] = []
    real_new_name = runner._new_container_name

    def spy_new_name() -> str:
        name = real_new_name()
        created.append(name)
        return name

    monkeypatch.setattr(runner, "_new_container_name", spy_new_name)

    collectors = []
    real_collector = runner._OutputCollector

    def observe_collector(*args, **kwargs):
        collector = real_collector(*args, **kwargs)
        collectors.append(collector)
        return collector

    monkeypatch.setattr(runner, "_OutputCollector", observe_collector)

    def cancel_check() -> bool:
        # Cancel only after the real analysis has started, not while pulling/launching.
        return bool(collectors and b"CANCEL_STARTED" in collectors[0].stderr_bytes())

    result = runner.run_python_isolated(
        runner._docker_config(),
        code="import sys,time\nsys.stderr.write('CANCEL_STARTED\\n');sys.stderr.flush()\nwhile True: time.sleep(0.1)\n",
        input_payload={},
        cpu_quota=1.0,
        cpu_seconds=60,
        memory_mb=1024,
        pids_limit=64,
        scratch_mb=64,
        max_output_bytes=65536,
        max_stderr_bytes=65536,
        wall_timeout_seconds=10,
        cancel_check=cancel_check,
        endpoint=_locked_endpoint(),
    )
    assert result.status == "cancelled", f"取消未传播: {result.status}"
    assert result.reason_category == "cancelled", result.reason_category
    assert b"CANCEL_STARTED" in collectors[0].stderr_bytes()
    assert created, "未捕获到本次容器身份"
    for name in created:
        assert _exact_container_exists(name) is False, f"取消后容器 {name} 未清理"


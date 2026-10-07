"""Docker 隔离执行器第七批缺陷的回归测试（Request H；第八批 Request I 后保留有效项）。

第七批修复的实现（基于 ``subprocess.run`` + 文件重定向）在第八批被整体重写为
``Popen`` + 并发有界读取（见 ``python_docker_runner.py``）。原先直接断言该实现的
用例（stdin bytes、超时文件重定向等）已由 ``test_python_docker_defects2.py`` 与
``test_python_docker_runner.py`` 以新接口覆盖。

本文件保留**与实现细节解耦、仍然有效**的回归项：
- 构建上下文正确（Dockerfile COPY 源真实存在于 server/ 上下文）；
- 容器名逐次唯一；
- 启动链禁用宿主 Docker 日志驱动；
- docker CLI 用最小隔离环境（不含 DOCKER_HOST/DOCKER_CONTEXT）；
- 隔离配置逐项校验（缺项拒绝）。
"""
from __future__ import annotations

import subprocess
from pathlib import Path

import pytest

from server.app import python_docker_runner as runner


REPO_ROOT = Path(__file__).resolve().parents[2]
DOCKER_DIR = REPO_ROOT / "server" / "docker"


# ---------------------------------------------------------------------------
# 缺陷 1：构建上下文正确
# ---------------------------------------------------------------------------


def test_dockerfile_copy_source_exists_in_build_context():
    """Dockerfile 里每个 COPY 的**源**必须在**实际构建上下文**中存在。"""
    dockerfile = DOCKER_DIR / "Dockerfile"
    assert dockerfile.exists(), "缺少 Dockerfile"
    text = dockerfile.read_text(encoding="utf-8")

    assert "context" in text.lower(), "Dockerfile 未声明构建上下文，COPY 源无从校验"
    context = REPO_ROOT / "server"
    assert "server/" in text, "Dockerfile 未把构建上下文写清为 server/"

    copy_sources: list[str] = []
    for raw in text.splitlines():
        line = raw.strip()
        if line.upper().startswith("COPY "):
            parts = line.split()[1:]
            copy_sources.extend(parts[:-1])

    assert copy_sources, "Dockerfile 没有任何 COPY"
    for src in copy_sources:
        assert not src.startswith("-"), f"COPY 源不应是选项：{src}"
        resolved = (context / src).resolve()
        assert resolved.exists(), (
            f"COPY 源 '{src}' 在构建上下文 {context} 中不存在：{resolved}"
        )
    assert any(s.endswith("docker_worker.py") for s in copy_sources), (
        "Dockerfile 未把 docker_worker.py 拷进镜像"
    )


# ---------------------------------------------------------------------------
# 缺陷 3：容器名逐次唯一
# ---------------------------------------------------------------------------


def test_container_name_is_unique_per_run():
    """容器名必须逐次唯一（不同运行不会误停别人的容器）。"""
    names = {runner._new_container_name() for _ in range(50)}
    assert len(names) == 50, "容器名出现重复，无法区分并发运行"
    assert all(n.startswith(runner._CONTAINER_NAME_PREFIX) for n in names), (
        "容器名未带可信前缀"
    )


# ---------------------------------------------------------------------------
# 缺陷 4：Docker 日志隐私
# ---------------------------------------------------------------------------


def test_build_run_arguments_disables_docker_log_driver():
    """容器不得把 stdout/stderr 落进宿主 Docker 日志（模型代码输出可能是隐私）。"""
    config = runner.DockerRuntimeConfig(image="img", image_digest="sha256:" + "0" * 64)
    args, _env = runner.build_run_arguments(
        config, cpu_quota=1.0, cpu_seconds=60, memory_mb=1024, pids_limit=2,
        scratch_mb=256, container_name="excel-bro-py-x",
    )
    joined = " ".join(args)
    assert "--log-driver none" in joined, (
        "容器未禁用日志驱动，模型代码输出会落进宿主 Docker 日志（隐私）"
    )


# ---------------------------------------------------------------------------
# 缺陷 5：引擎端点绑定（最小隔离环境）与隔离配置校验
# ---------------------------------------------------------------------------


def test_verify_engine_calls_use_host_isolated_env(monkeypatch):
    """docker CLI 用**最小隔离环境**（不含 DOCKER_HOST / DOCKER_CONTEXT）。

    判据（可变异）：宿主注入远端端点变量后，``_run_docker_cli`` 传给 subprocess
    的 env 里不得出现这两个变量；若实现改为继承宿主环境则变红。
    """
    captured_envs: list[dict] = []
    real_run = subprocess.run

    def capture_run(cmd, **kwargs):
        captured_envs.append(kwargs.get("env") or {})
        return subprocess.CompletedProcess(cmd, 0, "linux\n", "")

    monkeypatch.setattr(runner.subprocess, "run", capture_run)
    monkeypatch.setenv("DOCKER_HOST", "tcp://remote-daemon:2375")
    monkeypatch.setenv("DOCKER_CONTEXT", "remote")
    runner._run_docker_cli(["info", "--format", "{{.OSType}}"])

    assert captured_envs, "没有捕获到任何 docker CLI 调用"
    for env in captured_envs:
        assert "DOCKER_HOST" not in env, "docker CLI 继承了远端 DOCKER_HOST"
        assert "DOCKER_CONTEXT" not in env, "docker CLI 继承了远端 DOCKER_CONTEXT"


def test_isolation_config_validation_rejects_incomplete_config():
    """隔离配置必须**逐项**校验，缺任一项即拒绝执行（不能只信默认值）。"""
    assert hasattr(runner, "validate_isolation_config"), (
        "缺少 validate_isolation_config：隔离参数未被显式校验，配置错误会静默执行"
    )
    digest = "sha256:" + "0" * 64
    bad = runner.DockerRuntimeConfig(
        image="img", image_digest=digest,
        network="host", read_only=False, cap_drop="", no_new_privileges=False,
        non_root_user="0", log_driver="json-file",
    )
    ok, reason = runner.validate_isolation_config(bad)
    assert ok is False
    assert reason
    # 对照：完整合法配置必须通过
    good = runner.DockerRuntimeConfig(image="img", image_digest=digest)
    ok2, _ = runner.validate_isolation_config(good)
    assert ok2 is True


def test_isolation_config_validation_rejects_root_uid_variants():
    digest = "sha256:" + "0" * 64
    for user in ("0", "0:0", "root", "root:root", "0:1000", "root:1000"):
        cfg = runner.DockerRuntimeConfig(image="img", image_digest=digest, non_root_user=user)
        ok, _ = runner.validate_isolation_config(cfg)
        assert ok is False, f"root UID 变体 {user!r} 未被拒绝"

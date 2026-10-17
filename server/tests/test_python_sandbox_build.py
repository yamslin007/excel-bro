"""The build submits an explicit archive, never the repository or user Docker config."""
import io
import json
import tarfile
from pathlib import Path

import pytest

from scripts import build_python_sandbox as build
from server.app.python_docker_runner import VerifiedEndpoint


def test_context_has_only_the_three_explicit_files(tmp_path):
    for name in build.CONTEXT_FILES:
        path = tmp_path / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"approved-file")
    (tmp_path / ".env").write_text("SYNTHETIC_SECRET")
    (tmp_path / "app" / "new_unrelated_source.py").write_text("PRIVATE_SOURCE")
    archive = build.build_context(tmp_path)
    with tarfile.open(fileobj=io.BytesIO(archive)) as context:
        assert set(context.getnames()) == {
            ".dockerignore", "docker/Dockerfile", "app/docker_worker.py"}
        assert all(item.isfile() for item in context.getmembers())
    assert b"SYNTHETIC_SECRET" not in archive
    assert b"PRIVATE_SOURCE" not in archive


def test_context_rejects_symlink_instead_of_following_it(tmp_path):
    for name in build.CONTEXT_FILES:
        path = tmp_path / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"approved-file")
    outside = tmp_path / "outside"
    outside.write_bytes(b"PRIVATE_SOURCE")
    worker = tmp_path / "app" / "docker_worker.py"
    worker.unlink()
    try:
        worker.symlink_to(outside)
    except OSError:
        pytest.skip("symlink creation unavailable")
    if not worker.is_symlink():
        # 有些运行环境（含受限沙箱）会把 symlink_to 降级成"建一个空文件"：
        # 此时 is_symlink() 为假，本用例的前提（真的造出了符号链接）不成立，
        # 继续断言只会得到一条假失败。真正判定"是不是链接"的守卫由下面那条
        # 不依赖宿主能力的用例来验。
        pytest.skip("symlink creation unavailable")
    with pytest.raises(ValueError):
        build.build_context(tmp_path)


@pytest.mark.parametrize("member", ["file", "parent", "root"])
def test_context_rejects_a_symlink_at_every_context_path(monkeypatch, tmp_path, member):
    """符号链接守卫必须在**每个**上下文路径上都生效。

    上一条用例依赖宿主能不能真的造出符号链接（受限环境下会 skip），这一条不
    依赖：`is_symlink` 是唯一的判据，把它在被测对象上置真即可证明守卫真的会
    拒绝，而不是靠 `is_file`/路径包含检查蒙对。
    """
    for name in build.CONTEXT_FILES:
        path = tmp_path / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"approved-file")

    # 正向锚点：上下文里没有链接时，同一份上下文是可以正常打包的 ——
    # 证明下面那条 ValueError 是守卫拦的，不是打包恒失败。
    assert build.build_context(tmp_path)

    selected = {
        "file": tmp_path / "app" / "docker_worker.py",
        "parent": tmp_path / "app",
        "root": tmp_path,
    }[member]
    real_is_symlink = Path.is_symlink
    monkeypatch.setattr(
        Path,
        "is_symlink",
        lambda self: True if self == selected else real_is_symlink(self),
    )
    with pytest.raises(ValueError):
        build.build_context(tmp_path)


@pytest.mark.parametrize("driver", ["remote", "docker-container", "kubernetes"])
def test_remote_or_separate_builder_is_rejected(driver):
    with pytest.raises(ValueError):
        build.validate_default_builder(f"Name: default\nDriver: {driver}\nNodes:\nName: default\nEndpoint: default\nStatus: running\n")


def test_only_one_running_daemon_bound_default_builder_is_accepted():
    report = "Name: default\nDriver: docker\nNodes:\nName: default\nEndpoint: default\nStatus: running\n"
    build.validate_default_builder(report)
    for bad in [report.replace("running", "stopped"), report.replace("Endpoint: default", "Endpoint: remote"), report + "Endpoint: remote\n"]:
        with pytest.raises(ValueError):
            build.validate_default_builder(bad)


def test_client_environment_and_arguments_do_not_inherit_remote_overrides(monkeypatch, tmp_path):
    monkeypatch.setenv("DOCKER_HOST", "tcp://remote:2375")
    monkeypatch.setenv("BUILDX_BUILDER", "remote-builder")
    monkeypatch.setenv("OPENAI_API_KEY", "SYNTHETIC_SECRET")
    client = build.BuildClient(Path("docker"), tmp_path, VerifiedEndpoint("npipe", "npipe:////./pipe/docker_engine"))
    args = client.command(["buildx", "build", "--builder", "default", "-"])
    assert args[:5] == ["docker", "--config", str(tmp_path), "-H", "npipe:////./pipe/docker_engine"]
    env = build.client_environment()
    assert not {"DOCKER_HOST", "DOCKER_CONTEXT", "BUILDX_BUILDER", "OPENAI_API_KEY"} & env.keys()


def test_build_arguments_pin_base_and_builder_and_accept_only_archive_stdin():
    args = build.build_arguments("python@sha256:" + "a" * 64, "excel-bro-python-sandbox:test")
    assert args[-1] == "-"
    assert args[args.index("--builder") + 1] == "default"
    assert args[args.index("--build-arg") + 1] == "PYTHON_BASE=python@sha256:" + "a" * 64
    assert args[args.index("--file") + 1] == "docker/Dockerfile"


def test_base_identity_must_be_real_digest_and_linux_image():
    data = [{"Id": "sha256:" + "b" * 64, "Os": "linux", "Architecture": "amd64", "RepoDigests": ["python@sha256:" + "a" * 64]}]
    assert build.pinned_python_base(json.dumps(data)) == "python@sha256:" + "a" * 64
    for field, value in [("Os", "windows"), ("RepoDigests", []), ("RepoDigests", ["python:3.13-slim"])]:
        bad = [dict(data[0], **{field: value})]
        with pytest.raises(ValueError):
            build.pinned_python_base(json.dumps(bad))

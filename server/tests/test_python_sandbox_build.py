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

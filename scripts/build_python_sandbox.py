"""Authorized local build: python -m scripts.build_python_sandbox.

Only CONTEXT_FILES enter the stdin archive. Never reads the user's Docker
configuration, enables the product tool, or accepts model-supplied build options.
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import tarfile
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path

from server.app import python_docker_runner as runner

REPO_ROOT = Path(__file__).resolve().parents[1]
SERVER_ROOT = REPO_ROOT / "server"
CONTEXT_FILES = (".dockerignore", "docker/Dockerfile", "app/docker_worker.py")
DIGEST = re.compile(r"sha256:[0-9a-f]{64}\Z")


def build_context(root: Path = SERVER_ROOT) -> bytes:
    """Closed file-set snapshot: do not walk the server directory."""
    stream = io.BytesIO()
    with tarfile.open(fileobj=stream, mode="w") as archive:
        for name in CONTEXT_FILES:
            path = root / name
            if any(part.is_symlink() for part in (path, path.parent, root)):
                raise ValueError("Build input must not be a symbolic link")
            if not path.is_file() or not path.resolve().is_relative_to(root.resolve()):
                raise ValueError("Build input is missing or outside the build root")
            content = path.read_bytes()
            info = tarfile.TarInfo(name)
            info.size, info.mode = len(content), 0o644
            archive.addfile(info, io.BytesIO(content))
    return stream.getvalue()


def client_environment() -> dict[str, str]:
    return runner._cli_env()


@dataclass(frozen=True)
class BuildClient:
    executable: Path
    config_directory: Path
    endpoint: runner.VerifiedEndpoint

    def command(self, arguments: list[str]) -> list[str]:
        return [str(self.executable), "--config", str(self.config_directory),
                *runner.endpoint_cli_prefix(self.endpoint), *arguments]

    def run(self, arguments: list[str], *, data: bytes | None = None,
            visible: bool = False, timeout: int = 30) -> str:
        proc = subprocess.run(self.command(arguments), input=data, shell=False,
                              env=client_environment(), timeout=timeout,
                              stdout=None if visible else subprocess.PIPE,
                              stderr=None if visible else subprocess.PIPE)
        if proc.returncode:
            raise RuntimeError("Local Docker build operation failed")
        return (proc.stdout or b"").decode("utf-8", errors="strict")


def validate_default_builder(report: str) -> None:
    # Default + docker driver is bound to the daemon explicitly selected by -H.
    values = lambda key: re.findall(rf"^{key}:\s*(.*?)\s*$", report, re.MULTILINE)
    if (values("Driver") != ["docker"] or values("Endpoint") != ["default"]
            or values("Status") != ["running"] or values("Name")[:1] != ["default"]):
        raise ValueError("Builder is not the single running daemon-bound default")


def pinned_python_base(raw: str) -> str:
    images = json.loads(raw)
    if len(images) != 1 or images[0].get("Os") != "linux":
        raise ValueError("Python base image is not Linux")
    for ref in images[0].get("RepoDigests", []):
        name, sep, digest = ref.partition("@")
        if sep and name in ("python", "docker.io/library/python") and DIGEST.fullmatch(digest):
            return ref
    raise ValueError("Python base image has no verified repository digest")


def build_arguments(base: str, image: str) -> list[str]:
    return ["buildx", "build", "--builder", "default", "--load",
            "--provenance=false", "--sbom=false", "--progress", "plain",
            "--build-arg", "PYTHON_BASE=" + base, "--tag", image,
            "--file", "docker/Dockerfile", "-"]


def main() -> None:
    endpoint = runner.verified_endpoint()
    if not runner.verify_engine(endpoint):
        raise RuntimeError("Verified local Linux engine is not ready")
    executable = shutil.which("docker")
    if not executable:
        raise RuntimeError("Docker CLI is not available")
    plugin_directory = Path(executable).resolve().parent.parent / "cli-plugins"
    plugin_name = "docker-buildx.exe" if os.name == "nt" else "docker-buildx"
    if not (plugin_directory / plugin_name).is_file():
        raise RuntimeError("Bundled local buildx plugin was not found")
    context = build_context()
    with tarfile.open(fileobj=io.BytesIO(context)) as archive:
        hashes = {name: hashlib.sha256(archive.extractfile(name).read()).hexdigest()
                  for name in CONTEXT_FILES}
    image = "excel-bro-python-sandbox:build-" + uuid.uuid4().hex[:12]
    with tempfile.TemporaryDirectory(prefix=".tmp-python-build-client-", dir=REPO_ROOT) as temporary:
        config_directory = Path(temporary)
        (config_directory / "config.json").write_text(json.dumps({
            "cliPluginsExtraDirs": [str(plugin_directory)], "auths": {},
        }), encoding="utf-8")
        client = BuildClient(Path(executable), config_directory, endpoint)
        before = json.loads(client.run(["info", "--format", "{{json .}} "]))
        if before.get("OSType") != "linux" or not before.get("ID"):
            raise RuntimeError("Local engine identity is incomplete")
        validate_default_builder(client.run(["buildx", "inspect", "default"]))
        print("Verified local builder; archive files: " + ", ".join(CONTEXT_FILES), flush=True)
        client.run(["pull", "python:3.13-slim"], visible=True, timeout=600)
        base = pinned_python_base(client.run(["image", "inspect", "python:3.13-slim"]))
        print("Pinned Python base: " + base, flush=True)
        client.run(build_arguments(base, image), data=context, visible=True, timeout=1200)
        images = json.loads(client.run(["image", "inspect", image]))
        after = json.loads(client.run(["info", "--format", "{{json .}} "]))
        if after.get("ID") != before["ID"]:
            raise RuntimeError("Engine identity changed during the build")
        if (len(images) != 1 or not DIGEST.fullmatch(images[0].get("Id", ""))
                or images[0].get("Os") != "linux"
                or images[0].get("Config", {}).get("User") != "65534:65534"):
            raise RuntimeError("Built image identity or user could not be verified")
        receipt = {"image": image, "imageId": images[0]["Id"],
                   "repoDigests": images[0].get("RepoDigests", []), "baseImage": base,
                   "engineId": before["ID"], "endpointHost": endpoint.host,
                   "contextFiles": hashes, "productEnabled": False}
        print("BUILD_RECEIPT=" + json.dumps(receipt), flush=True)


if __name__ == "__main__":
    main()

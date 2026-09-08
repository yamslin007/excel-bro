from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import server.app.main as main_module
from server.app.folder_workbooks import restore_managed_folder_backup
from server.app.main import app


def _fingerprint(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _write_recovery_point(
    root: Path,
    *,
    undo_id: str,
    files: list[tuple[str, bytes | None, bytes]],
    status: str = "ready",
    revision: int = 1,
) -> Path:
    execution = root / ".excel-bro-backups" / "executions" / undo_id
    execution.mkdir(parents=True)
    manifest_files: list[dict[str, object]] = []
    for relative_text, prior, post in files:
        target = root / relative_text
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(post)
        if prior is None:
            backup_relative = None
            recovery_action = "remove"
        else:
            backup = execution / relative_text
            backup.parent.mkdir(parents=True, exist_ok=True)
            backup.write_bytes(prior)
            backup_relative = Path(relative_text).as_posix()
            recovery_action = "restore"
        manifest_files.append(
            {
                "originalRelativePath": Path(relative_text).as_posix(),
                "backupRelativePath": backup_relative,
                "recoveryAction": recovery_action,
                "priorFingerprint": _fingerprint(prior) if prior is not None else None,
                "postFingerprint": _fingerprint(post),
                "sizeBytes": len(prior) if prior is not None else 0,
                "state": "applied",
            }
        )
    (execution / "manifest.json").write_text(
        json.dumps(
            {
                "version": 2,
                "revision": revision,
                "undoId": undo_id,
                "status": status,
                "createdAt": "2026-09-08T10:00:00+08:00",
                "planId": "plan-recovery-test",
                "planTitle": "恢复测试",
                "files": manifest_files,
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return execution


def test_managed_recovery_restores_existing_file_and_removes_new_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    root = tmp_path / "work"
    undo_id = "b" * 32
    _write_recovery_point(
        root,
        undo_id=undo_id,
        files=[
            ("existing.xlsx", b"before", b"after"),
            ("new.xlsx", None, b"created"),
        ],
    )

    result = restore_managed_folder_backup(
        root,
        undo_id=undo_id,
        expected_revision=1,
        recovery_operation_id="recovery_test_1",
    )

    assert (root / "existing.xlsx").read_bytes() == b"before"
    assert not (root / "new.xlsx").exists()
    assert sorted(result.filesRestored) == ["existing.xlsx", "new.xlsx"]
    manifest = json.loads(
        (
            root
            / ".excel-bro-backups"
            / "executions"
            / undo_id
            / "manifest.json"
        ).read_text(encoding="utf-8")
    )
    assert manifest["status"] == "restored"
    assert not (
        root / ".excel-bro-backups" / "operations" / "recovery_test_1"
    ).exists()


def test_managed_recovery_refuses_to_overwrite_later_user_edit(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    root = tmp_path / "work"
    undo_id = "c" * 32
    _write_recovery_point(
        root,
        undo_id=undo_id,
        files=[("edited.xlsx", b"before", b"after")],
    )
    (root / "edited.xlsx").write_bytes(b"user-edit")

    with pytest.raises(ValueError, match="又发生了变化"):
        restore_managed_folder_backup(
            root,
            undo_id=undo_id,
            expected_revision=1,
            recovery_operation_id="recovery_test_2",
        )

    assert (root / "edited.xlsx").read_bytes() == b"user-edit"


def test_interrupted_managed_recovery_continues_only_remaining_files(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    root = tmp_path / "work"
    undo_id = "d" * 32
    operation_id = "recovery_partial_1"
    _write_recovery_point(
        root,
        undo_id=undo_id,
        files=[
            ("first.xlsx", b"first-before", b"first-after"),
            ("second.xlsx", b"second-before", b"second-after"),
        ],
        status="recovering",
    )
    (root / "first.xlsx").write_bytes(b"first-before")
    operation = root / ".excel-bro-backups" / "operations" / operation_id
    operation.mkdir(parents=True)
    (operation / "manifest.json").write_text(
        json.dumps(
            {
                "version": 1,
                "operationId": operation_id,
                "operationType": "undo",
                "status": "applying",
                "sourceUndoIds": [undo_id],
                "files": [],
            }
        ),
        encoding="utf-8",
    )

    result = restore_managed_folder_backup(
        root,
        undo_id=undo_id,
        expected_revision=1,
        recovery_operation_id=operation_id,
    )

    assert (root / "first.xlsx").read_bytes() == b"first-before"
    assert (root / "second.xlsx").read_bytes() == b"second-before"
    assert result.filesRestored == ["second.xlsx"]
    assert not operation.exists()


def test_managed_recovery_api_is_idempotent_after_completion(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    main_module._folder_operations.clear()
    root = tmp_path / "work"
    undo_id = "e" * 32
    _write_recovery_point(
        root,
        undo_id=undo_id,
        files=[("book.xlsx", b"before", b"after")],
    )
    directory = main_module._folder_operation_ledger.register_directory(str(root))
    payload = {
        "directoryId": directory.directory_id,
        "undoId": undo_id,
        "expectedRevision": 1,
        "operationId": "recovery_api_1",
        "confirmed": True,
    }
    client = TestClient(app)

    first = client.post("/api/folders/recovery/restore", json=payload)
    second = client.post("/api/folders/recovery/restore", json=payload)

    assert first.status_code == 200
    assert second.status_code == 200
    assert first.json() == second.json()
    assert (root / "book.xlsx").read_bytes() == b"before"


def test_managed_recovery_reenters_running_identity_after_uncommitted_completion(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("EXCEL_BRO_CONFIG_DIR", str(tmp_path / "app-data"))
    root = tmp_path / "work"
    undo_id = "f" * 32
    _write_recovery_point(
        root,
        undo_id=undo_id,
        files=[("book.xlsx", b"before", b"after")],
    )
    directory = main_module._folder_operation_ledger.register_directory(str(root))
    operation_id = "recovery_uncommitted_1"
    request_fingerprint = hashlib.sha256(
        json.dumps(
            {
                "directoryId": directory.directory_id,
                "undoId": undo_id,
                "expectedRevision": 1,
                "direction": "restore_before",
            },
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    ).hexdigest()
    main_module._folder_operation_ledger.register(
        operation_id,
        request_fingerprint,
        work_directory=str(root),
        plan_id=f"recovery:{undo_id}",
    )

    response = TestClient(app).post(
        "/api/folders/recovery/restore",
        json={
            "directoryId": directory.directory_id,
            "undoId": undo_id,
            "expectedRevision": 1,
            "operationId": operation_id,
            "confirmed": True,
        },
    )

    assert response.status_code == 200
    assert response.json()["status"] == "completed"
    assert (root / "book.xlsx").read_bytes() == b"before"
    assert main_module._folder_operation_ledger.get(operation_id).status == "completed"

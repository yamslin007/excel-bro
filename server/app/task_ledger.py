from __future__ import annotations

import json
import hashlib
import os
import sqlite3
import threading
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Literal

from .app_data import task_state_path


FolderOperationStatus = Literal[
    "running", "completed", "failed", "needs_reconciliation"
]


@dataclass(frozen=True)
class FolderOperationRecord:
    operation_id: str
    request_fingerprint: str
    status: FolderOperationStatus
    last_sequence: int
    result: dict[str, Any] | None
    error: dict[str, Any] | None
    created_at: float
    updated_at: float
    executor_id: str
    revision: int
    work_directory: str | None
    session_id: str | None
    plan_id: str | None
    execution_id: str | None
    undo_id: str | None


ManagedDirectoryStatus = Literal["active", "unavailable", "unsafe", "removed"]


class ExpiredOperationError(RuntimeError):
    def __init__(self, operation_id: str) -> None:
        super().__init__(f"任务记录已过保留期：{operation_id}")
        self.operation_id = operation_id


@dataclass(frozen=True)
class ManagedDirectoryRecord:
    directory_id: str
    path: str
    status: ManagedDirectoryStatus
    registered_at: float
    last_seen_at: float
    last_scanned_at: float | None
    recovery_count: int
    attention_count: int
    operation_count: int
    last_error: str | None
    revision: int


class FolderOperationLedger:
    """SQLite authority for folder-operation identity and event replay."""

    def __init__(
        self,
        path_provider: Callable[[], Path] = task_state_path,
        *,
        executor_id: str | None = None,
    ) -> None:
        self._path_provider = path_provider
        self.executor_id = executor_id or uuid.uuid4().hex
        self._schema_lock = threading.Lock()
        self._initialized_paths: set[Path] = set()

    def _path(self) -> Path:
        return self._path_provider().resolve()

    def _connect(self) -> sqlite3.Connection:
        path = self._path()
        path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(path, timeout=10, isolation_level=None)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA busy_timeout = 10000")
        self._ensure_schema(connection, path)
        return connection

    def _ensure_schema(self, connection: sqlite3.Connection, path: Path) -> None:
        if path in self._initialized_paths:
            return
        with self._schema_lock:
            if path in self._initialized_paths:
                return
            connection.execute("PRAGMA journal_mode = WAL")
            connection.execute("PRAGMA synchronous = FULL")
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS folder_operations (
                    operation_id TEXT PRIMARY KEY,
                    request_fingerprint TEXT NOT NULL,
                    status TEXT NOT NULL CHECK (
                        status IN ('running', 'completed', 'failed', 'needs_reconciliation')
                    ),
                    last_sequence INTEGER NOT NULL DEFAULT 0 CHECK (last_sequence >= 0),
                    result_json TEXT,
                    error_json TEXT,
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL,
                    executor_id TEXT NOT NULL,
                    revision INTEGER NOT NULL DEFAULT 1 CHECK (revision >= 1),
                    work_directory TEXT,
                    session_id TEXT,
                    plan_id TEXT,
                    execution_id TEXT,
                    undo_id TEXT
                );
                CREATE TABLE IF NOT EXISTS folder_operation_events (
                    operation_id TEXT NOT NULL,
                    sequence INTEGER NOT NULL CHECK (sequence > 0),
                    payload_json TEXT NOT NULL,
                    created_at REAL NOT NULL,
                    PRIMARY KEY (operation_id, sequence),
                    FOREIGN KEY (operation_id)
                        REFERENCES folder_operations(operation_id) ON DELETE CASCADE
                );
                CREATE INDEX IF NOT EXISTS idx_folder_operations_status_updated
                    ON folder_operations(status, updated_at);
                CREATE TABLE IF NOT EXISTS folder_operation_tombstones (
                    operation_id TEXT PRIMARY KEY,
                    request_fingerprint TEXT NOT NULL,
                    expired_at REAL NOT NULL
                );
                CREATE TABLE IF NOT EXISTS managed_directories (
                    directory_id TEXT PRIMARY KEY,
                    path TEXT NOT NULL UNIQUE,
                    status TEXT NOT NULL CHECK (
                        status IN ('active', 'unavailable', 'unsafe', 'removed')
                    ),
                    registered_at REAL NOT NULL,
                    last_seen_at REAL NOT NULL,
                    last_scanned_at REAL,
                    recovery_count INTEGER NOT NULL DEFAULT 0 CHECK (recovery_count >= 0),
                    attention_count INTEGER NOT NULL DEFAULT 0 CHECK (attention_count >= 0),
                    operation_count INTEGER NOT NULL DEFAULT 0 CHECK (operation_count >= 0),
                    last_error TEXT,
                    revision INTEGER NOT NULL DEFAULT 1 CHECK (revision >= 1)
                );
                CREATE INDEX IF NOT EXISTS idx_managed_directories_status_seen
                    ON managed_directories(status, last_seen_at);
                """
            )
            operation_columns = {
                str(row["name"])
                for row in connection.execute(
                    "PRAGMA table_info(folder_operations)"
                ).fetchall()
            }
            for column in (
                "work_directory",
                "session_id",
                "plan_id",
                "execution_id",
                "undo_id",
            ):
                if column not in operation_columns:
                    connection.execute(
                        f"ALTER TABLE folder_operations ADD COLUMN {column} TEXT"
                    )
            connection.execute("PRAGMA user_version = 3")
            self._initialized_paths.add(path)

    @staticmethod
    def _decode(value: str | None) -> dict[str, Any] | None:
        if value is None:
            return None
        decoded = json.loads(value)
        if not isinstance(decoded, dict):
            raise RuntimeError("持久任务记录包含无效 JSON 对象")
        return decoded

    @classmethod
    def _record(cls, row: sqlite3.Row) -> FolderOperationRecord:
        return FolderOperationRecord(
            operation_id=str(row["operation_id"]),
            request_fingerprint=str(row["request_fingerprint"]),
            status=str(row["status"]),  # type: ignore[arg-type]
            last_sequence=int(row["last_sequence"]),
            result=cls._decode(row["result_json"]),
            error=cls._decode(row["error_json"]),
            created_at=float(row["created_at"]),
            updated_at=float(row["updated_at"]),
            executor_id=str(row["executor_id"]),
            revision=int(row["revision"]),
            work_directory=(
                str(row["work_directory"]) if row["work_directory"] is not None else None
            ),
            session_id=str(row["session_id"]) if row["session_id"] is not None else None,
            plan_id=str(row["plan_id"]) if row["plan_id"] is not None else None,
            execution_id=(
                str(row["execution_id"]) if row["execution_id"] is not None else None
            ),
            undo_id=str(row["undo_id"]) if row["undo_id"] is not None else None,
        )

    def register(
        self,
        operation_id: str,
        request_fingerprint: str,
        *,
        work_directory: str | None = None,
        session_id: str | None = None,
        plan_id: str | None = None,
    ) -> tuple[FolderOperationRecord, bool]:
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            tombstone = connection.execute(
                "SELECT operation_id FROM folder_operation_tombstones WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if tombstone is not None:
                connection.rollback()
                raise ExpiredOperationError(operation_id)
            row = connection.execute(
                "SELECT * FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if row is None:
                connection.execute(
                    """
                    INSERT INTO folder_operations (
                        operation_id, request_fingerprint, status, created_at,
                        updated_at, executor_id, work_directory, session_id, plan_id
                    ) VALUES (?, ?, 'running', ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        operation_id,
                        request_fingerprint,
                        now,
                        now,
                        self.executor_id,
                        work_directory,
                        session_id,
                        plan_id,
                    ),
                )
                row = connection.execute(
                    "SELECT * FROM folder_operations WHERE operation_id = ?",
                    (operation_id,),
                ).fetchone()
                connection.commit()
                assert row is not None
                return self._record(row), True
            if str(row["request_fingerprint"]) == request_fingerprint:
                connection.execute(
                    """
                    UPDATE folder_operations
                    SET work_directory = COALESCE(work_directory, ?),
                        session_id = COALESCE(session_id, ?),
                        plan_id = COALESCE(plan_id, ?)
                    WHERE operation_id = ?
                    """,
                    (work_directory, session_id, plan_id, operation_id),
                )
                row = connection.execute(
                    "SELECT * FROM folder_operations WHERE operation_id = ?",
                    (operation_id,),
                ).fetchone()
            connection.commit()
            assert row is not None
            return self._record(row), False

    def get(self, operation_id: str) -> FolderOperationRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        return self._record(row) if row is not None else None

    def expired_fingerprint(self, operation_id: str) -> str | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT request_fingerprint FROM folder_operation_tombstones WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        return str(row["request_fingerprint"]) if row is not None else None

    def compact_terminal_operations(self, retention_days: int) -> int:
        """Drop bulky terminal payloads while permanently reserving operation IDs."""

        cutoff = time.time() - max(1, retention_days) * 86400
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            rows = connection.execute(
                """
                SELECT operation_id, request_fingerprint
                FROM folder_operations
                WHERE status IN ('completed', 'failed') AND updated_at < ?
                """,
                (cutoff,),
            ).fetchall()
            for row in rows:
                connection.execute(
                    """
                    INSERT OR IGNORE INTO folder_operation_tombstones (
                        operation_id, request_fingerprint, expired_at
                    ) VALUES (?, ?, ?)
                    """,
                    (row["operation_id"], row["request_fingerprint"], now),
                )
            if rows:
                connection.executemany(
                    "DELETE FROM folder_operations WHERE operation_id = ?",
                    [(row["operation_id"],) for row in rows],
                )
            connection.commit()
        return len(rows)

    def reopen_for_recovery(
        self, operation_id: str, request_fingerprint: str
    ) -> FolderOperationRecord:
        """Reclaim an explicitly retried recovery operation, never a normal execution."""

        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE folder_operations
                SET status = 'running', error_json = NULL, updated_at = ?,
                    executor_id = ?, revision = revision + 1
                WHERE operation_id = ? AND request_fingerprint = ?
                    AND status IN ('failed', 'needs_reconciliation')
                    AND plan_id LIKE 'recovery:%'
                """,
                (
                    now,
                    self.executor_id,
                    operation_id,
                    request_fingerprint,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("恢复操作状态已变化，无法重新领取")
            row = connection.execute(
                "SELECT * FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._record(row)

    def events(self, operation_id: str) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT payload_json FROM folder_operation_events
                WHERE operation_id = ? ORDER BY sequence
                """,
                (operation_id,),
            ).fetchall()
        values: list[dict[str, Any]] = []
        for row in rows:
            value = json.loads(str(row["payload_json"]))
            if not isinstance(value, dict):
                raise RuntimeError("持久任务事件包含无效 JSON 对象")
            values.append(value)
        return values

    def append_event(
        self,
        operation_id: str,
        sequence: int,
        payload: dict[str, Any],
        *,
        execution_id: str | None = None,
    ) -> None:
        now = time.time()
        encoded = json.dumps(
            payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        )
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT status, last_sequence FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if row is None:
                raise RuntimeError("无法为未登记的文件夹任务保存进度")
            if row["status"] != "running":
                raise RuntimeError("文件夹任务已停止，不能继续保存进度")
            expected = int(row["last_sequence"]) + 1
            if sequence != expected:
                raise RuntimeError(
                    f"文件夹任务进度序号不连续：预期 {expected}，收到 {sequence}"
                )
            connection.execute(
                """
                INSERT INTO folder_operation_events (
                    operation_id, sequence, payload_json, created_at
                ) VALUES (?, ?, ?, ?)
                """,
                (operation_id, sequence, encoded, now),
            )
            connection.execute(
                """
                UPDATE folder_operations
                SET last_sequence = ?, updated_at = ?, revision = revision + 1,
                    execution_id = COALESCE(execution_id, ?)
                WHERE operation_id = ?
                """,
                (sequence, now, execution_id, operation_id),
            )
            connection.commit()

    def finish(
        self,
        operation_id: str,
        *,
        status: Literal["completed", "failed", "needs_reconciliation"],
        result: dict[str, Any] | None = None,
        error: dict[str, Any] | None = None,
        execution_id: str | None = None,
        undo_id: str | None = None,
    ) -> FolderOperationRecord:
        if status == "completed" and result is None:
            raise ValueError("completed 文件夹任务必须保存结果")
        if status != "completed" and error is None:
            raise ValueError("未完成的文件夹任务必须保存错误或核对原因")
        now = time.time()
        result_json = (
            json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
            if result is not None
            else None
        )
        error_json = (
            json.dumps(error, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
            if error is not None
            else None
        )
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE folder_operations
                SET status = ?, result_json = ?, error_json = ?, updated_at = ?,
                    revision = revision + 1,
                    execution_id = COALESCE(execution_id, ?),
                    undo_id = COALESCE(undo_id, ?)
                WHERE operation_id = ? AND status = 'running'
                """,
                (
                    status,
                    result_json,
                    error_json,
                    now,
                    execution_id,
                    undo_id,
                    operation_id,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("文件夹任务状态已经改变，不能重复提交结果")
            row = connection.execute(
                "SELECT * FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._record(row)

    @staticmethod
    def _directory_record(row: sqlite3.Row) -> ManagedDirectoryRecord:
        return ManagedDirectoryRecord(
            directory_id=str(row["directory_id"]),
            path=str(row["path"]),
            status=str(row["status"]),  # type: ignore[arg-type]
            registered_at=float(row["registered_at"]),
            last_seen_at=float(row["last_seen_at"]),
            last_scanned_at=(
                float(row["last_scanned_at"])
                if row["last_scanned_at"] is not None
                else None
            ),
            recovery_count=int(row["recovery_count"]),
            attention_count=int(row["attention_count"]),
            operation_count=int(row["operation_count"]),
            last_error=str(row["last_error"]) if row["last_error"] is not None else None,
            revision=int(row["revision"]),
        )

    @staticmethod
    def _directory_identity(path: str) -> str:
        normalized = os.path.normcase(os.path.abspath(path))
        return hashlib.sha256(normalized.encode("utf-8")).hexdigest()

    def register_directory(
        self, path: str, *, max_entries: int = 50
    ) -> ManagedDirectoryRecord:
        normalized = os.path.abspath(path)
        directory_id = self._directory_identity(normalized)
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            existing = connection.execute(
                "SELECT status FROM managed_directories WHERE directory_id = ?",
                (directory_id,),
            ).fetchone()
            if existing is None or existing["status"] == "removed":
                active_count = int(
                    connection.execute(
                        "SELECT COUNT(*) FROM managed_directories WHERE status != 'removed'"
                    ).fetchone()[0]
                )
                if active_count >= max(1, max_entries):
                    connection.rollback()
                    raise RuntimeError(
                        "已登记工作目录数量达到上限，请先移除不再使用的目录"
                    )
            connection.execute(
                """
                INSERT INTO managed_directories (
                    directory_id, path, status, registered_at, last_seen_at
                ) VALUES (?, ?, 'active', ?, ?)
                ON CONFLICT(directory_id) DO UPDATE SET
                    path = excluded.path,
                    status = 'active',
                    last_seen_at = excluded.last_seen_at,
                    last_error = NULL,
                    revision = managed_directories.revision + 1
                """,
                (directory_id, normalized, now, now),
            )
            row = connection.execute(
                "SELECT * FROM managed_directories WHERE directory_id = ?",
                (directory_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._directory_record(row)

    def managed_directories(
        self, *, include_removed: bool = False
    ) -> list[ManagedDirectoryRecord]:
        sql = "SELECT * FROM managed_directories"
        parameters: tuple[object, ...] = ()
        if not include_removed:
            sql += " WHERE status != ?"
            parameters = ("removed",)
        sql += " ORDER BY last_seen_at DESC"
        with self._connect() as connection:
            rows = connection.execute(sql, parameters).fetchall()
        return [self._directory_record(row) for row in rows]

    def managed_directory(self, directory_id: str) -> ManagedDirectoryRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM managed_directories WHERE directory_id = ?",
                (directory_id,),
            ).fetchone()
        return self._directory_record(row) if row is not None else None

    def record_directory_scan(
        self,
        directory_id: str,
        *,
        status: Literal["active", "unavailable", "unsafe"],
        recovery_count: int,
        attention_count: int,
        operation_count: int,
        error: str | None = None,
    ) -> ManagedDirectoryRecord:
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE managed_directories
                SET status = ?, last_scanned_at = ?, recovery_count = ?,
                    attention_count = ?, operation_count = ?, last_error = ?,
                    revision = revision + 1
                WHERE directory_id = ? AND status != 'removed'
                """,
                (
                    status,
                    now,
                    max(recovery_count, 0),
                    max(attention_count, 0),
                    max(operation_count, 0),
                    error[:2000] if error else None,
                    directory_id,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("已登记工作目录不存在或已移除")
            row = connection.execute(
                "SELECT * FROM managed_directories WHERE directory_id = ?",
                (directory_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._directory_record(row)

    def remove_directory(self, directory_id: str) -> bool:
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE managed_directories
                SET status = 'removed', last_seen_at = ?, revision = revision + 1
                WHERE directory_id = ? AND status != 'removed'
                """,
                (now, directory_id),
            ).rowcount
            connection.commit()
        return changed == 1

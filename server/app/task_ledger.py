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
from .capabilities import capability_int


FolderOperationStatus = Literal[
    "running", "completed", "failed", "needs_reconciliation"
]
PlanningTaskStatus = Literal[
    "running", "waiting_user", "completed", "failed", "cancelled"
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
    request: dict[str, Any] | None
    authorization: dict[str, Any] | None


@dataclass(frozen=True)
class FolderOperationCheckpoint:
    operation_id: str
    generation: int
    kind: str
    payload: dict[str, Any]
    created_at: float
    updated_at: float


@dataclass(frozen=True)
class PlanningTaskRecord:
    turn_id: str
    request_fingerprint: str
    request: dict[str, Any]
    status: PlanningTaskStatus
    last_sequence: int
    checkpoint: dict[str, Any] | None
    result: dict[str, Any] | None
    error: dict[str, Any] | None
    pause_reason: str | None
    active_elapsed_seconds: float
    approved_wall_time_seconds: float
    extension_count: int
    executor_id: str
    revision: int
    created_at: float
    updated_at: float


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
                CREATE TABLE IF NOT EXISTS folder_operation_checkpoints (
                    operation_id TEXT PRIMARY KEY,
                    generation INTEGER NOT NULL CHECK (generation >= 1),
                    kind TEXT NOT NULL,
                    payload_json TEXT NOT NULL,
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL,
                    FOREIGN KEY (operation_id)
                        REFERENCES folder_operations(operation_id) ON DELETE CASCADE
                );
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
                CREATE TABLE IF NOT EXISTS planning_tasks (
                    turn_id TEXT PRIMARY KEY,
                    request_fingerprint TEXT NOT NULL,
                    request_json TEXT NOT NULL,
                    status TEXT NOT NULL CHECK (
                        status IN ('running', 'waiting_user', 'completed', 'failed', 'cancelled')
                    ),
                    last_sequence INTEGER NOT NULL DEFAULT 0 CHECK (last_sequence >= 0),
                    checkpoint_json TEXT,
                    result_json TEXT,
                    error_json TEXT,
                    pause_reason TEXT,
                    active_elapsed_seconds REAL NOT NULL DEFAULT 0 CHECK (active_elapsed_seconds >= 0),
                    approved_wall_time_seconds REAL NOT NULL CHECK (approved_wall_time_seconds > 0),
                    extension_count INTEGER NOT NULL DEFAULT 0 CHECK (extension_count >= 0),
                    executor_id TEXT NOT NULL,
                    revision INTEGER NOT NULL DEFAULT 1 CHECK (revision >= 1),
                    created_at REAL NOT NULL,
                    updated_at REAL NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_planning_tasks_status_updated
                    ON planning_tasks(status, updated_at);
                CREATE TABLE IF NOT EXISTS planning_task_events (
                    turn_id TEXT NOT NULL,
                    sequence INTEGER NOT NULL CHECK (sequence > 0),
                    payload_json TEXT NOT NULL,
                    created_at REAL NOT NULL,
                    PRIMARY KEY (turn_id, sequence),
                    FOREIGN KEY (turn_id)
                        REFERENCES planning_tasks(turn_id) ON DELETE CASCADE
                );
                CREATE TABLE IF NOT EXISTS planning_budget_extensions (
                    turn_id TEXT NOT NULL,
                    request_id TEXT NOT NULL,
                    added_seconds REAL NOT NULL CHECK (added_seconds >= 0),
                    created_at REAL NOT NULL,
                    PRIMARY KEY (turn_id, request_id),
                    FOREIGN KEY (turn_id)
                        REFERENCES planning_tasks(turn_id) ON DELETE CASCADE
                );
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
                "request_json",
                "authorization_json",
            ):
                if column not in operation_columns:
                    connection.execute(
                        f"ALTER TABLE folder_operations ADD COLUMN {column} TEXT"
                    )
            connection.execute("PRAGMA user_version = 5")
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
            request=cls._decode(row["request_json"]),
            authorization=cls._decode(row["authorization_json"]),
        )

    @classmethod
    def _checkpoint_record(
        cls, row: sqlite3.Row
    ) -> FolderOperationCheckpoint:
        payload = cls._decode(str(row["payload_json"]))
        if payload is None:
            raise RuntimeError("持久任务检查点缺少内容")
        return FolderOperationCheckpoint(
            operation_id=str(row["operation_id"]),
            generation=int(row["generation"]),
            kind=str(row["kind"]),
            payload=payload,
            created_at=float(row["created_at"]),
            updated_at=float(row["updated_at"]),
        )

    @classmethod
    def _planning_record(cls, row: sqlite3.Row) -> PlanningTaskRecord:
        request = cls._decode(str(row["request_json"]))
        if request is None:
            raise RuntimeError("规划任务缺少原始请求")
        return PlanningTaskRecord(
            turn_id=str(row["turn_id"]),
            request_fingerprint=str(row["request_fingerprint"]),
            request=request,
            status=str(row["status"]),  # type: ignore[arg-type]
            last_sequence=int(row["last_sequence"]),
            checkpoint=cls._decode(row["checkpoint_json"]),
            result=cls._decode(row["result_json"]),
            error=cls._decode(row["error_json"]),
            pause_reason=(
                str(row["pause_reason"]) if row["pause_reason"] is not None else None
            ),
            active_elapsed_seconds=float(row["active_elapsed_seconds"]),
            approved_wall_time_seconds=float(row["approved_wall_time_seconds"]),
            extension_count=int(row["extension_count"]),
            executor_id=str(row["executor_id"]),
            revision=int(row["revision"]),
            created_at=float(row["created_at"]),
            updated_at=float(row["updated_at"]),
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

    @staticmethod
    def _encode(value: dict[str, Any]) -> str:
        return json.dumps(
            value, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        )

    def save_execution_context(
        self,
        operation_id: str,
        *,
        request: dict[str, Any],
        authorization: dict[str, Any],
    ) -> FolderOperationRecord:
        """Persist the confirmed request and exact authorization before any write."""

        request_json = self._encode(request)
        authorization_json = self._encode(authorization)
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT * FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if row is None:
                connection.rollback()
                raise RuntimeError("无法为未登记的文件夹任务保存执行现场")
            if row["status"] != "running" or row["executor_id"] != self.executor_id:
                connection.rollback()
                raise RuntimeError("文件夹任务已被其他执行者接管或停止")
            existing_request = row["request_json"]
            existing_authorization = row["authorization_json"]
            if existing_request is not None and str(existing_request) != request_json:
                connection.rollback()
                raise RuntimeError("已登记任务的确认请求不可被替换")
            if (
                existing_authorization is not None
                and str(existing_authorization) != authorization_json
            ):
                connection.rollback()
                raise RuntimeError("已登记任务的授权范围不可被扩大或替换")
            connection.execute(
                """
                UPDATE folder_operations
                SET request_json = COALESCE(request_json, ?),
                    authorization_json = COALESCE(authorization_json, ?),
                    updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND status = 'running'
                    AND executor_id = ?
                """,
                (
                    request_json,
                    authorization_json,
                    now,
                    operation_id,
                    self.executor_id,
                ),
            )
            updated = connection.execute(
                "SELECT * FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert updated is not None
        return self._record(updated)

    def checkpoint(self, operation_id: str) -> FolderOperationCheckpoint | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM folder_operation_checkpoints WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        return self._checkpoint_record(row) if row is not None else None

    def commit_checkpoint(
        self,
        operation_id: str,
        *,
        kind: str,
        payload: dict[str, Any],
    ) -> FolderOperationCheckpoint:
        """Atomically advance the durable checkpoint for the current owner."""

        if not kind.strip():
            raise ValueError("检查点类型不能为空")
        encoded = self._encode(payload)
        encoded_size = len(encoded.encode("utf-8"))
        maximum_size = capability_int("taskRecovery", "maxCheckpointBytes")
        if encoded_size > maximum_size:
            raise ValueError(
                f"任务检查点超过容量上限（{encoded_size} > {maximum_size} 字节）"
            )
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            operation = connection.execute(
                "SELECT status, executor_id FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if operation is None:
                connection.rollback()
                raise RuntimeError("无法为未登记的文件夹任务保存检查点")
            if (
                operation["status"] != "running"
                or operation["executor_id"] != self.executor_id
            ):
                connection.rollback()
                raise RuntimeError("过期执行者不能提交新的检查点")
            existing = connection.execute(
                """
                SELECT generation, created_at
                FROM folder_operation_checkpoints WHERE operation_id = ?
                """,
                (operation_id,),
            ).fetchone()
            generation = int(existing["generation"]) + 1 if existing else 1
            created_at = float(existing["created_at"]) if existing else now
            connection.execute(
                """
                INSERT INTO folder_operation_checkpoints (
                    operation_id, generation, kind, payload_json, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(operation_id) DO UPDATE SET
                    generation = excluded.generation,
                    kind = excluded.kind,
                    payload_json = excluded.payload_json,
                    updated_at = excluded.updated_at
                """,
                (operation_id, generation, kind, encoded, created_at, now),
            )
            connection.execute(
                """
                UPDATE folder_operations
                SET updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND status = 'running'
                    AND executor_id = ?
                """,
                (now, operation_id, self.executor_id),
            )
            row = connection.execute(
                "SELECT * FROM folder_operation_checkpoints WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._checkpoint_record(row)

    def claim_for_resume(
        self, operation_id: str, *, expected_revision: int
    ) -> FolderOperationRecord:
        """CAS claim for an explicit restart/resume request."""

        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE folder_operations
                SET status = 'running', error_json = NULL, executor_id = ?,
                    updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND revision = ?
                    AND status IN ('running', 'needs_reconciliation')
                    AND request_json IS NOT NULL AND authorization_json IS NOT NULL
                    AND EXISTS (
                        SELECT 1 FROM folder_operation_checkpoints checkpoint
                        WHERE checkpoint.operation_id = folder_operations.operation_id
                    )
                """,
                (self.executor_id, now, operation_id, expected_revision),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("任务状态已变化、现场不完整或已被其他执行者领取")
            row = connection.execute(
                "SELECT * FROM folder_operations WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._record(row)

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
                """
                SELECT status, last_sequence, executor_id
                FROM folder_operations WHERE operation_id = ?
                """,
                (operation_id,),
            ).fetchone()
            if row is None:
                raise RuntimeError("无法为未登记的文件夹任务保存进度")
            if row["status"] != "running":
                raise RuntimeError("文件夹任务已停止，不能继续保存进度")
            if row["executor_id"] != self.executor_id:
                raise RuntimeError("过期执行者不能继续保存进度")
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
                WHERE operation_id = ? AND executor_id = ?
                """,
                (sequence, now, execution_id, operation_id, self.executor_id),
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
                WHERE operation_id = ? AND status = 'running' AND executor_id = ?
                """,
                (
                    status,
                    result_json,
                    error_json,
                    now,
                    execution_id,
                    undo_id,
                    operation_id,
                    self.executor_id,
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

    def register_planning_task(
        self,
        turn_id: str,
        request_fingerprint: str,
        request: dict[str, Any],
        *,
        approved_wall_time_seconds: float,
    ) -> tuple[PlanningTaskRecord, bool]:
        if approved_wall_time_seconds <= 0:
            raise ValueError("规划任务的初始运行份额必须大于 0")
        encoded_request = self._encode(request)
        if len(encoded_request.encode("utf-8")) > capability_int(
            "taskRecovery", "maxCheckpointBytes"
        ):
            raise ValueError("规划请求超过可持久化容量上限")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
            if row is None:
                connection.execute(
                    """
                    INSERT INTO planning_tasks (
                        turn_id, request_fingerprint, request_json, status,
                        approved_wall_time_seconds, executor_id, created_at, updated_at
                    ) VALUES (?, ?, ?, 'running', ?, ?, ?, ?)
                    """,
                    (
                        turn_id,
                        request_fingerprint,
                        encoded_request,
                        approved_wall_time_seconds,
                        self.executor_id,
                        now,
                        now,
                    ),
                )
                row = connection.execute(
                    "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
                ).fetchone()
                connection.commit()
                assert row is not None
                return self._planning_record(row), True
            connection.commit()
        record = self._planning_record(row)
        if record.request_fingerprint != request_fingerprint:
            raise RuntimeError("该 turnId 已用于另一份规划请求")
        return record, False

    def planning_task(self, turn_id: str) -> PlanningTaskRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
        return self._planning_record(row) if row is not None else None

    def planning_events(self, turn_id: str) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT payload_json FROM planning_task_events
                WHERE turn_id = ? ORDER BY sequence
                """,
                (turn_id,),
            ).fetchall()
        result: list[dict[str, Any]] = []
        for row in rows:
            value = self._decode(str(row["payload_json"]))
            if value is None:
                raise RuntimeError("规划进度事件缺少内容")
            result.append(value)
        return result

    def append_planning_event(
        self, turn_id: str, sequence: int, payload: dict[str, Any]
    ) -> PlanningTaskRecord:
        encoded = self._encode(payload)
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                """
                SELECT status, last_sequence, executor_id
                FROM planning_tasks WHERE turn_id = ?
                """,
                (turn_id,),
            ).fetchone()
            if row is None:
                connection.rollback()
                raise RuntimeError("无法为未登记的规划任务保存进度")
            if row["status"] != "running" or row["executor_id"] != self.executor_id:
                connection.rollback()
                raise RuntimeError("过期执行者不能保存规划进度")
            expected = int(row["last_sequence"]) + 1
            if sequence != expected:
                connection.rollback()
                raise RuntimeError(
                    f"规划进度序号不连续：预期 {expected}，收到 {sequence}"
                )
            connection.execute(
                """
                INSERT INTO planning_task_events (
                    turn_id, sequence, payload_json, created_at
                ) VALUES (?, ?, ?, ?)
                """,
                (turn_id, sequence, encoded, now),
            )
            connection.execute(
                """
                UPDATE planning_tasks
                SET last_sequence = ?, updated_at = ?, revision = revision + 1
                WHERE turn_id = ? AND executor_id = ? AND status = 'running'
                """,
                (sequence, now, turn_id, self.executor_id),
            )
            updated = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
            connection.commit()
        assert updated is not None
        return self._planning_record(updated)

    def save_planning_checkpoint(
        self,
        turn_id: str,
        checkpoint: dict[str, Any],
        *,
        active_elapsed_seconds: float,
    ) -> PlanningTaskRecord:
        encoded = self._encode(checkpoint)
        if len(encoded.encode("utf-8")) > capability_int(
            "taskRecovery", "maxCheckpointBytes"
        ):
            raise ValueError("规划检查点超过可持久化容量上限")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE planning_tasks
                SET checkpoint_json = ?, active_elapsed_seconds = ?,
                    updated_at = ?, revision = revision + 1
                WHERE turn_id = ? AND status = 'running' AND executor_id = ?
                """,
                (
                    encoded,
                    max(active_elapsed_seconds, 0.0),
                    now,
                    turn_id,
                    self.executor_id,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("过期执行者不能保存规划检查点")
            row = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._planning_record(row)

    def pause_planning_task(
        self,
        turn_id: str,
        checkpoint: dict[str, Any],
        *,
        active_elapsed_seconds: float,
        reason: str,
    ) -> PlanningTaskRecord:
        encoded = self._encode(checkpoint)
        if len(encoded.encode("utf-8")) > capability_int(
            "taskRecovery", "maxCheckpointBytes"
        ):
            raise ValueError("规划检查点超过可持久化容量上限")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE planning_tasks
                SET status = 'waiting_user', checkpoint_json = ?,
                    active_elapsed_seconds = ?, pause_reason = ?, updated_at = ?,
                    revision = revision + 1
                WHERE turn_id = ? AND status = 'running' AND executor_id = ?
                """,
                (
                    encoded,
                    max(active_elapsed_seconds, 0.0),
                    reason[:1000],
                    now,
                    turn_id,
                    self.executor_id,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("规划任务状态已变化，不能重复暂停")
            row = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._planning_record(row)

    def claim_planning_continuation(
        self,
        turn_id: str,
        *,
        expected_revision: int,
        request_id: str,
        extension_seconds: float,
        maximum_extensions: int,
        hard_wall_time_seconds: float,
    ) -> tuple[PlanningTaskRecord, bool]:
        """CAS claim; a stable request_id grants a budget extension at most once."""

        if not request_id.strip():
            raise ValueError("预算追加请求标识不能为空")
        if extension_seconds < 0 or hard_wall_time_seconds <= 0:
            raise ValueError("规划预算参数无效")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
            if row is None:
                connection.rollback()
                raise RuntimeError("未找到可继续的规划任务")
            duplicate = connection.execute(
                """
                SELECT added_seconds FROM planning_budget_extensions
                WHERE turn_id = ? AND request_id = ?
                """,
                (turn_id, request_id),
            ).fetchone()
            if duplicate is not None:
                if row["status"] not in {"running", "waiting_user"}:
                    connection.rollback()
                    return self._planning_record(row), False
                connection.execute(
                    """
                    UPDATE planning_tasks
                    SET status = 'running', executor_id = ?, pause_reason = NULL,
                        updated_at = ?, revision = revision + 1
                    WHERE turn_id = ?
                    """,
                    (self.executor_id, now, turn_id),
                )
                updated = connection.execute(
                    "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
                ).fetchone()
                connection.commit()
                assert updated is not None
                return self._planning_record(updated), True
            if int(row["revision"]) != expected_revision:
                connection.rollback()
                raise RuntimeError("规划任务状态已变化，请刷新后重试")
            if row["status"] not in {"waiting_user", "running"}:
                connection.rollback()
                raise RuntimeError("规划任务当前不能继续")
            active_elapsed = float(row["active_elapsed_seconds"])
            approved = float(row["approved_wall_time_seconds"])
            extension_count = int(row["extension_count"])
            added = 0.0
            if row["status"] == "waiting_user":
                if extension_count >= max(maximum_extensions, 0):
                    connection.rollback()
                    raise RuntimeError("规划任务已达到允许的继续次数上限")
                remaining_hard = hard_wall_time_seconds - approved
                added = min(extension_seconds, max(remaining_hard, 0.0))
                if added <= 0 or active_elapsed >= hard_wall_time_seconds:
                    connection.rollback()
                    raise RuntimeError("规划任务已达到有效运行总时长上限")
            connection.execute(
                """
                INSERT INTO planning_budget_extensions (
                    turn_id, request_id, added_seconds, created_at
                ) VALUES (?, ?, ?, ?)
                """,
                (turn_id, request_id, added, now),
            )
            changed = connection.execute(
                """
                UPDATE planning_tasks
                SET status = 'running', executor_id = ?, pause_reason = NULL,
                    approved_wall_time_seconds = approved_wall_time_seconds + ?,
                    extension_count = extension_count + ?, updated_at = ?,
                    revision = revision + 1
                WHERE turn_id = ? AND revision = ?
                """,
                (
                    self.executor_id,
                    added,
                    1 if added > 0 else 0,
                    now,
                    turn_id,
                    expected_revision,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("规划任务已被其他请求领取")
            updated = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
            connection.commit()
        assert updated is not None
        return self._planning_record(updated), True

    def finish_planning_task(
        self,
        turn_id: str,
        *,
        status: Literal["completed", "failed", "cancelled"],
        result: dict[str, Any] | None = None,
        error: dict[str, Any] | None = None,
    ) -> PlanningTaskRecord:
        if status == "completed" and result is None:
            raise ValueError("完成的规划任务必须保存结果")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE planning_tasks
                SET status = ?, result_json = ?, error_json = ?, pause_reason = NULL,
                    updated_at = ?, revision = revision + 1
                WHERE turn_id = ? AND status = 'running' AND executor_id = ?
                """,
                (
                    status,
                    self._encode(result) if result is not None else None,
                    self._encode(error) if error is not None else None,
                    now,
                    turn_id,
                    self.executor_id,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("规划任务状态已经改变，不能重复提交结果")
            row = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._planning_record(row)

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

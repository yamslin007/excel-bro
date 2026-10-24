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

_READONLY_RETRY_ATTEMPTS = 6
# 当前 schema 版本（与 _initialize_schema 里的 PRAGMA user_version 一致）。
_SCHEMA_VERSION = 8
# SQLite 只读相关错误码（用于按**准确类别**判定重试，而非匹配错误字符串）。
# 见 https://www.sqlite.org/rescode.html ：SQLITE_READONLY(8) 及其扩展码。
_SQLITE_READONLY_CODES = {8, 8 | (1 << 8), 8 | (2 << 8), 8 | (3 << 8)}


def _is_readonly_error(error: sqlite3.Error) -> bool:
    """按**准确错误类别**判断是否为只读错误（不依赖错误字符串）。

    优先读取 CPython 在 SQLite 异常上暴露的 ``sqlite_errorcode`` /
    ``sqlite_errorname``（3.11+）：``SQLITE_READONLY`` 扩展码归为只读。
    """
    code = getattr(error, "sqlite_errorcode", None)
    if isinstance(code, int) and code in _SQLITE_READONLY_CODES:
        return True
    name = getattr(error, "sqlite_errorname", None)
    if isinstance(name, str) and name.startswith("SQLITE_READONLY"):
        return True
    # 兜底：极旧 SQLite/CPython 未暴露诊断属性时，退回字符串匹配。
    return "readonly" in str(error).lower()


def _schema_is_current(path: Path) -> bool:
    """判断已有库是否已达当前 schema 版本（避免重跑 bootstrap 与迁移）。

    库不存在、不是 SQLite 库或读取失败 → 返回 False（走完整 bootstrap，
    由 bootstrap 自己失败关闭）。
    """
    if not path.exists():
        return False
    try:
        with sqlite3.connect(path, timeout=5) as connection:
            row = connection.execute("PRAGMA user_version").fetchone()
    except sqlite3.Error:
        return False
    return bool(row) and int(row[0]) >= _SCHEMA_VERSION


def _connection_is_writable(connection: sqlite3.Connection) -> bool:
    """探测连接是否可写，用于识别 Windows 上 WAL 短暂只读句柄。

    只用来决定是否重开连接；不改变任何业务写入路径，也不把只读当成功。
    注意：此探测会取写锁，**只能用于可写连接**，绝不能用于纯读取路径
    （否则纯读也会被串行到写锁，见第十批 F4）。
    """

    try:
        connection.execute("BEGIN IMMEDIATE")
        connection.execute("ROLLBACK")
        return True
    except sqlite3.OperationalError as error:
        if _is_readonly_error(error):
            return False
        raise


FolderOperationStatus = Literal[
    "running", "completed", "failed", "needs_reconciliation"
]
PlanningTaskStatus = Literal[
    "running", "waiting_user", "completed", "failed", "cancelled"
]
OfficeExecutionStatus = Literal[
    "running", "needs_reconciliation", "completed", "failed"
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
    current_model_id: str | None
    model_switches: tuple[dict[str, Any], ...]
    executor_id: str
    revision: int
    created_at: float
    updated_at: float


@dataclass(frozen=True)
class OfficeExecutionRecord:
    operation_id: str
    request_fingerprint: str
    plan: dict[str, Any]
    authorization: dict[str, Any]
    workbook_identity: str
    task_id: str | None
    status: OfficeExecutionStatus
    next_action_index: int
    in_flight_action_index: int | None
    before_snapshot: dict[str, Any] | None
    checkpoint: dict[str, Any] | None
    result: dict[str, Any] | None
    error: dict[str, Any] | None
    owner_token: str
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
        """打开**可写**连接（用于写入/迁移）。

        可写性由连接**本身**保证：写路径的 ``BEGIN IMMEDIATE`` 会在只读句柄上
        如实失败（失败关闭），不需要在每次连接时主动探测可写——那种探测会让
        **纯读取**也被强制取得写锁（第十批 F4）。仅当本路径尚未完成 schema
        bootstrap 时才做一次 bootstrap。
        """
        path = self._path()
        path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_schema(path)
        return self._open_writable_connection(path)

    def _connect_readonly(self) -> sqlite3.Connection:
        """打开**只读**连接（用于纯查询）。

        读取不得被串行到写锁：只读连接不参与写事务，SQLite 在 WAL 下允许
        读与写并发。schema 若尚未初始化，先确保 bootstrap（幂等、只做一次）。
        """
        path = self._path()
        path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_schema(path)
        return self._open_readonly_connection(path)

    def _open_writable_connection(self, path: Path) -> sqlite3.Connection:
        # WAL 模式下 Windows/SQLite 偶发把刚打开的连接持成只读句柄；这是短暂状态，
        # 重开即可恢复，不能把短暂只读当成永久失败，更不能降级为只读或跳过写入。
        # 但**不得**在每次连接时用 BEGIN IMMEDIATE 主动探测可写（会使纯读也取写锁）。
        last_error: sqlite3.OperationalError | None = None
        for attempt in range(_READONLY_RETRY_ATTEMPTS):
            connection = sqlite3.connect(path, timeout=10, isolation_level=None)
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys = ON")
            connection.execute("PRAGMA busy_timeout = 10000")
            if _connection_is_writable(connection):
                return connection
            connection.close()
            last_error = sqlite3.OperationalError("attempt to write a readonly database")
            time.sleep(0.005 * (attempt + 1))
        raise last_error if last_error is not None else sqlite3.OperationalError(
            "无法获得可写数据库连接"
        )

    def _open_readonly_connection(self, path: Path) -> sqlite3.Connection:
        """只读连接：``mode=ro`` 打开，不参与写事务，不取写锁。"""
        connection = sqlite3.connect(
            f"file:{path.as_posix()}?mode=ro", timeout=10, isolation_level=None, uri=True
        )
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA busy_timeout = 10000")
        return connection

    def _ensure_schema(self, path: Path) -> None:
        if path in self._initialized_paths:
            return
        with self._schema_lock:
            if path in self._initialized_paths:
                return
            self._run_schema_bootstrap(path)
            self._initialized_paths.add(path)

    def _run_schema_bootstrap(self, path: Path) -> None:
        # 已有库（user_version 已达当前版本）无需重建：直接返回，避免每次新实例
        # 都重跑 journal_mode 切换与迁移——那会与并发写事务冲突（第十批 F4）。
        if _schema_is_current(path):
            return
        # bootstrap 连接自身也可能被 WAL 切换持成只读；同样重开，不得降级。
        last_error: sqlite3.OperationalError | None = None
        for attempt in range(_READONLY_RETRY_ATTEMPTS):
            bootstrap = sqlite3.connect(path, timeout=10, isolation_level=None)
            try:
                bootstrap.row_factory = sqlite3.Row
                self._initialize_schema(bootstrap)
            except sqlite3.OperationalError as error:
                bootstrap.close()
                if not _is_readonly_error(error):
                    raise
                last_error = error
                time.sleep(0.005 * (attempt + 1))
                continue
            bootstrap.close()
            return
        raise last_error if last_error is not None else sqlite3.OperationalError(
            "无法完成账本 schema 初始化"
        )

    def _initialize_schema(self, connection: sqlite3.Connection) -> None:
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
                current_model_id TEXT,
                model_switches_json TEXT NOT NULL DEFAULT '{"items":[]}',
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
                model_id TEXT,
                created_at REAL NOT NULL,
                PRIMARY KEY (turn_id, request_id),
                FOREIGN KEY (turn_id)
                    REFERENCES planning_tasks(turn_id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS office_executions (
                operation_id TEXT PRIMARY KEY,
                request_fingerprint TEXT NOT NULL,
                plan_json TEXT NOT NULL,
                authorization_json TEXT NOT NULL,
                workbook_identity TEXT NOT NULL,
                task_id TEXT,
                status TEXT NOT NULL CHECK (
                    status IN ('running', 'needs_reconciliation', 'completed', 'failed')
                ),
                next_action_index INTEGER NOT NULL DEFAULT 0 CHECK (next_action_index >= 0),
                in_flight_action_index INTEGER,
                before_snapshot_json TEXT,
                checkpoint_json TEXT,
                result_json TEXT,
                error_json TEXT,
                owner_token TEXT NOT NULL,
                revision INTEGER NOT NULL DEFAULT 1 CHECK (revision >= 1),
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_office_executions_status_updated
                ON office_executions(status, updated_at);
            CREATE INDEX IF NOT EXISTS idx_office_executions_task
                ON office_executions(task_id);
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
        planning_columns = {
            str(row["name"])
            for row in connection.execute(
                "PRAGMA table_info(planning_tasks)"
            ).fetchall()
        }
        if "current_model_id" not in planning_columns:
            connection.execute(
                "ALTER TABLE planning_tasks ADD COLUMN current_model_id TEXT"
            )
        if "model_switches_json" not in planning_columns:
            connection.execute(
                "ALTER TABLE planning_tasks ADD COLUMN model_switches_json "
                "TEXT NOT NULL DEFAULT '{\"items\":[]}'"
            )
        extension_columns = {
            str(row["name"])
            for row in connection.execute(
                "PRAGMA table_info(planning_budget_extensions)"
            ).fetchall()
        }
        if "model_id" not in extension_columns:
            connection.execute(
                "ALTER TABLE planning_budget_extensions ADD COLUMN model_id TEXT"
            )
        office_columns = {
            str(row["name"])
            for row in connection.execute(
                "PRAGMA table_info(office_executions)"
            ).fetchall()
        }
        if "task_id" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN task_id TEXT"
            )
        connection.execute("PRAGMA user_version = %d" % _SCHEMA_VERSION)

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
    def _office_record(cls, row: sqlite3.Row) -> OfficeExecutionRecord:
        plan = cls._decode(str(row["plan_json"]))
        authorization = cls._decode(str(row["authorization_json"]))
        if plan is None or authorization is None:
            raise RuntimeError("Office 执行记录缺少计划或授权现场")
        return OfficeExecutionRecord(
            operation_id=str(row["operation_id"]),
            request_fingerprint=str(row["request_fingerprint"]),
            plan=plan,
            authorization=authorization,
            workbook_identity=str(row["workbook_identity"]),
            task_id=str(row["task_id"]) if row["task_id"] is not None else None,
            status=str(row["status"]),  # type: ignore[arg-type]
            next_action_index=int(row["next_action_index"]),
            in_flight_action_index=(
                int(row["in_flight_action_index"])
                if row["in_flight_action_index"] is not None
                else None
            ),
            before_snapshot=cls._decode(row["before_snapshot_json"]),
            checkpoint=cls._decode(row["checkpoint_json"]),
            result=cls._decode(row["result_json"]),
            error=cls._decode(row["error_json"]),
            owner_token=str(row["owner_token"]),
            revision=int(row["revision"]),
            created_at=float(row["created_at"]),
            updated_at=float(row["updated_at"]),
        )

    @classmethod
    def _planning_record(cls, row: sqlite3.Row) -> PlanningTaskRecord:
        request = cls._decode(str(row["request_json"]))
        if request is None:
            raise RuntimeError("规划任务缺少原始请求")
        raw_switches = cls._decode(row["model_switches_json"]) or {"items": []}
        switch_items = raw_switches.get("items", [])
        if not isinstance(switch_items, list) or not all(
            isinstance(item, dict) for item in switch_items
        ):
            raise RuntimeError("规划任务包含无效的模型切换记录")
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
            current_model_id=(
                str(row["current_model_id"])
                if row["current_model_id"] is not None
                else None
            ),
            model_switches=tuple(dict(item) for item in switch_items),
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
        with self._connect_readonly() as connection:
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
        with self._connect_readonly() as connection:
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
        with self._connect_readonly() as connection:
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
        with self._connect_readonly() as connection:
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
                        approved_wall_time_seconds, current_model_id,
                        executor_id, created_at, updated_at
                    ) VALUES (?, ?, ?, 'running', ?, ?, ?, ?, ?)
                    """,
                    (
                        turn_id,
                        request_fingerprint,
                        encoded_request,
                        approved_wall_time_seconds,
                        (
                            str(request.get("modelId"))
                            if request.get("modelId") is not None
                            else None
                        ),
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
        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT * FROM planning_tasks WHERE turn_id = ?", (turn_id,)
            ).fetchone()
        return self._planning_record(row) if row is not None else None

    def planning_events(self, turn_id: str) -> list[dict[str, Any]]:
        with self._connect_readonly() as connection:
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
        error: dict[str, Any] | None = None,
    ) -> PlanningTaskRecord:
        encoded = self._encode(checkpoint)
        encoded_error = self._encode(error) if error is not None else None
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
                    active_elapsed_seconds = ?, pause_reason = ?, error_json = ?, updated_at = ?,
                    revision = revision + 1
                WHERE turn_id = ? AND status = 'running' AND executor_id = ?
                """,
                (
                    encoded,
                    max(active_elapsed_seconds, 0.0),
                    reason[:1000],
                    encoded_error,
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
        model_id: str | None = None,
        model_switch_reason: str | None = None,
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
                SELECT added_seconds, model_id FROM planning_budget_extensions
                WHERE turn_id = ? AND request_id = ?
                """,
                (turn_id, request_id),
            ).fetchone()
            if duplicate is not None:
                duplicate_model_id = (
                    str(duplicate["model_id"])
                    if duplicate["model_id"] is not None
                    else None
                )
                if duplicate_model_id != model_id:
                    connection.rollback()
                    raise RuntimeError("同一继续请求不能改用另一模型连接")
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
            current_model_id = (
                str(row["current_model_id"])
                if row["current_model_id"] is not None
                else None
            )
            switches = self._decode(row["model_switches_json"]) or {"items": []}
            switch_items = switches.get("items", [])
            if not isinstance(switch_items, list):
                connection.rollback()
                raise RuntimeError("规划任务的模型切换记录已损坏")
            if model_id is not None and model_id != current_model_id:
                switch_items = [
                    *switch_items,
                    {
                        "fromModelId": current_model_id,
                        "toModelId": model_id,
                        "reason": (model_switch_reason or "用户明确选择备用模型")[:500],
                        "createdAt": now,
                    },
                ]
                current_model_id = model_id
            encoded_switches = self._encode({"items": switch_items})
            added = 0.0
            if row["status"] == "waiting_user" and extension_seconds > 0:
                # 2026-10-10 改造：一次"继续"即**结清到总活跃执行时限**，而不是
                # 每次只加 ``extension_seconds``（60 秒）。旧语义下初批预算就等于
                # 提醒阈值（60 秒），若仍按 60 秒递增，用户要点几十次才能跑到真实
                # 任务需要的时长 —— 那正是"被短时预算反复打断"的另一面。
                # 现在一次点击进入新策略：批准预算抬到 hard_wall_time，仍受
                # hard_wall_time 硬封顶，且累计耗时继续累加、不复位。
                already_settled = approved >= hard_wall_time_seconds
                if already_settled and extension_count >= max(maximum_extensions, 0):
                    # 只有**已经**结清到总时限的任务才继续受次数上限约束。
                    # 旧短时预算暂停的任务（approved 仍是 60 秒）不受此限，
                    # 否则它会被历史次数卡住，永远进不了新策略。
                    connection.rollback()
                    raise RuntimeError("规划任务已达到允许的继续次数上限")
                added = max(hard_wall_time_seconds - approved, 0.0)
                if active_elapsed >= hard_wall_time_seconds:
                    connection.rollback()
                    raise RuntimeError("规划任务已达到有效运行总时长上限")
            elif active_elapsed >= hard_wall_time_seconds:
                connection.rollback()
                raise RuntimeError("规划任务已达到有效运行总时长上限")
            connection.execute(
                """
                INSERT INTO planning_budget_extensions (
                    turn_id, request_id, added_seconds, model_id, created_at
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (turn_id, request_id, added, model_id, now),
            )
            changed = connection.execute(
                """
                UPDATE planning_tasks
                SET status = 'running', executor_id = ?, pause_reason = NULL,
                    error_json = NULL, current_model_id = ?, model_switches_json = ?,
                    approved_wall_time_seconds = approved_wall_time_seconds + ?,
                    extension_count = extension_count + ?, updated_at = ?,
                    revision = revision + 1
                WHERE turn_id = ? AND revision = ?
                """,
                (
                    self.executor_id,
                    current_model_id,
                    encoded_switches,
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
        with self._connect_readonly() as connection:
            rows = connection.execute(sql, parameters).fetchall()
        return [self._directory_record(row) for row in rows]

    def managed_directory(self, directory_id: str) -> ManagedDirectoryRecord | None:
        with self._connect_readonly() as connection:
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

    def register_office_execution(
        self,
        operation_id: str,
        request_fingerprint: str,
        *,
        plan: dict[str, Any],
        authorization: dict[str, Any],
        workbook_identity: str,
        identity_prefix: str = "url:",
        task_id: str | None = None,
    ) -> tuple[OfficeExecutionRecord, bool]:
        """Register the durable intent before the first supported host write.

        `identity_prefix` 是宿主身份命名空间：Office 用 ``url:``（绑定
        ``Office.context.document.url``），WPS 用 ``wps-path:``（绑定已保存工作簿的
        小写磁盘完整路径）。两种前缀互不冒充 —— 已存在记录的身份按**完整字符串**
        比较，跨前缀的重放必然命中"属于另一个工作簿"而拒绝，不会借前缀蒙混。
        """

        encoded_plan = self._encode(plan)
        encoded_authorization = self._encode(authorization)
        maximum_size = capability_int("taskRecovery", "maxCheckpointBytes")
        if len(encoded_plan.encode("utf-8")) > maximum_size:
            raise ValueError("执行计划超过可持久化容量上限")
        if len(encoded_authorization.encode("utf-8")) > maximum_size:
            raise ValueError("执行授权现场超过可持久化容量上限")
        if not workbook_identity.startswith(identity_prefix):
            raise ValueError("未保存或无法稳定识别的工作簿不能启用自动断点续跑")
        now = time.time()
        owner_token = uuid.uuid4().hex
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if row is None:
                connection.execute(
                    """
                    INSERT INTO office_executions (
                        operation_id, request_fingerprint, plan_json,
                        authorization_json, workbook_identity, task_id, status,
                        owner_token, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, 'running', ?, ?, ?)
                    """,
                    (
                        operation_id,
                        request_fingerprint,
                        encoded_plan,
                        encoded_authorization,
                        workbook_identity,
                        task_id,
                        owner_token,
                        now,
                        now,
                    ),
                )
                row = connection.execute(
                    "SELECT * FROM office_executions WHERE operation_id = ?",
                    (operation_id,),
                ).fetchone()
                connection.commit()
                assert row is not None
                return self._office_record(row), True
            connection.commit()
        record = self._office_record(row)
        if record.request_fingerprint != request_fingerprint:
            raise RuntimeError("该 operationId 已用于另一份执行请求")
        if record.workbook_identity != workbook_identity:
            raise RuntimeError("该执行属于另一个工作簿")
        if task_id is not None and record.task_id != task_id:
            raise RuntimeError("该 operationId 已绑定到另一个任务")
        return record, False

    def has_execution_for_task(self, task_id: str) -> bool:
        """这个规划任务是否已经登记过写入执行。

        "继续分析"必须先过这一关：只要有执行记录（无论已完成、部分执行还是
        结果未知），就**不能**用重新规划绕过去 —— 那等于在没有核对原表的情况下
        再写一遍。调用方据此把用户导向"继续执行 / 核对"或明确说明需要重新发起。
        """
        if not task_id:
            return False
        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT 1 FROM office_executions WHERE task_id = ? LIMIT 1",
                (task_id,),
            ).fetchone()
        return row is not None

    def office_execution(self, operation_id: str) -> OfficeExecutionRecord | None:
        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        return self._office_record(row) if row is not None else None

    def begin_office_action(
        self,
        operation_id: str,
        *,
        expected_revision: int,
        owner_token: str,
        action_index: int,
        before_snapshot: dict[str, Any],
    ) -> OfficeExecutionRecord:
        encoded_snapshot = self._encode(before_snapshot)
        if len(encoded_snapshot.encode("utf-8")) > capability_int(
            "taskRecovery", "maxCheckpointBytes"
        ):
            raise ValueError("Office 动作写前快照超过容量上限")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE office_executions
                SET in_flight_action_index = ?, before_snapshot_json = ?,
                    error_json = NULL, updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND revision = ? AND owner_token = ?
                    AND status = 'running' AND next_action_index = ?
                    AND in_flight_action_index IS NULL
                """,
                (
                    action_index,
                    encoded_snapshot,
                    now,
                    operation_id,
                    expected_revision,
                    owner_token,
                    action_index,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("Office 动作状态已变化，已停止写入")
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._office_record(row)

    def complete_office_action(
        self,
        operation_id: str,
        *,
        expected_revision: int,
        owner_token: str,
        action_index: int,
        checkpoint: dict[str, Any],
    ) -> OfficeExecutionRecord:
        encoded_checkpoint = self._encode(checkpoint)
        if len(encoded_checkpoint.encode("utf-8")) > capability_int(
            "taskRecovery", "maxCheckpointBytes"
        ):
            raise ValueError("Office 动作检查点超过容量上限")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE office_executions
                SET next_action_index = ?, in_flight_action_index = NULL,
                    before_snapshot_json = NULL, checkpoint_json = ?,
                    error_json = NULL, updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND revision = ? AND owner_token = ?
                    AND status = 'running' AND next_action_index = ?
                    AND in_flight_action_index = ?
                """,
                (
                    action_index + 1,
                    encoded_checkpoint,
                    now,
                    operation_id,
                    expected_revision,
                    owner_token,
                    action_index,
                    action_index,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("Office 动作回执已过期，写入结果需要核对")
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._office_record(row)

    def suspend_office_execution(
        self,
        operation_id: str,
        *,
        expected_revision: int,
        owner_token: str,
        error: dict[str, Any],
    ) -> OfficeExecutionRecord:
        encoded_error = self._encode(error)
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                """
                UPDATE office_executions
                SET status = CASE
                        WHEN in_flight_action_index IS NULL THEN 'failed'
                        ELSE 'needs_reconciliation'
                    END,
                    error_json = ?, updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND revision = ? AND owner_token = ?
                    AND status = 'running'
                """,
                (
                    encoded_error,
                    now,
                    operation_id,
                    expected_revision,
                    owner_token,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("Office 执行状态已变化，无法重复挂起")
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._office_record(row)

    def claim_office_resume(
        self,
        operation_id: str,
        *,
        expected_revision: int,
        require_idle: bool = False,
    ) -> OfficeExecutionRecord:
        """Rotate the owner token; an old task pane can no longer commit receipts."""

        now = time.time()
        owner_token = uuid.uuid4().hex
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            if require_idle:
                self._assert_wps_idle(connection, operation_id, now)
            changed = connection.execute(
                """
                UPDATE office_executions
                SET status = 'running', owner_token = ?, error_json = NULL,
                    updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND revision = ?
                    AND status IN ('running', 'failed', 'needs_reconciliation')
                    AND in_flight_action_index IS NULL
                """,
                (owner_token, now, operation_id, expected_revision),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("Office 任务需要先核对在途动作，或状态已经变化")
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._office_record(row)

    def reconcile_office_action(
        self,
        operation_id: str,
        *,
        expected_revision: int,
        action_index: int,
        outcome: Literal["applied", "not_applied"],
        checkpoint: dict[str, Any] | None = None,
        require_idle: bool = False,
    ) -> OfficeExecutionRecord:
        if outcome == "applied" and checkpoint is None:
            raise ValueError("补记已应用动作必须提供完整检查点")
        encoded_checkpoint = self._encode(checkpoint) if checkpoint is not None else None
        if encoded_checkpoint and len(encoded_checkpoint.encode("utf-8")) > capability_int(
            "taskRecovery", "maxCheckpointBytes"
        ):
            raise ValueError("Office 核对检查点超过容量上限")
        now = time.time()
        owner_token = uuid.uuid4().hex
        next_action_index = action_index + 1 if outcome == "applied" else action_index
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            if require_idle:
                self._assert_wps_idle(connection, operation_id, now)
            changed = connection.execute(
                """
                UPDATE office_executions
                SET status = 'running', next_action_index = ?,
                    in_flight_action_index = NULL, before_snapshot_json = NULL,
                    checkpoint_json = COALESCE(?, checkpoint_json),
                    error_json = NULL, owner_token = ?, updated_at = ?,
                    revision = revision + 1
                WHERE operation_id = ? AND revision = ?
                    AND status IN ('running', 'needs_reconciliation')
                    AND next_action_index = ? AND in_flight_action_index = ?
                """,
                (
                    next_action_index,
                    encoded_checkpoint,
                    owner_token,
                    now,
                    operation_id,
                    expected_revision,
                    action_index,
                    action_index,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("Office 在途动作状态已变化，无法补记")
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert row is not None
        return self._office_record(row)

    @staticmethod
    def _assert_wps_idle(connection: sqlite3.Connection, operation_id: str, now: float) -> None:
        row = connection.execute(
            "SELECT workbook_identity, status, updated_at FROM office_executions WHERE operation_id = ?",
            (operation_id,),
        ).fetchone()
        if row is None or not str(row["workbook_identity"]).startswith("wps-path:"):
            raise RuntimeError("未找到这次 WPS 执行记录")
        if row["status"] == "running" and float(row["updated_at"]) + capability_int(
            "taskRecovery", "wpsOwnerLeaseSeconds"
        ) > now:
            raise RuntimeError("原执行者仍持有 WPS 执行权，请等待或由原执行者挂起后再核对")

    def finish_office_execution(
        self,
        operation_id: str,
        *,
        expected_revision: int,
        owner_token: str,
        result: dict[str, Any],
    ) -> OfficeExecutionRecord:
        encoded_result = self._encode(result)
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT plan_json FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if row is None:
                connection.rollback()
                raise RuntimeError("未找到 Office 执行记录")
            plan = self._decode(str(row["plan_json"])) or {}
            action_count = len(plan.get("actions", []))
            changed = connection.execute(
                """
                UPDATE office_executions
                SET status = 'completed', result_json = ?, error_json = NULL,
                    updated_at = ?, revision = revision + 1
                WHERE operation_id = ? AND revision = ? AND owner_token = ?
                    AND status = 'running' AND next_action_index = ?
                    AND in_flight_action_index IS NULL
                """,
                (
                    encoded_result,
                    now,
                    operation_id,
                    expected_revision,
                    owner_token,
                    action_count,
                ),
            ).rowcount
            if changed != 1:
                connection.rollback()
                raise RuntimeError("Office 任务尚未完成全部动作或完成回执已过期")
            completed = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        assert completed is not None
        return self._office_record(completed)

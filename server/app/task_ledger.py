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
_SCHEMA_VERSION = 13
# `office_executions` 上"后加的列"清单：`_schema_is_current` 必须逐个确认它们**真的存在**。
# 版本号只是意图，列才是事实 —— 版本号被写成最新而列没建时，早退会让补列永久跳过
# （2026-10-27 真机：恢复包 PUT 因此 500）。新增列必须同时进这里与迁移。
_REQUIRED_OFFICE_COLUMNS = (
    "recovery_package_expires_at",
    "discovery_dismissed_at",
    "abandoned_at",
    "undone_at",
    "recovery_package_digest",
    "recovery_package_bytes",
    "recovery_package_retired_at",
    "recovery_package_removed_at",
)
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

    **版本号不等于列存在**：只有版本号够新、且所有必填列都在，才算就绪。旧库的
    `user_version` 可能已经被写成最新而列还没建（升级中途失败、或升级前就在的库），
    只看版本号会让补列永久跳过 —— 2026-10-27 真机的 `no such column` 就是这么来的。

    库不存在、不是 SQLite 库或读取失败 → 返回 False（走完整 bootstrap，
    由 bootstrap 自己失败关闭）。
    """
    if not path.exists():
        return False
    try:
        with sqlite3.connect(path, timeout=5) as connection:
            row = connection.execute("PRAGMA user_version").fetchone()
            columns = {str(column[1]) for column in connection.execute("PRAGMA table_info(office_executions)")}
    except sqlite3.Error:
        return False
    return (
        bool(row)
        and int(row[0]) >= _SCHEMA_VERSION
        and all(column in columns for column in _REQUIRED_OFFICE_COLUMNS)
    )


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
    #: 用户终态处置的首次时刻（`abandoned_at` = 放弃变更，`undone_at` = 撤销本次执行）。
    #: `None` 只表示"没有这条留痕"，不代表"用户没处置过"（上线前的处置没有留痕）。
    abandoned_at: float | None = None
    undone_at: float | None = None


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
                -- 恢复材料的**首次**到期时间（ISO 文本）。只写一次，之后只读：
                -- 重复 PUT、读回水合、保留期清理都不许改写这个值，否则"幂等保存"
                -- 就变成了免费的续期按钮。这里只存时间戳，不存任何单元格正文。
                recovery_package_expires_at TEXT,
                -- 「恢复提示已被用户清除」的处置时间（epoch 秒）。**只**用来抑制发现提示：
                -- 它不推进 revision、不改 status、不动动作游标、不删恢复材料 —— "清除了提示"
                -- 与"任务已经结束"是两件不同的事实。`dismissed_at >= updated_at` 时该记录
                -- 不再出现在发现结果里；执行一旦有新变化（updated_at 前进）就重新可见。
                discovery_dismissed_at REAL,
                -- 用户对这次执行的**处置留痕**：各自记"第一次发生"的时刻（epoch 秒）。
                -- `abandoned_at` = 审阅阶段点了「放弃变更」；`undone_at` = 确认之后点了
                -- 「撤销本次执行」。两列都**只写一次**（首次写赢）、都不推进 `revision`、
                -- 不改 `status`、不动动作游标、不删恢复材料 —— 它们是"用户做过什么"的
                -- 审计事实，不是进度。`NULL` 只有一个含义：**没有这条记录**（不是"没做过"）。
                abandoned_at REAL,
                undone_at REAL,
                -- 「被回收的那份恢复材料是哪一份」的留痕（U5，2026-10-30）。
                -- **只存指纹与字节数，不存材料正文**：恢复包里含单元格原值，账本里一个值
                -- 都不许有（`test_recovery_material_never_enters_the_task_ledger` 钉着这条）。
                -- 前两列 + `retired_at` **一起**写、**只写一次**（首次写赢）；`NULL` = 没有这条
                -- 留痕，不是"材料没被回收过"。它回答的只有一件事：删除那一刻磁盘上那堆字节
                -- 是什么 —— 拿它跟任何一份留底比，就能确认"当时删的是不是这一份"。
                recovery_package_digest TEXT,
                recovery_package_bytes INTEGER,
                recovery_package_retired_at REAL,
                -- **文件确已删除**的时刻（epoch 秒）。`NULL` = **没有确认**删除（可能是删除
                -- 报错、也可能探测到有材料却删到时已不在）。
                -- 关键区别：`retired_at` 只说明"我记下了要删的是哪一份"，**不能**当成删除成功；
                -- 只有 `retired_at` 与 `removed_at` 都有，才能说"那一份确实被删掉了"。
                -- 单独一列而不是复用 `retired_at`：落账成功、删除失败时两者必须能分开读。
                recovery_package_removed_at REAL,
                owner_token TEXT NOT NULL,
                revision INTEGER NOT NULL DEFAULT 1 CHECK (revision >= 1),
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_office_executions_status_updated
                ON office_executions(status, updated_at);
            CREATE INDEX IF NOT EXISTS idx_office_executions_task
                ON office_executions(task_id);
            -- 「实际写入了几次」的审计：只存**计数与状态**，不存单元格值、公式正文
            -- 或工具参数。追加式（同一次动作可以先后记两条：写成功一条、补记一条），
            -- 汇总时按 `(operation_id, action_index)` 取最后一条。
            CREATE TABLE IF NOT EXISTS office_host_writes (
                operation_id TEXT NOT NULL,
                action_index INTEGER NOT NULL CHECK (action_index >= 0),
                payload_json TEXT NOT NULL,
                recorded_at REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_office_host_writes_operation
                ON office_host_writes(operation_id, action_index, recorded_at);
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
        if "recovery_package_expires_at" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN recovery_package_expires_at TEXT"
            )
        if "discovery_dismissed_at" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN discovery_dismissed_at REAL"
            )
        if "abandoned_at" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN abandoned_at REAL"
            )
        if "undone_at" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN undone_at REAL"
            )
        # U5：旧库补三列。**列才是事实**，版本号只是意图 —— 三列必须逐个补，
        # 漏一列就会让"删除前的留痕"在旧库上静默变成 no-op。
        if "recovery_package_digest" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN recovery_package_digest TEXT"
            )
        if "recovery_package_bytes" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN recovery_package_bytes INTEGER"
            )
        if "recovery_package_retired_at" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN recovery_package_retired_at REAL"
            )
        if "recovery_package_removed_at" not in office_columns:
            connection.execute(
                "ALTER TABLE office_executions ADD COLUMN recovery_package_removed_at REAL"
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
            abandoned_at=(
                float(row["abandoned_at"]) if row["abandoned_at"] is not None else None
            ),
            undone_at=float(row["undone_at"]) if row["undone_at"] is not None else None,
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

    def reserve_recovery_package_expiry(self, operation_id: str, expires_at: str) -> str | None:
        """登记一次执行的恢复材料**首次**到期时间，返回最终生效的值。

        语义（"首次写赢"）：

        - 该执行不存在 ⇒ 返回 `None`（调用方按"未登记"处理，不许自己造一个期限）；
        - 已登记 ⇒ 原样返回**第一次**写入的值，绝不被后来的调用推后；
        - 未登记 ⇒ 只在这时写入 `expires_at`。

        为什么放在账本：保留期清掉磁盘材料之后，仍然要能回答"这次执行的恢复材料本来
        什么时候到期"，否则删掉文件再 PUT 就能白拿一个新的 7 天。这里只存时间戳，
        不含任何计划或单元格正文，也不推进 `revision`（它不是执行进度）。
        """

        with self._connect() as connection:
            row = connection.execute(
                "SELECT recovery_package_expires_at FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if row is None:
                return None
            existing = row["recovery_package_expires_at"]
            if existing is not None:
                return str(existing)
            connection.execute(
                """
                UPDATE office_executions SET recovery_package_expires_at = ?
                WHERE operation_id = ? AND recovery_package_expires_at IS NULL
                """,
                (expires_at, operation_id),
            )
            connection.commit()
            row = connection.execute(
                "SELECT recovery_package_expires_at FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        value = None if row is None else row["recovery_package_expires_at"]
        return str(value) if value is not None else None

    def recovery_package_expiry(self, operation_id: str) -> str | None:
        """只读：已登记的首次到期时间（没有就 `None`）。"""

        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT recovery_package_expires_at FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        value = None if row is None else row["recovery_package_expires_at"]
        return str(value) if value is not None else None

    def office_execution(self, operation_id: str) -> OfficeExecutionRecord | None:
        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT * FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        return self._office_record(row) if row is not None else None

    def pending_office_executions(
        self, workbook_identity: str, *, limit: int
    ) -> tuple[list[OfficeExecutionRecord], bool]:
        """按**完整相等**的工作簿身份找出未完成执行（只读）。

        这是"重新发现未完成执行"的唯一后端入口：浏览器里的发现索引只是缓存，丢了也能
        从这里找回来。三条纪律钉在 SQL 里：

        - `workbook_identity = ?` **完整串相等**：不按文件名、不按前缀、不按"最近一条"；
        - `status IN ('running','needs_reconciliation','failed')`：**已完成不是待续写**；
        - `discovery_dismissed_at IS NULL OR < updated_at`：用户清掉的提示不再复活，但
          执行一旦有新变化（`updated_at` 前进）就重新可见。处置列不参与进度判断。

        排序把"有在途动作的"放最前、"写前失败（一个动作都没推进）"放最后 —— 一条刚失败
        的写前记录不许盖住真正可能已落表的那条。返回 `(records, truncated)`，
        `truncated` 由多取一条判定，**不得**把截断静默当成"只有这些"。
        """

        maximum = capability_int("taskRecovery", "pendingDiscoveryMaxCandidates")
        effective = max(1, min(int(limit), maximum))
        with self._connect_readonly() as connection:
            rows = connection.execute(
                """
                SELECT * FROM office_executions
                WHERE workbook_identity = ?
                    AND status IN ('running', 'needs_reconciliation', 'failed')
                    AND (discovery_dismissed_at IS NULL OR discovery_dismissed_at < updated_at)
                ORDER BY
                    CASE
                        WHEN in_flight_action_index IS NOT NULL
                            OR status = 'needs_reconciliation' THEN 0
                        WHEN status = 'failed'
                            AND next_action_index = 0
                            AND in_flight_action_index IS NULL THEN 2
                        ELSE 1
                    END,
                    updated_at DESC,
                    operation_id ASC
                LIMIT ?
                """,
                (workbook_identity, effective + 1),
            ).fetchall()
        records = [self._office_record(row) for row in rows[:effective]]
        return records, len(rows) > effective

    def dismiss_office_execution_discovery(self, operation_id: str) -> bool:
        """记下"用户已核对并清除了这条发现提示"（幂等）。

        **只写一列时间戳**：不改 `status`、不推进 `revision`、不动动作游标与在途动作、
        不删恢复材料。拿它冒充"任务已完成"或"没写过"，会让下一次真正的续跑失去依据。
        返回该执行是否存在（不存在就是 `False`，调用方不得事后补建）。
        """

        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            changed = connection.execute(
                "UPDATE office_executions SET discovery_dismissed_at = ? WHERE operation_id = ?",
                (now, operation_id),
            ).rowcount
            connection.commit()
        return changed == 1

    #: 用户处置的两条合法取值 → `office_executions` 上对应的列名。
    #: 不放进数据库 CHECK：SQLite 的 `ADD COLUMN` 加不了带 CHECK 的列，两处各写一份
    #: 约束只会在迁移后不一致。合法取值只在这里判一次。
    DISPOSITION_COLUMNS = {"abandoned": "abandoned_at", "undone": "undone_at"}

    def record_office_disposition(
        self, operation_id: str, disposition: str
    ) -> tuple[bool, float | None]:
        """记下用户对这次执行做过的终态处置（**首次写赢**）。

        只写一列时间戳，别的一概不做：

        - **不推进 `revision`**、不改 `status`、不动 `next/in_flight`、不删恢复材料 ——
          "用户放弃了/撤销了"不是执行进度，写进进度会让续跑判据错位；
        - **重复上报不改写历史**：这一列回答的是"第一次发生在什么时候"，重试、双端各报一次、
          用户再点一次，都不该把那一刻往后退（与 `reserve_recovery_package_expiry` 同一口径）；
        - 记录不存在就返回 `(False, None)`：**绝不**为了留痕补建一条账本。

        返回 `(记录是否存在, 该处置第一次被记下的时刻)`。
        """

        column = self.DISPOSITION_COLUMNS.get(disposition)
        if column is None:
            raise ValueError("未知的执行处置类型")
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            exists = connection.execute(
                "SELECT 1 FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if exists is None:
                connection.rollback()
                return False, None
            connection.execute(
                f"UPDATE office_executions SET {column} = ? "
                f"WHERE operation_id = ? AND {column} IS NULL",
                (time.time(), operation_id),
            )
            row = connection.execute(
                f"SELECT {column} FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        recorded = None if row is None else row[column]
        return True, (float(recorded) if recorded is not None else None)

    def office_disposition(self, operation_id: str) -> dict[str, float | None] | None:
        """只读：这次执行被放弃/撤销的**首次时刻**（没有记录就是 `None`）。

        `NULL` 的含义必须读作"**没有这条留痕**"：本功能上线前发生过的放弃/撤销，
        以及从未做过处置的执行，看起来完全一样。**不要**把 `None` 说成"用户没放弃过"。
        """

        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT abandoned_at, undone_at FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        if row is None:
            return None
        return {
            "abandonedAt": None if row["abandoned_at"] is None else float(row["abandoned_at"]),
            "undoneAt": None if row["undone_at"] is None else float(row["undone_at"]),
        }

    def record_recovery_package_retirement(
        self, operation_id: str, digest: str, size: int
    ) -> tuple[bool, dict[str, Any] | None]:
        """记下「这次执行的恢复材料被交出回收」（U5，**删除前**写、**首次写赢**）。

        为什么需要它：材料一删就没得算了。删掉之后任何"当时那份材料是不是这一份"的回答
        都只能是推断 —— 而终态回收恰恰是唯一合法删除点（原地撤销成功），验收时最想知道
        的正是"消失的是哪一份"。

        **和 `record_office_disposition` 的两处差别，不许照抄**：

        - **时机相反**：处置留痕必须等宿主恢复**成功之后**才写（没恢复就不许说"撤销了"）；
          这一条必须在删除**之前**写 —— 它记的是"即将被删掉的那堆字节"，删完只剩猜。
        - **失败不许被吞掉**：调用方（DELETE 端点）拿不到留痕就**不许删**。否则"删了但
          没记录"正是这一列要防的那种状态。上报方（端上 `retireWpsRecoveryPackage`）本来
          就吞错，于是失败的表现是"材料留着、等保留期到期"，不是"材料没了还没证据"。

        三列（指纹/字节数/时刻）**同一次 UPDATE 一起写**：拆成三次会出现"有指纹没时刻"
        这种半条记录。`NULL` = 没有留痕，不是"材料没被回收过"。
        """

        fingerprint = (digest or "").strip().lower()
        if len(fingerprint) != 64 or any(ch not in "0123456789abcdef" for ch in fingerprint):
            raise ValueError("恢复材料指纹必须是 64 位十六进制摘要")
        if not isinstance(size, int) or isinstance(size, bool) or size < 0:
            raise ValueError("恢复材料字节数必须是非负整数")
        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            exists = connection.execute(
                "SELECT 1 FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if exists is None:
                connection.rollback()
                return False, None
            connection.execute(
                "UPDATE office_executions "
                "SET recovery_package_digest = ?, recovery_package_bytes = ?, "
                "    recovery_package_retired_at = ? "
                "WHERE operation_id = ? AND recovery_package_retired_at IS NULL",
                (fingerprint, int(size), now, operation_id),
            )
            row = connection.execute(
                "SELECT recovery_package_digest, recovery_package_bytes, "
                "       recovery_package_retired_at, recovery_package_removed_at "
                "FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        if row is None or row["recovery_package_retired_at"] is None:
            # 记录在、但这一列仍空：只可能是并发下 UPDATE 没命中（不会发生），
            # 或者从来没写过。两种都读作"没有留痕"，不编一个时刻出来。
            return True, None
        return True, self._retirement_row(row)

    def mark_recovery_package_removed(
        self, operation_id: str
    ) -> tuple[bool, dict[str, Any] | None]:
        """记下「文件**确实**删掉了」（U5 的第二个事实，首次写赢）。

        为什么必须单独一列，不能复用 `retired_at`：落账成功、删除失败是完全可能的状态
        （磁盘占用、权限、进程还在读）。那时 `retired_at` 已经有值，如果拿它当"删掉了"，
        读取方就会得到一条**假的删除结论**。所以：

        - `retired_at` = "我记下了要删的是这一份"（删前摘要，**只是记录**）；
        - `removed_at` = "`remove()` 报告真的删掉了"（**确认删除**）；
        - 只有两者都有，才能说"那一份确实被删掉了"。

        调用点**只**在 `remove()` 返回 True 之后。`remove()` 返回 False（探测到有、删到时
        已不在）或抛错，都不写这一列 —— "不是删掉的"就不许记成删掉的。
        """

        now = time.time()
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            exists = connection.execute(
                "SELECT 1 FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            if exists is None:
                connection.rollback()
                return False, None
            connection.execute(
                "UPDATE office_executions SET recovery_package_removed_at = ? "
                "WHERE operation_id = ? AND recovery_package_removed_at IS NULL",
                (now, operation_id),
            )
            row = connection.execute(
                "SELECT recovery_package_digest, recovery_package_bytes, "
                "       recovery_package_retired_at, recovery_package_removed_at "
                "FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
            connection.commit()
        if row is None:
            return True, None
        return True, self._retirement_row(row)

    @staticmethod
    def _retirement_row(row: Any) -> dict[str, Any] | None:
        """把留痕那一行折成对外形状；`retired_at` 为空就是"没有留痕"。"""

        if row["recovery_package_retired_at"] is None:
            return None
        return {
            "digest": row["recovery_package_digest"],
            "bytes": None if row["recovery_package_bytes"] is None else int(row["recovery_package_bytes"]),
            "retiredAt": float(row["recovery_package_retired_at"]),
            "removedAt": None if row["recovery_package_removed_at"] is None
            else float(row["recovery_package_removed_at"]),
        }

    def recovery_package_retirement(self, operation_id: str) -> dict[str, Any] | None:
        """只读：这份恢复材料被回收时的指纹与两个时刻（没有留痕就是 `None`）。

        `None` 必须读作"**没有这条留痕**"：本功能上线前删掉的材料、以及从未回收过的
        执行，看起来完全一样。**不要**把 `None` 说成"材料还在"。

        `removedAt is None` 必须读作"**没有确认删除**"：不能因为 `retiredAt` 有值就宣称
        删掉了 —— 那正是这条留痕最容易读错的地方。
        """

        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT recovery_package_digest, recovery_package_bytes, "
                "       recovery_package_retired_at, recovery_package_removed_at "
                "FROM office_executions WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        if row is None:
            return None
        return self._retirement_row(row)

    def append_office_host_writes(
        self, operation_id: str, entries: list[dict[str, Any]]
    ) -> None:
        """追加一次宿主写调用审计（**只追加**，不改执行状态、不推进 revision）。

        为什么不是"更新一行"：这份审计是观测，不是恢复点。把它写成 upsert 会让
        晚到的补记覆盖先到的现场；写成"必须成功"又会给关键写入路径加一个失败点
        —— 调用方（端上）本来就是 best-effort，失败只该少一条证据，不该影响写入。
        """
        if not entries:
            return
        now = time.time()
        rows = [
            (
                operation_id,
                int(entry.get("actionIndex") or 0),
                json.dumps(entry, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
                now,
            )
            for entry in entries
        ]
        with self._connect() as connection:
            connection.executemany(
                """
                INSERT INTO office_host_writes (
                    operation_id, action_index, payload_json, recorded_at
                ) VALUES (?, ?, ?, ?)
                """,
                rows,
            )
            connection.commit()

    def office_host_writes(self, operation_id: str) -> list[dict[str, Any]]:
        """按动作序号取**每个动作最后一条**审计（同动作多条时以最后一条为准）。"""
        with self._connect_readonly() as connection:
            rows = connection.execute(
                """
                SELECT action_index, payload_json FROM office_host_writes
                WHERE operation_id = ?
                ORDER BY recorded_at ASC, rowid ASC
                """,
                (operation_id,),
            ).fetchall()
        latest: dict[int, dict[str, Any]] = {}
        for row in rows:
            try:
                payload = json.loads(str(row["payload_json"]))
            except (TypeError, ValueError):
                continue
            if isinstance(payload, dict):
                latest[int(row["action_index"])] = payload
        return [latest[key] for key in sorted(latest)]

    def office_host_writes_last_recorded_at(self, operation_id: str) -> float | None:
        """最近一次收到上报的时刻；一条都没有时返回 ``None``。

        **不要**把它当成 0 或"现在"：`None` 就是"还没收到"，这正是"未采集"与
        "明确零"的分界依据之一。
        """
        with self._connect_readonly() as connection:
            row = connection.execute(
                "SELECT MAX(recorded_at) FROM office_host_writes WHERE operation_id = ?",
                (operation_id,),
            ).fetchone()
        value = row[0] if row is not None else None
        return float(value) if value is not None else None

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

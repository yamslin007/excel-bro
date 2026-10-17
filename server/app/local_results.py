"""Volatile, task-scoped registry of local query results.

The value matrix of a local query never leaves the trusted local side: the model
only ever receives an opaque ``resultRef``. This registry is the single place
that stores the real values, bounded by ``config/capabilities.json``
(``agentToolBridge``).

Three lifecycles, modelled separately
-------------------------------------
* **Authorization** — an :class:`OwnerClaim` records which host / workbook /
  authorized-scope a task was planned under. Re-authorization, a workbook change
  or a different session bumps the claim revision and drops the retained results.
* **Session** — the volatile tool bridge session. It ends when the turn returns.
* **Result** — registered values outlive the session (the terminal result card is
  rendered *after* the turn returns) but stay bounded by their own TTL, the
  per-task cap and the global retention cap.

Boundaries kept here:

* No persistence: the registry is in-process memory only. It is deliberately
  separate from the durable execution ledger and from chat history.
* No model exposure: ``resolve`` is only called by trusted server code (result
  compilation / local result card assembly), never on a path that serializes
  into a model prompt.
"""

from __future__ import annotations

import hashlib
import json
import math
import secrets
import threading
import time
from collections import OrderedDict
from dataclasses import dataclass, replace
from typing import Any

from .capabilities import capability_int
from .models import CellValue, DataToolResult


def fallback_workbook_identity(workbook_name: str) -> str:
    """Identity used when the client does not supply a stable one.

    Kept explicit (and single-sourced) so that the bridge binding and the
    registry claim can never disagree about what "this workbook" means.
    """
    return f"name:{workbook_name}"


@dataclass(frozen=True)
class OwnerClaim:
    """Who owns what, for one planning task. Revision bumps on re-authorization."""

    task_id: str
    host_id: str
    workbook_name: str
    workbook_identity: str
    authorized_key: str
    source_fingerprint: str | None
    revision: int
    created_at: float

    def same_authorization(self, other: "OwnerClaim") -> bool:
        return (
            self.host_id == other.host_id
            and self.workbook_name == other.workbook_name
            and self.workbook_identity == other.workbook_identity
            and self.authorized_key == other.authorized_key
        )

    def accepts(
        self,
        *,
        host_id: str,
        workbook_name: str,
        workbook_identity: str | None,
    ) -> bool:
        identity = workbook_identity or fallback_workbook_identity(workbook_name)
        return (
            self.host_id == host_id
            and self.workbook_name == workbook_name
            and self.workbook_identity == identity
        )


@dataclass
class RegisteredResult:
    handle: str
    task_id: str
    title: str
    headers: list[CellValue]
    rows: list[list[CellValue]]
    source_sheets: list[str]
    scanned_rows: int
    complete: bool
    calculation: str
    warnings: list[str]
    source_fingerprint: str | None
    created_at: float
    fingerprint: str
    # 登记时的授权版本。授权一变，旧结果立刻不可解析（不靠 TTL 兜底）。
    claim_revision: int | None = None

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def column_count(self) -> int:
        return len(self.headers)


def _encode_value(value: CellValue) -> str:
    return json.dumps(value, ensure_ascii=False, allow_nan=False, separators=(",", ":"))


def _validate_matrix(result: DataToolResult) -> None:
    """Independent bounds check.

    Do not trust Pydantic field limits alone: a crafted result must not be able
    to materialize an oversized matrix in the trusted registry.
    """
    max_rows = capability_int("agentToolBridge", "maxResultRows")
    max_columns = capability_int("agentToolBridge", "maxResultColumns")
    max_cells = capability_int("agentToolBridge", "maxResultCells")
    max_chars = capability_int("agentToolBridge", "maxResultCellChars")

    rows = len(result.rows)
    columns = len(result.headers)
    if not 1 <= columns <= max_columns or not 0 <= rows <= max_rows:
        raise ValueError()
    if rows * columns > max_cells:
        raise ValueError()
    for row in result.rows:
        if len(row) != columns:
            raise ValueError()
        for value in row:
            if value is None or isinstance(value, bool):
                continue
            if isinstance(value, int):
                continue
            if isinstance(value, float) and math.isfinite(value):
                continue
            if isinstance(value, str) and len(value) <= max_chars:
                continue
            raise ValueError()


class LocalResultRegistry:
    def __init__(self) -> None:
        self._by_handle: OrderedDict[str, RegisteredResult] = OrderedDict()
        self._by_task: dict[str, list[str]] = {}
        self._claims: dict[str, OwnerClaim] = {}
        #  Registry writes happen on the worker thread that runs the isolated
        # container (see python_coordinator.execute_python_analysis_async) while
        # reads happen on the event loop. Reentrant so internal helpers
        # (_prune -> discard/release) can nest without deadlocking.
        self._lock = threading.RLock()

    # -- lifecycle: authorization -------------------------------------------

    def bind(
        self,
        *,
        task_id: str,
        host_id: str,
        workbook_name: str,
        workbook_identity: str | None,
        authorized_key: str,
        source_fingerprint: str | None,
    ) -> OwnerClaim:
        """Record the authorization a task is currently planned under.

        A change in host / workbook / authorized scope invalidates every result
        previously retained for that task: those values belong to an
        authorization that no longer applies.
        """
        with self._lock:
            return self._bind_locked(
                task_id=task_id,
                host_id=host_id,
                workbook_name=workbook_name,
                workbook_identity=workbook_identity,
                authorized_key=authorized_key,
                source_fingerprint=source_fingerprint,
            )

    def _bind_locked(self, *, task_id, host_id, workbook_name, workbook_identity,
                     authorized_key, source_fingerprint) -> OwnerClaim:
        identity = workbook_identity or fallback_workbook_identity(workbook_name)
        existing = self._claims.get(task_id)
        revision = 1
        if existing is not None:
            candidate = replace(
                existing,
                host_id=host_id,
                workbook_name=workbook_name,
                workbook_identity=identity,
                authorized_key=authorized_key,
                source_fingerprint=source_fingerprint,
            )
            if existing.same_authorization(candidate):
                return existing
            self._discard_task_locked(task_id)
            revision = existing.revision + 1
        self._prune()
        claim = OwnerClaim(
            task_id=task_id,
            host_id=host_id,
            workbook_name=workbook_name,
            workbook_identity=identity,
            authorized_key=authorized_key,
            source_fingerprint=source_fingerprint,
            revision=revision,
            created_at=time.monotonic(),
        )
        self._claims[task_id] = claim
        return claim

    def claim(self, task_id: str) -> OwnerClaim | None:
        with self._lock:
            self._prune()
            return self._claims.get(task_id)

    def claim_accepts(
        self,
        task_id: str,
        *,
        host_id: str,
        workbook_name: str,
        workbook_identity: str | None,
    ) -> bool:
        claim = self.claim(task_id)
        if claim is None:
            return False
        return claim.accepts(
            host_id=host_id,
            workbook_name=workbook_name,
            workbook_identity=workbook_identity,
        )

    # -- lifecycle: results --------------------------------------------------

    def _prune(self) -> None:
        with self._lock:
            self._prune_locked()

    def _prune_locked(self) -> None:
        ttl = capability_int("agentToolBridge", "resultTtlSeconds")
        cutoff = time.monotonic() - ttl
        expired = [
            handle
            for handle, item in self._by_handle.items()
            if item.created_at < cutoff
        ]
        for handle in expired:
            self._discard_locked(handle)
        for task_id, claim in list(self._claims.items()):
            if claim.created_at < cutoff:
                self._discard_task_locked(task_id)
                self._claims.pop(task_id, None)
        # 全局保留上限：结果生命周期长于会话，必须有独立容量边界。
        max_retained = capability_int("agentToolBridge", "maxRetainedResults")
        while len(self._by_handle) > max_retained:
            handle, item = self._by_handle.popitem(last=False)
            handles = self._by_task.get(item.task_id)
            if handles and handle in handles:
                handles.remove(handle)
                if not handles:
                    self._by_task.pop(item.task_id, None)

    def register(
        self, task_id: str, result: DataToolResult, source_fingerprint: str | None
    ) -> RegisteredResult:
        with self._lock:
            self._prune_locked()
            return self._register_locked(task_id, result, source_fingerprint)

    def _register_locked(self, task_id: str, result: DataToolResult,
                         source_fingerprint: str | None) -> RegisteredResult:
        _validate_matrix(result)
        handles = self._by_task.setdefault(task_id, [])
        if len(handles) >= capability_int("agentToolBridge", "maxResultsPerTask"):
            raise ValueError("本任务登记的结果过多")

        claim = self._claims.get(task_id)
        payload = result.model_dump(mode="json")
        digest = hashlib.sha256(
            json.dumps(
                payload, ensure_ascii=False, allow_nan=False, sort_keys=True
            ).encode("utf-8")
        ).hexdigest()
        handle = f"res_{secrets.token_hex(16)}"
        registered = RegisteredResult(
            handle=handle,
            task_id=task_id,
            title=result.title,
            headers=list(result.headers),
            rows=[list(row) for row in result.rows],
            source_sheets=list(result.sourceSheets),
            scanned_rows=result.scannedRows,
            complete=result.complete,
            calculation=result.calculation,
            warnings=list(result.warnings),
            source_fingerprint=source_fingerprint,
            created_at=time.monotonic(),
            fingerprint=digest,
            claim_revision=claim.revision if claim is not None else None,
        )
        self._by_handle[handle] = registered
        handles.append(handle)
        return registered

    def resolve(self, task_id: str, handle: str) -> RegisteredResult | None:
        with self._lock:
            self._prune_locked()
            return self._resolve_locked(task_id, handle)

    def _resolve_locked(self, task_id: str, handle: str) -> RegisteredResult | None:
        item = self._by_handle.get(handle)
        if item is None or item.task_id != task_id:
            # Knowing the string is not enough: ownership is checked here.
            return None
        claim = self._claims.get(task_id)
        if (
            claim is not None
            and item.claim_revision is not None
            and item.claim_revision != claim.revision
        ):
            # 登记时的授权已经被替换 —— 旧值不得再被读出。
            return None
        return item

    def resolve_owned(
        self,
        task_id: str,
        handle: str,
        *,
        host_id: str,
        workbook_name: str,
        workbook_identity: str | None,
    ) -> RegisteredResult | None:
        """Resolve for the loopback result endpoint.

        Requires the caller to still hold the task's authorization; the session
        may already be closed (the result card is rendered after the turn).
        """
        if not self.claim_accepts(
            task_id,
            host_id=host_id,
            workbook_name=workbook_name,
            workbook_identity=workbook_identity,
        ):
            return None
        return self.resolve(task_id, handle)

    def discard(self, handle: str) -> None:
        with self._lock:
            self._discard_locked(handle)

    def _discard_locked(self, handle: str) -> None:
        item = self._by_handle.pop(handle, None)
        if item is None:
            return
        handles = self._by_task.get(item.task_id)
        if handles and handle in handles:
            handles.remove(handle)
            if not handles:
                self._by_task.pop(item.task_id, None)

    def discard_task(self, task_id: str) -> None:
        with self._lock:
            self._discard_task_locked(task_id)

    def _discard_task_locked(self, task_id: str) -> None:
        for handle in list(self._by_task.get(task_id, [])):
            self._by_handle.pop(handle, None)
        self._by_task.pop(task_id, None)

    def release(self, task_id: str) -> None:
        """Full teardown: results *and* the ownership claim."""
        with self._lock:
            self._discard_task_locked(task_id)
            self._claims.pop(task_id, None)

    def invalidate_source(self, task_id: str, source_fingerprint: str | None) -> int:
        """Drop results whose bound source no longer matches. Returns count."""
        with self._lock:
            removed = 0
            for handle in list(self._by_task.get(task_id, [])):
                item = self._by_handle.get(handle)
                if item is not None and item.source_fingerprint != source_fingerprint:
                    self._discard_locked(handle)
                    removed += 1
            return removed

    def count_task(self, task_id: str) -> int:
        with self._lock:
            return len(self._by_task.get(task_id, []))

    def count(self) -> int:
        with self._lock:
            return len(self._by_handle)

    def reset(self) -> None:
        with self._lock:
            self._by_handle.clear()
            self._by_task.clear()
            self._claims.clear()


local_results = LocalResultRegistry()


def structure_of(item: RegisteredResult, *, truncated: bool = False) -> dict[str, Any]:
    """The only shape of a result that may be shown to the model."""
    return {
        "headers": [
            header if isinstance(header, str) else str(header)
            for header in item.headers
        ],
        "rowCount": item.row_count,
        "columnCount": item.column_count,
        "complete": bool(item.complete) and not truncated,
        "truncated": bool(truncated),
    }

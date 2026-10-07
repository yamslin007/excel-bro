"""Task-scoped, bounded, volatile data. Never model content or a persistent store.

The host adapter proves the live workbook identity before and after reading.
This registry checks exact bindings, not the truth of a caller's Excel claim.
No paths, logs, disk spill, checkpoint hooks or Python execution in this module.
"""
from __future__ import annotations

import hashlib
import json
import math
import secrets
import threading
import time
from dataclasses import dataclass, field

from .capabilities import capability_int
from .column_mapping import _range_box
from .models import PythonDataScope, PythonDatasetReceipt, PythonResultReceipt
from .python_results import encode_matrix_result


class DatasetRejected(ValueError):
    def __init__(self):
        super().__init__("授权数据不可用，请重新确认目标并生成预览。")


@dataclass(repr=False)
class _Grant:
    scope: dict
    origin: str
    expires: float
    datasets: dict[str, tuple[bytes, str, int]] = field(default_factory=dict)
    results: dict[str, "_Result"] = field(default_factory=dict)


@dataclass(repr=False, frozen=True)
class _Result:
    encoded: bytes
    # Order matters: code receives datasets in the caller's declared order.
    datasets: tuple[tuple[str, str], ...]
    receipt: PythonResultReceipt


@dataclass
class _Budget:
    calls: int
    expires: float


class PythonDatasetRegistry:
    def __init__(self, *, clock=time.monotonic):
        self._clock = clock
        self._grants: dict[str, _Grant] = {}
        self._budgets: dict[tuple[str, str], _Budget] = {}
        self._lock = threading.RLock()

    def _prune(self):
        now = self._clock()
        for token in list(self._grants):
            if self._grants[token].expires <= now:
                del self._grants[token]
        for key in list(self._budgets):
            if self._budgets[key].expires <= now:
                del self._budgets[key]

    def _grant(self, token: str, scope: PythonDataScope, origin: str):
        self._prune()
        grant = self._grants.get(token)
        if not grant or grant.scope != scope.model_dump() or grant.origin != origin:
            raise DatasetRejected()
        return grant

    def create(self, scope: PythonDataScope, origin: str) -> str:
        with self._lock:
            self._prune()
            sources = scope.sources
            limit = capability_int("pythonData", "maxDatasetsPerScope")
            if len(sources) > limit or len({(s.sheetId, s.range) for s in sources}) != len(sources):
                raise DatasetRejected()
            for source in sources:
                r1, c1, r2, c2 = _range_box(source.range)
                if r2 > 1048576 or c2 > 16384 or (r2-r1+1)*(c2-c1+1) > capability_int("pythonData", "maxCellsPerDataset"):
                    raise DatasetRejected()
            key = origin, scope.taskId
            budget = self._budgets.get(key)
            if budget is None:
                if len(self._budgets) >= capability_int("pythonData", "maxScopes"):
                    raise DatasetRejected()
                budget = _Budget(0, self._clock() + capability_int("pythonData", "scopeTtlSeconds"))
            # A replacement authorization for the same local task atomically
            # invalidates all old handles, even when names look unchanged.
            for token, grant in list(self._grants.items()):
                if grant.origin == origin and grant.scope["taskId"] == scope.taskId:
                    del self._grants[token]
            if len(self._grants) >= capability_int("pythonData", "maxScopes"):
                raise DatasetRejected()
            token = secrets.token_urlsafe(32)
            self._budgets[key] = budget
            self._grants[token] = _Grant(scope.model_dump(), origin, budget.expires)
            return token

    def register(self, token: str, scope: PythonDataScope, origin: str,
                 index: int, values: list[list]) -> PythonDatasetReceipt:
        with self._lock:
            grant = self._grant(token, scope, origin)
            if type(index) is not int or not 0 <= index < len(scope.sources):
                raise DatasetRejected()
            r1, c1, r2, c2 = _range_box(scope.sources[index].range)
            if type(values) is not list or len(values) != r2-r1+1:
                raise DatasetRejected()
            for row in values:
                if type(row) is not list or len(row) != c2-c1+1:
                    raise DatasetRejected()
                for value in row:
                    if value is not None and type(value) not in (str, int, float, bool):
                        raise DatasetRejected()
                    if type(value) is float and not math.isfinite(value):
                        raise DatasetRejected()
            try:
                encoded = json.dumps(values, ensure_ascii=False, allow_nan=False,
                    separators=(",", ":")).encode("utf-8")
            except (ValueError, UnicodeError):
                raise DatasetRejected() from None
            total = sum(len(data[0]) for g in self._grants.values() for data in g.datasets.values())
            if (len(encoded) > capability_int("pythonData", "maxDatasetBytes")
                or total + len(encoded) > capability_int("pythonData", "maxRegistryBytes")):
                raise DatasetRejected()
            # One snapshot per exact source. Re-registration invalidates its
            # previous handle rather than retaining stale revisions.
            for handle, data in list(grant.datasets.items()):
                if data[2] == index:
                    del grant.datasets[handle]
                    for result_handle, result in list(grant.results.items()):
                        if any(bound[0] == handle for bound in result.datasets):
                            del grant.results[result_handle]
            handle = "dataset://" + secrets.token_urlsafe(32)
            fingerprint = hashlib.sha256(encoded).hexdigest()
            grant.datasets[handle] = encoded, fingerprint, index
            return PythonDatasetReceipt(datasetHandle=handle, contentFingerprint=fingerprint)

    def resolve(self, token: str, scope: PythonDataScope, origin: str,
                receipts: list[PythonDatasetReceipt], *, charge=True) -> dict:
        with self._lock:
            grant = self._grant(token, scope, origin)
            if not receipts or len(receipts) > capability_int("pythonData", "maxDatasetsPerScope"):
                raise DatasetRejected()
            if len({r.datasetHandle for r in receipts}) != len(receipts):
                raise DatasetRejected()
            datasets = []
            total_bytes = 0
            for receipt in receipts:
                data = grant.datasets.get(receipt.datasetHandle)
                if not data or data[1] != receipt.contentFingerprint:
                    raise DatasetRejected()
                total_bytes += len(data[0])
                if total_bytes > capability_int("pythonData", "maxAnalysisInputBytes"):
                    raise DatasetRejected()
                datasets.append({"source": grant.scope["sources"][data[2]].copy(),
                                 "values": json.loads(data[0])})
            if charge:
                budget = self._budgets[(grant.origin, grant.scope["taskId"])]
                if budget.calls >= capability_int("pythonExecution", "maxCallsPerTask"):
                    raise DatasetRejected()
                budget.calls += 1
            return {"datasets": datasets}

    def revoke(self, token: str, scope: PythonDataScope, origin: str):
        with self._lock:
            self._grant(token, scope, origin)
            del self._grants[token]

    def valid(self, token: str, scope: PythonDataScope, origin: str,
              receipts: list[PythonDatasetReceipt]) -> bool:
        """Cheap cancellation/publication guard; never deserializes cell data."""
        with self._lock:
            try:
                grant = self._grant(token, scope, origin)
            except DatasetRejected:
                return False
            return (0 < len(receipts) <= capability_int("pythonData", "maxDatasetsPerScope")
                and len({r.datasetHandle for r in receipts}) == len(receipts) and all(
                r.datasetHandle in grant.datasets and grant.datasets[r.datasetHandle][1] == r.contentFingerprint
                for r in receipts))

    def publish_result(self, token: str, scope: PythonDataScope, origin: str,
                       receipts: list[PythonDatasetReceipt], result) -> PythonResultReceipt:
        with self._lock:
            grant = self._grant(token, scope, origin)
            if not self.valid(token, scope, origin, receipts):
                raise DatasetRejected()
            try:
                encoded, rows, columns = encode_matrix_result(result)
            except (ValueError, TypeError, UnicodeError):
                raise DatasetRejected() from None
            total = sum(len(r.encoded) for g in self._grants.values() for r in g.results.values())
            if (len(grant.results) >= capability_int("pythonResult", "maxResultsPerScope")
                or total + len(encoded) > capability_int("pythonResult", "maxRegistryBytes")):
                raise DatasetRejected()
            # Recheck expiry after potentially expensive untrusted validation.
            grant = self._grant(token, scope, origin)
            handle = "result://" + secrets.token_urlsafe(32)
            receipt = PythonResultReceipt(resultHandle=handle,
                contentFingerprint=hashlib.sha256(encoded).hexdigest(), rows=rows, columns=columns)
            grant.results[handle] = _Result(encoded,
                tuple((r.datasetHandle, r.contentFingerprint) for r in receipts), receipt.model_copy())
            return receipt

    def _result(self, token, scope, origin, receipts, receipt):
        grant = self._grant(token, scope, origin)
        result = grant.results.get(receipt.resultHandle)
        if (not result or not self.valid(token, scope, origin, receipts)
            or result.datasets != tuple((r.datasetHandle, r.contentFingerprint) for r in receipts)
            or result.receipt != receipt):
            raise DatasetRejected()
        return grant, result

    def read_result(self, token, scope, origin, receipts, receipt) -> dict:
        with self._lock:
            _, result = self._result(token, scope, origin, receipts, receipt)
            return json.loads(result.encoded)

    def discard_result(self, token, scope, origin, receipts, receipt):
        with self._lock:
            grant, _ = self._result(token, scope, origin, receipts, receipt)
            del grant.results[receipt.resultHandle]


registry = PythonDatasetRegistry()

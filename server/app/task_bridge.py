"""Volatile task tool bridge: SDK tool handler ⇄ local host adapter.

Why a bridge exists
-------------------
The Claude Agent SDK runs in the Python backend; Office.js / WPS read the
*open, possibly unsaved* workbook inside the task pane. A Python MCP handler
therefore cannot read the workbook itself. The bridge lets an SDK tool handler
submit a query request that is handed to the host adapter owning the task; the
adapter executes the existing deterministic ``query_table`` locally and posts a
receipt back over a loopback endpoint.

Ownership and safety
--------------------
Every call is bound at request time to (taskId, hostId, workbook identity,
source fingerprint, authorized-sheet key, operation id and cancel generation).
A model- or client-reported id never grants access on its own: every delivery is
matched against the live binding, and a receipt that arrives after cancel /
supersession / source change is rejected as late instead of silently accepted.

The tool request channel is deliberately separate from the durable progress
ledger: the progress stream only carries an opaque pointer (``bridgeTool``),
while the request arguments live only in this volatile store and are fetched by
the host adapter over the loopback poll endpoint.
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import logging
import secrets
import time
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Awaitable, Callable

from .capabilities import capability_int
from . import dev_fault_target, dev_result_ref_fault, dev_source_changed_fault
from .local_results import fallback_workbook_identity, local_results
from .planning_source_guard import (
    SOURCE_CHANGED_RECEIPT_CODE,
    is_source_changed_receipt,
)
from .python_coordinator import coordinator
from .python_datasets import registry
from .models import (
    BridgeDatasetReceipt,
    DataToolResult,
    PlanRequest,
    PythonDataScope,
    PythonDataSource,
)

TERMINAL_CANCEL_ERROR_CODES = frozenset(
    {"cancelled", "timeout", "budget_exhausted", "source_changed", "host_unavailable"}
)

logger = logging.getLogger(__name__)


def bridge_step_id(tool: str, operation_id: str) -> str:
    """进度事件里"一次工具调用"的步骤身份。

    身份必须包含**这一次调用**，不能只有工具名：同一个工具连续调两次是两个步骤。
    只用工具名时，第二次调用的 `started` 会覆盖第一次的步骤，界面上"读了两次"
    看起来只有一次，而且第一次永远不会被标成完成。

    重放同一次调用时 `operation_id` 不变，所以事件仍然落到同一个步骤上（只更新、
    不新增）；换一次调用 `operation_id` 就是新的，于是独立成步。这正是"重试身份"
    与"单次调用身份"的分界。

    后缀取 12 位十六进制（48 bit）：足以区分同一任务内的所有调用，
    又不会把步骤 ID 撑到影响日志可读性。分隔符用 `:`——协议对 `stepId` 的约束是
    `^[A-Za-z0-9_.:-]+$`（`server/app/models.py`），`#` 会被校验直接拒掉。
    """
    suffix = operation_id.removeprefix("op_")[:12] or "unknown"
    return f"bridge.{tool}:{suffix}"


class BridgeError(RuntimeError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def _source_key(source_fingerprint: str | None) -> str | None:
    return source_fingerprint


@dataclass(frozen=True)
class BridgeBinding:
    task_id: str
    host_id: str
    workbook_name: str
    source_fingerprint: str | None
    authorized_key: str
    workbook_identity: str = ""

    @property
    def identity(self) -> str:
        return self.workbook_identity or fallback_workbook_identity(self.workbook_name)

    @classmethod
    def from_request(cls, request: PlanRequest, *, task_id: str) -> "BridgeBinding":
        names = sorted(sheet.name for sheet in request.workbook.worksheets)
        authorized_key = hashlib.sha256(
            json.dumps(
                {"sheets": names, "fingerprint": request.workbook.sourceFingerprint},
                ensure_ascii=False,
                sort_keys=True,
            ).encode("utf-8")
        ).hexdigest()
        return cls(
            task_id=task_id,
            host_id=request.hostId or "excel-office",
            workbook_name=request.workbook.name,
            source_fingerprint=_source_key(request.workbook.sourceFingerprint),
            authorized_key=authorized_key,
            workbook_identity=(
                request.workbook.workbookIdentity
                or fallback_workbook_identity(request.workbook.name)
            ),
        )

    def matches(
        self,
        *,
        host_id: str,
        workbook_name: str,
        workbook_identity: str | None = None,
    ) -> bool:
        # Identity is host instance + workbook *identity*, not just the display
        # name: two different workbooks can share a name. The source fingerprint
        # is deliberately *not* compared here — a differing fingerprint is a
        # legitimate signal that the source changed while the task ran, handled
        # in ``deliver`` as ``source_changed``.
        echoed = workbook_identity or fallback_workbook_identity(workbook_name)
        return (
            self.host_id == host_id
            and self.workbook_name == workbook_name
            and self.identity == echoed
        )


def arguments_fingerprint(tool: str, arguments: dict[str, Any]) -> str:
    payload = json.dumps(
        {"tool": tool, "arguments": arguments},
        ensure_ascii=False,
        allow_nan=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


@dataclass
class _Pending:
    operation_id: str
    tool: str
    arguments: dict[str, Any]
    fingerprint: str
    future: "asyncio.Future[dict[str, Any]]"
    created_at: float


@dataclass
class _Completed:
    fingerprint: str
    receipt: dict[str, Any]
    recorded_at: float = field(default_factory=time.monotonic)


class ToolBridgeSession:
    """One bridge session per planning task; serial, single-consumer."""

    def __init__(
        self,
        binding: BridgeBinding,
        *,
        on_request: Callable[[str, str, dict[str, Any], str], Awaitable[None]] | None,
        on_complete: (
            Callable[[str, str, dict[str, Any]], Awaitable[None]] | None
        ) = None,
    ) -> None:
        self.binding = binding
        self.on_request = on_request
        # 一次桥调用的**真实结束**（成功 / 固定错误码 / 超时 / 取消）都在这里出口。
        # 以前没有任何"结束"回调，于是进度事件只有 started、没有 completed，
        # 前端只能靠"下一步出现了"倒推上一步完成 —— 那是对用户的猜测。
        self.on_complete = on_complete
        self.generation = 0
        self.cancel_reason: str | None = None
        # 来源变化是**本次规划的终态**：一经生产守卫确认，之后的工具调用不再派给
        # 宿主（也就不会再有读取），并且回同一个固定码，让模型知道原因不是"用户取消"。
        self.source_stopped = False
        self.call_count = 0
        self.total_read_cells = 0
        self.created_at = time.monotonic()
        self.updated_at = self.created_at
        self._pending: OrderedDict[str, _Pending] = OrderedDict()
        self._completed: OrderedDict[str, _Completed] = OrderedDict()

    # -- introspection -----------------------------------------------------

    @property
    def pending_calls(self) -> list[dict[str, Any]]:
        now = time.monotonic()
        ttl = capability_int("agentToolBridge", "operationTtlSeconds")
        calls = []
        for item in self._pending.values():
            calls.append(
                {
                    "operationId": item.operation_id,
                    "tool": item.tool,
                    "arguments": item.arguments,
                    "requestFingerprint": item.fingerprint,
                    "expiresAt": item.created_at + ttl,
                }
            )
        return calls

    def _prune_completed(self) -> None:
        ttl = capability_int("agentToolBridge", "operationTtlSeconds")
        cutoff = time.monotonic() - ttl
        stale = [
            key
            for key, item in self._completed.items()
            if item.recorded_at < cutoff
        ]
        for key in stale:
            self._completed.pop(key, None)

    # -- request path ------------------------------------------------------

    async def request(self, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
        if self.cancel_reason is not None:
            # 来源变化不是"用户取消"：回同一个固定码，且**不**再产生一次宿主读表。
            # 回 `cancelled` 会让模型以为是用户停的，进而以为"再试一次也许可以"。
            code = (
                SOURCE_CHANGED_RECEIPT_CODE
                if self.cancel_reason == SOURCE_CHANGED_RECEIPT_CODE
                else "cancelled"
            )
            return await self._notify_complete(tool, "op_none", _error_receipt(code))
        if dev_source_changed_fault.take(self.binding.task_id):
            # 开发验收专用（`EXCEL_BRO_SOURCE_CHANGED_FAULT=1` 时才可能为真）：
            # 只有 `open()` 绑定到**这一次任务**的启用才会命中（`bind_task`），
            # 另一本簿 / 同一个簿的另一个任务照常派给宿主。
            # 让"来源变化"在真实面板上成为一次**真实发生**的终止，而不是靠
            # "希望模型再读一次"。回执与生产完全同一份固定码，且**不派发**给
            # 宿主 —— 顺序也与生产一致：先交回这次回执，再停桥。
            receipt = _error_receipt(SOURCE_CHANGED_RECEIPT_CODE)
            await self._notify_complete(tool, "op_none", receipt)
            self.stop_for_source_change()
            return receipt
        max_calls = capability_int("agentToolBridge", "maxToolCallsPerTask")
        if self.call_count >= max_calls:
            return await self._notify_complete(
                tool, "op_none", _error_receipt("internal", extra_code="budget_exhausted")
            )
        if len(self._pending) >= capability_int("agentToolBridge", "maxPendingPerTask"):
            return await self._notify_complete(
                tool, "op_none", _error_receipt("internal", extra_code="busy")
            )

        fingerprint = arguments_fingerprint(tool, arguments)
        operation_id = f"op_{secrets.token_hex(16)}"
        loop = asyncio.get_running_loop()
        future: "asyncio.Future[dict[str, Any]]" = loop.create_future()
        pending = _Pending(
            operation_id=operation_id,
            tool=tool,
            arguments=arguments,
            fingerprint=fingerprint,
            future=future,
            created_at=time.monotonic(),
        )
        self._pending[operation_id] = pending
        self.call_count += 1
        self.updated_at = time.monotonic()

        if self.on_request is not None:
            await self.on_request(operation_id, tool, arguments, fingerprint)

        timeout = float(capability_int("agentToolBridge", "waitTimeoutSeconds"))
        try:
            return await self._notify_complete(
                tool, operation_id, await asyncio.wait_for(future, timeout=timeout)
            )
        except TimeoutError:
            self._pending.pop(operation_id, None)
            # 超时也是一个**真实终态**：这一步确实没做成，不能留在"进行中"。
            return await self._notify_complete(
                tool, operation_id, _error_receipt("timeout")
            )
        finally:
            self._pending.pop(operation_id, None)

    async def _notify_complete(
        self, tool: str, operation_id: str, receipt: dict[str, Any]
    ) -> dict[str, Any]:
        if self.on_complete is not None:
            try:
                await self.on_complete(operation_id, tool, receipt)
            except Exception:
                # 进度事件发不出去只影响展示，绝不能让它把工具回执吞掉。
                logger.warning("桥调用结束事件发送失败 tool=%s", tool, exc_info=True)
        return receipt

    # -- delivery path -----------------------------------------------------

    def deliver(
        self,
        *,
        operation_id: str,
        request_fingerprint: str,
        host_id: str,
        workbook_name: str,
        source_fingerprint: str | None,
        result: DataToolResult | None,
        error_code: str | None,
        workbook_identity: str | None = None,
        datasets: list[BridgeDatasetReceipt] | None = None,
    ) -> dict[str, Any]:
        """Validate and accept a host receipt. Raises :class:`BridgeError`.

        Idempotency: the same operation id with the same request fingerprint is
        accepted once and returns ``duplicate`` afterwards; the same id with a
        different fingerprint is a conflict. Receipts for unknown ids (cancelled,
        superseded, expired) are rejected as late.
        """
        if not self.binding.matches(
            host_id=host_id,
            workbook_name=workbook_name,
            workbook_identity=workbook_identity,
        ):
            raise BridgeError("stale_binding", "回执与本次任务绑定不一致")

        if self.cancel_reason is not None:
            raise BridgeError("cancelled", "本次任务已取消，回执不再接收")

        self._prune_completed()
        completed = self._completed.get(operation_id)
        if completed is not None:
            if completed.fingerprint != request_fingerprint:
                raise BridgeError("conflict", "同一操作 ID 收到了不同参数的回执")
            return {"status": "duplicate"}

        pending = self._pending.get(operation_id)
        if pending is None:
            raise BridgeError("late", "该操作已不在等待中，回执被拒绝")
        if pending.fingerprint != request_fingerprint:
            raise BridgeError("conflict", "回执参数指纹与请求不一致")

        if error_code is not None:
            receipt = _error_receipt(error_code)
        elif pending.tool == "prepare_python_data":
            receipt = self._deliver_python_datasets(
                pending, datasets, source_fingerprint
            )
        elif result is None:
            receipt = _error_receipt("internal")
        else:
            if (
                self.binding.source_fingerprint is not None
                and source_fingerprint is not None
                and source_fingerprint != self.binding.source_fingerprint
            ):
                # Source changed after the turn was bound: refuse to bind values
                # read from a different snapshot.
                receipt = _error_receipt("source_changed")
            else:
                try:
                    registered = local_results.register(
                        self.binding.task_id, result, source_fingerprint
                    )
                except ValueError:
                    receipt = _error_receipt("result_too_large")
                else:
                    receipt = {
                        "status": "ok",
                        "resultRef": registered.handle,
                        "structure": {
                            "headers": list(registered.headers),
                            "rowCount": registered.row_count,
                            "columnCount": registered.column_count,
                            "complete": bool(registered.complete),
                            "truncated": not registered.complete,
                        },
                    }

        self._pending.pop(operation_id, None)
        self._completed[operation_id] = _Completed(
            fingerprint=request_fingerprint, receipt=receipt
        )
        self.updated_at = time.monotonic()
        if not pending.future.done():
            pending.future.set_result(receipt)
        if is_source_changed_receipt(receipt):
            # 这一次调用拿到的是**本次规划的最后一次**宿主读取结果：来源已经不成立，
            # 之后任何再次读取、重新计算或提交都只能建立在旧版本数据上。
            # 必须在交回这次的 `source_changed` 之后再停 —— 顺序反了，模型收到的
            # 就会是 `cancelled`，看不出真正原因。
            self.stop_for_source_change()
        return {"status": "accepted"}

    def _deliver_python_datasets(
        self,
        pending: _Pending,
        datasets: list[BridgeDatasetReceipt] | None,
        source_fingerprint: str | None,
    ) -> dict[str, Any]:
        """``prepare_python_data`` 回执：登记原始值矩阵并建立协调器绑定。

        与 ``query_table`` 同一条来源闸：源指纹变化即拒绝。值只进易失
        :mod:`python_datasets` registry，句柄（``dataset://``）交回模型，
        fingerprint 从不外发 —— 模型自报句柄不能当授权凭据。
        """
        if (
            self.binding.source_fingerprint is not None
            and source_fingerprint is not None
            and source_fingerprint != self.binding.source_fingerprint
        ):
            return _error_receipt("source_changed")
        if datasets is None:
            return _error_receipt("internal")

        # 惰性导入避免模块级循环（python_coordinator 依赖 local_results/models）。
        from .python_coordinator import coordinator
        from .python_datasets import DatasetRejected, registry

        scope = self._build_python_scope(pending)
        if scope is None:
            return _error_receipt("invalid_arguments")

        origin = _bridge_origin(self.binding.host_id)
        # 完整性闸：**声明了几张来源，就必须正好交回几块**，`sourceIndex` 必须是
        # `0..n-1` 的**完整且不重复**集合。
        #
        # 少一块就放行 = 隔离 Python 在"少读了一张表"的数据上算出一份看起来成功
        # 的错答案；重复/越界则会让同一份值被登记两次或挂到错误的来源上。
        # 所以一律拒绝，不补齐、不猜测。
        indexes = [block.sourceIndex for block in datasets]
        if (
            len(datasets) != len(scope.sources)
            or len(set(indexes)) != len(indexes)
            or set(indexes) != set(range(len(scope.sources)))
        ):
            return _error_receipt("invalid_arguments")

        try:
            token = registry.create(scope, origin)
        except DatasetRejected:
            return _error_receipt("result_too_large")
        except Exception:
            return _error_receipt("internal")
        try:
            receipts = []
            source_indexes: dict[str, int] = {}
            for block in datasets:
                receipt = registry.register(
                    token, scope, origin, block.sourceIndex, block.values
                )
                receipts.append(receipt)
                # 回执块可能乱序交回：句柄 -> 来源下标要跟着登记顺序记下来，
                # 不能拿 receipts 的迭代顺序当来源顺序。
                source_indexes[receipt.datasetHandle] = block.sourceIndex
        except DatasetRejected:
            # 形状/类型/字节/限额任一不过 → 统一拒绝，不泄露哪一项不过；
            # 本次已经登记进去的块必须**当场清理**，不能留一份残缺的登记。
            _discard_python_grant(registry, token, scope, origin)
            return _error_receipt("result_too_large")
        except Exception:
            _discard_python_grant(registry, token, scope, origin)
            return _error_receipt("internal")

        coordinator.bind(
            task_id=self.binding.task_id,
            host_id=self.binding.host_id,
            workbook_name=self.binding.workbook_name,
            scope=scope,
            scope_token=token,
            origin=origin,
            receipts=receipts,
            # 本次数据登记时的源快照：之后结果登记要拿它标注"结果来自哪个源"，
            # 否则源一变就无法判断已登记结果算不算过期。
            source_fingerprint=source_fingerprint,
            source_indexes=source_indexes,
        )
        return {
            "status": "ok",
            "datasetHandles": [r.datasetHandle for r in receipts],
        }

    def _build_python_scope(self, pending: _Pending) -> PythonDataScope | None:
        """从 ``prepare_python_data`` 参数构造可信 ``PythonDataScope``。

        ``taskId``/``hostId``/``workbookIdentityHash``/``authorizationVersion`` 全部
        取自服务端 binding（模型拿不到），只有 ``sources`` 来自模型参数，但每个
        source 都要能被授权范围核实 —— 这里只做形状构造，范围核对由宿主读表
        前的授权复核兜底，登记再由 :mod:`python_datasets` 复核形状/类型/上界。
        """
        arguments = pending.arguments if isinstance(pending.arguments, dict) else {}
        sources_arg = arguments.get("sources")
        if not isinstance(sources_arg, list) or not sources_arg:
            return None
        if self.binding.host_id not in ("excel-office", "wps-et"):
            # folder 宿主没有隔离 Python 读表通道。
            return None
        sources: list[PythonDataSource] = []
        for item in sources_arg:
            if not isinstance(item, dict):
                return None
            sheet = item.get("sheet")
            range_ = item.get("range")
            if not isinstance(sheet, str) or not sheet.strip():
                return None
            if not isinstance(range_, str) or not range_.strip():
                return None
            try:
                sources.append(
                    PythonDataSource(
                        sheetId=sheet,
                        sheetName=sheet,
                        range=range_,
                    )
                )
            except Exception:
                return None
        if not sources:
            return None
        identity = self.binding.identity
        try:
            return PythonDataScope(
                version=1,
                taskId=self.binding.task_id,
                hostId=self.binding.host_id,
                workbookIdentityHash=hashlib.sha256(
                    identity.encode("utf-8")
                ).hexdigest(),
                authorizationVersion=self.binding.authorized_key,
                sources=sources,
            )
        except Exception:
            return None

    # -- lifecycle ---------------------------------------------------------

    def cancel(self, reason: str) -> None:
        self.cancel_reason = reason
        for pending in list(self._pending.values()):
            if not pending.future.done():
                pending.future.set_result(_error_receipt("cancelled"))
        self._pending.clear()
        self.updated_at = time.monotonic()

    def stop_for_source_change(self) -> None:
        """生产守卫确认来源变化：本次规划到此为止。

        做三件事，缺一件都会留出"无效重试"的空间：

        1. 在飞调用失败（不能等它们慢慢超时）；
        2. 会话进入**停止**态，后续 ``request`` 直接回固定码，**不再派给宿主**
           —— 否则模型每重试一次，宿主就真读一次来源表，读到的仍是新数据，
           而守卫每次都拒，白白烧掉模型轮数直到 ``MODEL_TURN_LIMIT``；
        3. 作废本任务的 Python 数据登记、协调器绑定与已登记结果
           —— 旧结果不能再被提交成计划。

        与 ``cancel`` 的区别只有一个：停止原因不可被后续动作改写，也不允许
        重新派发。重复调用是幂等的（同一次规划里可能有多条在飞调用）。
        """
        self.source_stopped = True
        self.cancel(SOURCE_CHANGED_RECEIPT_CODE)

    def supersede(self) -> None:
        """Invalidate in-flight calls after re-authorization / turn restart."""
        self.generation += 1
        self.cancel("superseded")


class TaskToolBridge:
    def __init__(self) -> None:
        self._sessions: OrderedDict[str, ToolBridgeSession] = OrderedDict()

    def _prune(self) -> None:
        ttl = capability_int("agentToolBridge", "operationTtlSeconds")
        cutoff = time.monotonic() - ttl
        expired = [
            task_id
            for task_id, session in self._sessions.items()
            if session.updated_at < cutoff
        ]
        for task_id in expired:
            # 只回收会话；结果有独立 TTL，由登记器自己管。
            self.close(task_id)
        max_sessions = capability_int("agentToolBridge", "maxSessions")
        while len(self._sessions) > max_sessions:
            task_id, _ = self._sessions.popitem(last=False)
            local_results.release(task_id)

    def open(
        self,
        binding: BridgeBinding,
        *,
        on_request: Callable[[str, str, dict[str, Any], str], Awaitable[None]] | None = None,
        on_complete: (
            Callable[[str, str, dict[str, Any]], Awaitable[None]] | None
        ) = None,
    ) -> ToolBridgeSession:
        self._prune()
        existing = self._sessions.get(binding.task_id)
        if existing is not None and (
            existing.binding.authorized_key != binding.authorized_key
            or existing.binding.host_id != binding.host_id
            or existing.binding.workbook_name != binding.workbook_name
            or existing.binding.identity != binding.identity
        ):
            # Re-authorization / workbook change: drop stale in-flight work.
            existing.supersede()
            # 同一任务在**新授权/新工作簿**下重开：旧的 Python 数据登记必须当场
            # 作废 —— 否则旧 ``dataset://`` 句柄仍指向旧工作簿/旧范围的值，
            # 隔离执行会拿它算出一份"看起来成功"的错结果。
            drop_python_state(binding.task_id)
        # 授权生命周期：绑定新授权即按版本作废旧结果（host/工作簿/范围任一变化）。
        local_results.bind(
            task_id=binding.task_id,
            host_id=binding.host_id,
            workbook_name=binding.workbook_name,
            workbook_identity=binding.identity,
            authorized_key=binding.authorized_key,
            source_fingerprint=binding.source_fingerprint,
        )
        # 开发验收入口（A1-E1）：此刻才知道"哪一次规划任务"，把用户的显式启用
        # 绑定到这个 task_id。另一本簿 / 另一个宿主的任务不会消费它。
        # 判据是可核验身份相等（`dev_fault_target`），任务侧只报名字时不绑。
        dev_result_ref_fault.bind_task(
            task_id=binding.task_id,
            host_id=binding.host_id,
            workbook_name=binding.workbook_name,
            workbook_identity=binding.identity,
        )
        # 开发验收入口（A1-E3）：此刻才知道"哪一次规划任务"，把用户的显式启用
        # 绑定到这个 task_id。归属放在这里而不是"下一次调用"—— 否则先说话的那本
        # 簿（可能是并发任务、另一本打开的簿）会把注入抢走：被抢的那轮莫名停在
        # 半路，而真正要验的簿一次都看不到停止。
        dev_source_changed_fault.bind_task(
            task_id=binding.task_id,
            host_id=binding.host_id,
            workbook_name=binding.workbook_name,
            workbook_identity=binding.identity,
        )
        # 两个开发入口共用一份"服务端见过哪本簿"的证据（`dev_fault_target`）：
        # 用户只报名字时，靠它判断这个名字下是不是只有一个可核验身份。只在开发
        # 开关打开时记录 —— 生产不留下任何额外状态。
        if dev_source_changed_fault.enabled() or dev_result_ref_fault.enabled():
            dev_fault_target.note_target(
                host_id=binding.host_id,
                workbook_name=binding.workbook_name,
                workbook_identity=binding.identity,
            )
        session = ToolBridgeSession(
            binding, on_request=on_request, on_complete=on_complete
        )
        self._sessions[binding.task_id] = session
        self._sessions.move_to_end(binding.task_id)
        return session

    def session(self, task_id: str) -> ToolBridgeSession | None:
        return self._sessions.get(task_id)

    def close(self, task_id: str) -> None:
        """End the *session*. Retained results keep living until their own TTL.

        The terminal result card is rendered after the turn returns, so the
        result must outlive the session; dropping it here is what made the
        terminal ``resultRef`` unresolvable.

        但**数据登记**必须一起作废：单元格值只在"用户刚授权、刚核对过源"的那一
        轮有意义，会话一结束就没有正当理由继续留在内存里。
        """
        session = self._sessions.pop(task_id, None)
        if session is not None:
            session.supersede()
        drop_python_state(task_id, keep_results=True)

    def cancel(self, task_id: str, reason: str = "cancelled") -> None:
        """取消一次任务：在飞调用失败、Python 绑定与数据登记当场作废。

        与 `close` 的区别：`close` 是正常收尾（结果卡还要用 `resultRef`），
        `cancel` 是用户/服务端主动中止 —— **迟到结果一律不登记**。
        """
        session = self._sessions.get(task_id)
        if session is not None:
            session.cancel(reason)
        drop_python_state(task_id)

    def stop_for_source_change(self, task_id: str) -> None:
        """来源变化终态：会话停止派发 + 本次规划的全部中间态作废。

        作废的是**基于旧来源快照**算出来的东西：数据集登记、协调器绑定、
        已登记结果。留着任何一项，模型都还有机会把旧结果提交成白名单计划。
        """
        session = self._sessions.get(task_id)
        if session is not None:
            session.stop_for_source_change()
        drop_python_state(task_id)

    def release(self, task_id: str) -> None:
        """End the session *and* drop the retained results and ownership."""
        self.close(task_id)
        local_results.release(task_id)

    def count(self) -> int:
        return len(self._sessions)

    def reset(self) -> None:
        for session in self._sessions.values():
            session.cancel("reset")
        for task_id in list(self._sessions):
            drop_python_state(task_id)
        self._sessions.clear()
        local_results.reset()


def _discard_python_grant(registry, token: str, scope, origin: str) -> None:
    """登记中途失败时把**本次**建出来的 grant 整块丢掉。

    不清理会留下一份"登记了一半"的数据：grant 还在、数据集残缺，协调器也没
    绑定 —— 下一次同任务重新登记时 `registry.create` 会把它顶掉，但在这之前
    那些残缺的单元格值一直留在内存里，没有任何正当理由。
    """
    try:
        registry.revoke(token, scope, origin)
    except Exception:
        registry._grants.pop(token, None)


def drop_python_state(task_id: str, *, keep_results: bool = False) -> None:
    """取消 / 重授权 / 换簿 / 源变化 / 会话结束时，把该任务的 Python 状态当场作废。

    要删的是**全部**中间态，缺一处都会留出可乘之机：

    - **数据集登记**（``registry``）：句柄指向的是某一轮授权下读到的**单元格值**。
      留着它，取消之后跑完的那次隔离执行仍能把迟到结果登记出去。
    - **任务绑定**（``coordinator``）：``run_python_analysis`` 靠它反查指纹。
      留着它，模型自报的旧句柄在换簿之后仍能通过形状校验。
    - **已登记结果**（``local_results``）：结果值同样是"在那一轮授权/那一份源快照
      下算出来的"。只在重绑时 bump 授权版本是不够的 —— **源变化**（同一授权、
      同工作簿，但表内容变了）不会 bump 版本，靠 ``claim_revision`` 判不出来，
      旧结果会继续可读。所以这里必须连带作废结果。

    例外：``close`` 是**正常收尾**，终端结果卡要在这之后才被渲染，
    必须保留结果（及其归属 claim）直到其结果自己的 TTL。这类调用传
    ``keep_results=True``，把"结束会话"与"作废授权态"两件事分开。

    这不是"清理内存"，是**授权失效链**：任何让授权不再成立的时刻，都必须让
    基于该授权产生的所有中间态一并失效。
    """
    registry.drop_task(task_id)
    coordinator.release(task_id)
    if not keep_results:
        local_results.release(task_id)


def _bridge_origin(host_id: str) -> str:
    """稳定来源键：registry/coordinator 的 ``origin`` 只作命名空间判别。

    与 :mod:`python_data_api` 里用 Origin 头作 ``origin`` 不同，桥的绑定已唯一
    确定宿主，这里按 ``host_id`` 派一个稳定键，保证 ``coordinator.bind`` 与
    ``registry.create`` 用的是**同一个** origin（``execute_python_analysis`` 用
    ``binding.origin`` 反查，三者必须一致）。绝不使用模型/客户端可伪造的字段。
    """
    if host_id == "wps-et":
        return "bridge:wps-et"
    return "bridge:excel-office"


def _error_receipt(code: str, *, extra_code: str | None = None) -> dict[str, Any]:
    return {"status": "error", "errorCode": extra_code or code}


task_tool_bridge = TaskToolBridge()


def register_bridge_result(
    task_id: str, result: DataToolResult, source_fingerprint: str | None
) -> str:
    """Register a result on behalf of a task that already holds an authorization.

    Refuses when the task has no live ownership claim: registering values for an
    unknown task would let any local caller mint results for someone else.
    """
    if local_results.claim(task_id) is None:
        raise BridgeError("unowned_task", "任务没有生效的授权，不能登记结果")
    return local_results.register(task_id, result, source_fingerprint).handle

"""可信协调器：把 ``run_python_analysis`` 从"产品未接通"接到真实执行。

设计依据：``docs/SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md`` §2/§3/§5。

数据流（两端共用，值只在可信本地内存，绝不进模型/账本/持久化）：

1. **数据登记**（宿主读表 → 值回执）：宿主（Office.js / WPS）在用户已授权、
   已确认的范围内读表，把值矩阵经本机工具桥回执交给服务端，登记进
   :mod:`python_datasets.PythonDatasetRegistry`。这一步建立
   ``taskId → PythonTaskBinding`` 的映射，绑定任务、宿主、工作簿身份、
   授权版本与源指纹。模型只拿到 ``dataset://`` 句柄字符串，**没有** fingerprint。

2. **隔离执行**（``execute_python_analysis``）：模型提供 ``code`` +
   ``datasetHandles``。协调器**按 taskId 查 binding**，从可信登记里按句柄
   反查 fingerprint（不信任模型给的 fingerprint），解析数据 → 隔离执行 →
   结果登记为 :mod:`local_results` 本地结果 → 返回 ``resultRef``。

3. **结果转计划**（复用既有链路）：模型用 ``resultRef`` 走
   ``submit_result_plan`` → :func:`result_plan.compile_result_plan` →
   白名单 ``writeValues`` 计划 → 双端预览 → 用户确认 → 写回 → 回读。

安全边界（不可越过）：
- 模型自报的 ``dataset://`` 句柄只是字符串，**不是授权凭据**；fingerprint
  一律从可信登记反查，句柄不在登记里或指纹对不上就失败关闭。
- 授权变化 / 换簿 / 源变化 / 取消 / 过期 / 迟到结果都使 binding 或结果失效。
- 原始 Python 代码只在任务受控内存短期存在，不进入任何持久化现场。
- 结果值矩阵不进模型回执、远端校验或持久账本；只作为 ``resultRef`` 由
  本地结果卡 / 结果计划在可信侧取回。
"""

from __future__ import annotations

import asyncio
import threading
import time
from dataclasses import dataclass, field
from typing import Any

from .capabilities import capability_int
from .local_results import local_results
from .models import (
    DataToolResult,
    PythonDataScope,
    PythonDatasetReceipt,
)
from .python_analysis import (
    REASON_DISABLED,
    REASON_EXECUTION_ERROR,
    REASON_NONE,
    REASON_RUNNER_MISSING,
    RECEIPT_STATUS_COMPLETED,
    _rejected_receipt,
    _unavailable_receipt,
)
from .python_datasets import DatasetRejected, registry

# 动态调用 python_analysis 模块级函数，而非 import 绑定：这样测试里对
# ``python_analysis.python_execution_enabled`` / ``_runner_status`` /
# ``run_registered_analysis_internal`` 的 monkeypatch 能对本协调器生效
# （import 绑定会把函数引用钉死在导入时刻）。
from . import python_analysis as _pa


@dataclass(frozen=True)
class PythonTaskBinding:
    """服务端可信侧为一次任务建立的数据绑定。

    ``scope`` / ``origin`` / ``receipts`` 是数据登记时产生的服务端凭据，
    模型永远拿不到（它只有 ``dataset://`` 句柄字符串）。绑定按 ``task_id``
    索引，换授权 / 换簿 / 换任务都建立新绑定，旧绑定作废。
    """

    task_id: str
    host_id: str
    workbook_name: str
    scope: PythonDataScope
    scope_token: str
    origin: str
    receipts: dict[str, PythonDatasetReceipt]  # datasetHandle -> receipt
    created_at: float
    # 数据登记那一刻的源快照指纹。结果登记时要用它给结果打上"来自哪个源"，
    # 不能留 None —— 否则 ``invalidate_source`` 无法判断这份结果算不算过期。
    source_fingerprint: str | None = None
    # datasetHandle -> 该块对应的 **来源下标**。回执块可能乱序交回，所以不能
    # 用 receipts 的迭代顺序当来源顺序。
    source_indexes: dict[str, int] = field(default_factory=dict)

    def find_receipt(self, handle: str) -> PythonDatasetReceipt | None:
        """按句柄反查可信 receipt；只有句柄字符串绝不返回（fingerprint 由此获得）。"""
        return self.receipts.get(handle)

    def sheets_of(self, handles: list[str]) -> list[str] | None:
        """这批句柄**实际**对应的来源工作表名。

        不能拿 ``scope.sources`` 全体冒充：模型可以只把其中一部分句柄交给
        ``run_python_analysis``，那份结果也就只依赖那几张表。查不到下标时返回
        ``None``（无法确认），由调用方失败关闭 —— 不猜、不扩大。
        """
        names: list[str] = []
        for handle in handles:
            index = self.source_indexes.get(handle)
            if not isinstance(index, int) or not 0 <= index < len(self.scope.sources):
                return None
            names.append(self.scope.sources[index].sheetName)
        # 去重且保持声明顺序。
        seen: list[str] = []
        for name in names:
            if name not in seen:
                seen.append(name)
        return seen


class PythonTaskCoordinator:
    """任务级 Python 数据绑定注册表（易失、有 TTL、有容量上限）。"""

    def __init__(self, *, clock=time.monotonic) -> None:
        self._clock = clock
        self._bindings: dict[str, PythonTaskBinding] = {}
        self._max_bindings = capability_int("pythonData", "maxScopes")

    def _prune(self) -> None:
        ttl = capability_int("pythonData", "scopeTtlSeconds")
        cutoff = self._clock() - ttl
        for task_id in [t for t, b in self._bindings.items() if b.created_at < cutoff]:
            self.release(task_id)
        while len(self._bindings) > self._max_bindings:
            oldest = min(self._bindings, key=lambda t: self._bindings[t].created_at)
            self.release(oldest)

    def bind(
        self,
        *,
        task_id: str,
        host_id: str,
        workbook_name: str,
        scope: PythonDataScope,
        scope_token: str,
        origin: str,
        receipts: list[PythonDatasetReceipt],
        source_fingerprint: str | None = None,
        source_indexes: dict[str, int] | None = None,
    ) -> PythonTaskBinding:
        """登记一次任务的数据绑定；同任务重绑会作废旧绑定（换授权/换簿）。"""
        self._prune()
        binding = PythonTaskBinding(
            task_id=task_id,
            host_id=host_id,
            workbook_name=workbook_name,
            scope=scope,
            scope_token=scope_token,
            origin=origin,
            receipts={r.datasetHandle: r for r in receipts},
            created_at=self._clock(),
            source_fingerprint=source_fingerprint,
            source_indexes=dict(source_indexes or {}),
        )
        self._bindings[task_id] = binding
        return binding

    def lookup(self, task_id: str) -> PythonTaskBinding | None:
        self._prune()
        return self._bindings.get(task_id)

    def release(self, task_id: str) -> None:
        self._bindings.pop(task_id, None)

    def reset(self) -> None:
        self._bindings.clear()

    def count(self) -> int:
        return len(self._bindings)


coordinator = PythonTaskCoordinator()


def _matrix_result_to_data_tool_result(
    result_value: Any, *, source_sheets: list[str]
) -> DataToolResult | None:
    """把隔离执行得到的 ``{kind: "matrix", values: [...]}`` 转成本地结果。

    只有严格的矩阵形态可转；其它形态（标量、非 matrix）返回 None，调用方
    失败关闭，不伪造结果卡。
    """
    if not isinstance(result_value, dict):
        return None
    if result_value.get("kind") != "matrix":
        return None
    values = result_value.get("values")
    if not isinstance(values, list) or not values:
        return None
    if not all(isinstance(row, list) and row for row in values):
        return None
    columns = len(values[0])
    if any(len(row) != columns for row in values):
        return None
    headers = [f"列{i + 1}" for i in range(columns)]
    return DataToolResult(
        requestId=f"python-{len(values)}",
        title="Python 分析结果",
        headers=headers,
        rows=[list(row) for row in values],
        sourceSheets=source_sheets,
        scannedRows=len(values),
        complete=True,
        calculation="隔离 Python 分析得到的固定值结果",
        warnings=[],
    )


def execute_python_analysis(request: Any, arguments: dict[str, Any],
                            *, cancel_check=None) -> dict[str, Any]:
    """``run_python_analysis`` 工具入口的可信正向实现（失败关闭）。

    与 :mod:`python_analysis.execute_python_analysis` 的区别：本函数在开关与
    运行器都就绪时**真的执行**（解析已登记数据 → 隔离执行 → 结果登记为本地
    结果），而不是返回 ``product_not_ready``。仍复用 :mod:`python_analysis`
    的入参校验、固定文案与投影；本函数只负责"接上真实执行"这一段。

    返回有类型回执（供 ``project_python_analysis_receipt`` 投影）：
    - 成功：``{status: "completed", resultHandle, message}``
    - 失败：``{status: unavailable|rejected, reasonCategory, message}``
    """
    if not _pa.python_execution_enabled():
        return _unavailable_receipt(REASON_DISABLED)
    code = arguments.get("code") if isinstance(arguments, dict) else None
    from .python_analysis import (
        PYTHON_CODE_WITHHELD_MARKER,
        REASON_CODE_NOT_PRESERVED,
        _validate_arguments,
    )
    if code == PYTHON_CODE_WITHHELD_MARKER:
        return _unavailable_receipt(REASON_CODE_NOT_PRESERVED)
    if not _validate_arguments(arguments):
        return _rejected_receipt()
    if _pa._runner_status() is None:
        return _unavailable_receipt(REASON_RUNNER_MISSING)

    # 任务绑定：必须存在（数据登记已通过工具桥完成），否则句柄无法验证。
    task_id = getattr(request, "turnId", None) if request is not None else None
    if not isinstance(task_id, str) or not task_id:
        return _rejected_receipt()
    binding = coordinator.lookup(task_id)
    if binding is None:
        # 没有可信数据登记 → 模型自报句柄无法验证，失败关闭。
        return _rejected_receipt()

    handles = arguments.get("datasetHandles") or []
    receipts: list[PythonDatasetReceipt] = []
    for handle in handles:
        receipt = binding.find_receipt(handle)
        if receipt is None:
            # 句柄不在可信登记里：不信任模型自报的句柄，失败关闭。
            return _rejected_receipt()
        receipts.append(receipt)

    if not receipts:
        return _rejected_receipt()

    result = _pa.run_registered_analysis_internal(
        data_registry=registry,
        scope_token=binding.scope_token,
        scope=binding.scope,
        origin=binding.origin,
        receipts=receipts,
        code=code,
        cancel_check=cancel_check,
    )

    status = result.get("status")
    if status != "completed":
        reason = result.get("reasonCategory")
        if status == "rejected":
            # 入参/授权变化/结果非法：统一 rejected，不泄露原因细节。
            return _rejected_receipt()
        if status == "cancelled":
            # 取消/授权变化竞态：明确不可用，不发布迟到结果。
            return _unavailable_receipt(REASON_EXECUTION_ERROR)
        # unavailable / error：映射 runner 固定失败类别到统一公开文案。
        return _unavailable_receipt(_normalize_reason(reason))

    result_value = result.get("result")
    # **实际**来源依赖：只算模型真正交进来并被解析的那一批句柄对应的表，
    # 不是「本次登记的全部来源」。登记了两张表、只用了其中一张，结果就只依赖
    # 那一张 —— 写成全体是虚报来源。
    source_sheets = binding.sheets_of(handles)
    if source_sheets is None:
        # 查不到句柄对应哪张来源 → 无法确认来源依赖，失败关闭，不猜。
        return _unavailable_receipt(REASON_EXECUTION_ERROR)
    data_result = _matrix_result_to_data_tool_result(result_value, source_sheets=source_sheets)
    if data_result is None:
        # 非矩阵结果不能转本地结果：失败关闭，不伪造结果卡。
        return _unavailable_receipt(REASON_EXECUTION_ERROR)

    try:
        registered = local_results.register(
            task_id, data_result, binding.source_fingerprint
        )
    except ValueError:
        return _unavailable_receipt(REASON_EXECUTION_ERROR)

    # 构建结果卡（用于 SSE 事件投递）
    headers = [
        header if isinstance(header, str) else str(header) for header in registered.headers
    ]
    result_card = {
        "resultRef": registered.handle,
        "title": "Python 分析结果",
        "headers": [header[:200] for header in headers][:50] or ["（空）"],
        "rowCount": registered.row_count,
        "columnCount": max(1, registered.column_count),
        "sourceSheets": [str(sheet)[:100] for sheet in registered.source_sheets][:100],
        "complete": bool(registered.complete),
    }

    return {
        "status": RECEIPT_STATUS_COMPLETED,
        "reasonCategory": REASON_NONE,
        "resultHandle": registered.handle,
        "resultCard": result_card,
        "message": "隔离 Python 分析完成，结果已登记为本地结果。",
    }


async def execute_python_analysis_async(
    request: Any, arguments: dict[str, Any], *, cancel_check=None, on_event=None
) -> dict[str, Any]:
    """把**阻塞的隔离执行**挪出事件循环，并让取消真实传导到容器。

    为什么必须挪走：``run_python_isolated`` 是一次同步容器运行（上限
    ``maxWallTimeSeconds``）。FastAPI 端点与 SDK 的 MCP 工具处理器都跑在**同一
    个**事件循环上，同步调用会把整台服务一起冻住 —— 其它轮次的流式事件、
    工具桥 ``/poll``、取消、健康检查都停摆。

    为什么不能"丢进线程就完事"：线程里的容器脱离了 asyncio 的取消语义 —— 外层
    task 被 cancel 时线程还在跑，容器没人终止（这正好违背"取消必须让真实执行
    终止"）。所以这里把取消信号接到执行器的 ``cancel_check``（执行器据此杀掉
    容器），并且**等到真实执行终止**才把 ``CancelledError`` 抛出去：绝不留下
    一个脱离监管的容器。
    """
    cancelled = threading.Event()

    def guard() -> bool:
        return cancelled.is_set() or bool(cancel_check and cancel_check())

    loop = asyncio.get_running_loop()
    future = loop.run_in_executor(
        None,
        lambda: execute_python_analysis(request, arguments, cancel_check=guard),
    )
    try:
        result = await asyncio.shield(future)
        # 如果执行成功且有 resultHandle，发送一个进度事件通知前端。
        if on_event and result.get("status") == RECEIPT_STATUS_COMPLETED:
            result_card = result.get("resultCard")
            if result_card:
                await on_event({
                    "phase": "planning",
                    "stepId": "python.completed",
                    "status": "completed",
                    "title": "隔离 Python 分析完成",
                    "summary": "结果已登记为本地结果",
                    "pythonResult": result_card,
                })
        return result
    except asyncio.CancelledError:
        # 先让执行器自己停下来，再向上抛取消。
        cancelled.set()
        while not future.done():
            try:
                await asyncio.shield(future)
            except asyncio.CancelledError:
                if future.done():
                    break
                await asyncio.sleep(0)
        raise


def _normalize_reason(reason: str | None) -> str:
    """把 runner 的固定失败类别映射回统一公开文案可识别的类别。"""
    from .python_analysis import (
        REASON_CLEANUP_UNVERIFIED,
        REASON_IMAGE_NOT_IMMUTABLE,
    )
    if reason in (REASON_CLEANUP_UNVERIFIED, REASON_IMAGE_NOT_IMMUTABLE):
        return REASON_EXECUTION_ERROR
    if reason in ("runner_missing", "image_missing", "endpoint_not_local"):
        return REASON_RUNNER_MISSING
    return REASON_EXECUTION_ERROR

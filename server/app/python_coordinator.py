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
import uuid
from dataclasses import dataclass, field
from typing import Any

from .capabilities import capability_int
from . import dev_result_ref_fault
from .local_results import local_results
from .python_code_artifacts import python_code_artifacts
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
    validate_python_arguments,
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


@dataclass(frozen=True)
class PreparedPythonExecution:
    """一次**已通过全部前置闸门、可以真的起执行器**的分析。

    `prepare` 与 `run` 拆开是为了"正在运行"这句话有证据可依：只有拿到这个对象
    才允许宣告 running，而 running 的发出点又下沉到执行器真正启动的那一刻
    （`python_analysis.run_isolated_analysis_internal` 的 `on_runner_start`）。
    """

    task_id: str
    code: str
    handles: list[str]
    receipts: list[PythonDatasetReceipt]
    binding: PythonTaskBinding


# 未开始执行时的两种**不同**结论：入参不合法是"校验没过"，
# 其余（功能关闭、运行器缺失、没有可信绑定、授权拒绝）是"根本没开始跑"。
STAGE_VALIDATION_FAILED = "validation_failed"
STAGE_NOT_STARTED = "not_started"


def prepare_python_execution(
    request: Any, arguments: dict[str, Any]
) -> tuple[PreparedPythonExecution | None, dict[str, Any], str | None]:
    """执行前的全部确定性检查。返回 ``(prepared, receipt, stage)``。

    ``prepared`` 为 ``None`` 时 ``receipt`` 是"为什么没开始"的固定回执，
    ``stage`` 说明该用哪个展示状态（校验未通过 / 未开始执行）；
    **此时执行器一次都没被调用过**，调用方不得发出 running / executedVersion。
    """
    if not _pa.python_execution_enabled():
        return None, _unavailable_receipt(REASON_DISABLED), STAGE_NOT_STARTED
    code = arguments.get("code") if isinstance(arguments, dict) else None
    from .python_analysis import (
        PYTHON_CODE_WITHHELD_MARKER,
        REASON_CODE_NOT_PRESERVED,
        _validate_arguments,
    )
    if code == PYTHON_CODE_WITHHELD_MARKER:
        return None, _unavailable_receipt(REASON_CODE_NOT_PRESERVED), STAGE_NOT_STARTED
    if not _validate_arguments(arguments):
        return None, _rejected_receipt(), STAGE_VALIDATION_FAILED
    if _pa._runner_status() is None:
        return None, _unavailable_receipt(REASON_RUNNER_MISSING), STAGE_NOT_STARTED

    # 任务绑定：必须存在（数据登记已通过工具桥完成），否则句柄无法验证。
    task_id = getattr(request, "turnId", None) if request is not None else None
    if not isinstance(task_id, str) or not task_id:
        return None, _rejected_receipt(), STAGE_NOT_STARTED
    binding = coordinator.lookup(task_id)
    if binding is None:
        # 没有可信数据登记 → 模型自报句柄无法验证，失败关闭（没开始执行）。
        return None, _rejected_receipt(), STAGE_NOT_STARTED

    handles = arguments.get("datasetHandles") or []
    receipts: list[PythonDatasetReceipt] = []
    for handle in handles:
        receipt = binding.find_receipt(handle)
        if receipt is None:
            # 句柄不在可信登记里：不信任模型自报的句柄，失败关闭（没开始执行）。
            return None, _rejected_receipt(), STAGE_NOT_STARTED
        receipts.append(receipt)

    if not receipts:
        return None, _rejected_receipt(), STAGE_NOT_STARTED

    return PreparedPythonExecution(
        task_id=task_id,
        code=code,
        handles=list(handles),
        receipts=receipts,
        binding=binding,
    ), {}, None


def run_prepared_execution(
    prepared: PreparedPythonExecution,
    *, cancel_check=None, on_runner_start=None,
) -> dict[str, Any]:
    """真的起隔离执行器，并把结果登记为本地结果。

    ``on_runner_start`` 由执行器收到工作程序执行确认后回调，不以创建进程代替。
    """
    binding = prepared.binding
    result = _pa.run_registered_analysis_internal(
        data_registry=registry,
        scope_token=binding.scope_token,
        scope=binding.scope,
        origin=binding.origin,
        receipts=prepared.receipts,
        code=prepared.code,
        cancel_check=cancel_check,
        on_runner_start=on_runner_start,
    )

    status = result.get("status")
    if status != "completed":
        reason = result.get("reasonCategory")
        if status == "rejected":
            # 入参/授权变化/结果非法：统一 rejected，不泄露原因细节。
            return _rejected_receipt()
        if status == "cancelled":
            reason = "authorization_changed" if reason == "authorization_changed" else "cancelled"
            return {"status": "cancelled", "reasonCategory": reason,
                    "message": _pa._fixed_receipt_message("cancelled", reason)}
        # unavailable / error：映射 runner 固定失败类别到统一公开文案。
        return _unavailable_receipt(_normalize_reason(reason))

    result_value = result.get("result")
    # **实际**来源依赖：只算模型真正交进来并被解析的那一批句柄对应的表，
    # 不是「本次登记的全部来源」。登记了两张表、只用了其中一张，结果就只依赖
    # 那一张 —— 写成全体是虚报来源。
    source_sheets = binding.sheets_of(prepared.handles)
    if source_sheets is None:
        # 查不到句柄对应哪张来源 → 无法确认来源依赖，失败关闭，不猜。
        return _unavailable_receipt(REASON_EXECUTION_ERROR)
    data_result = _matrix_result_to_data_tool_result(result_value, source_sheets=source_sheets)
    if data_result is None:
        # 非矩阵结果不能转本地结果：失败关闭，不伪造结果卡。
        return _unavailable_receipt(REASON_EXECUTION_ERROR)

    try:
        registered = local_results.register(
            binding.task_id, data_result, binding.source_fingerprint
        )
    except ValueError:
        return _unavailable_receipt(REASON_EXECUTION_ERROR)

    # 开发验收入口（A1-E1）：本条是**该任务的第一条 Python 结果**时，把用户的显式
    # 启用绑定到这个句柄。绑定只做登记，不改值、不改 TTL；失效发生在引用被使用那一刻。
    dev_result_ref_fault.bind_python_result(binding.task_id, registered.handle)

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


def execute_python_analysis(request: Any, arguments: dict[str, Any],
                            *, cancel_check=None, on_runner_start=None) -> dict[str, Any]:
    """``run_python_analysis`` 工具入口的可信正向实现（失败关闭）。

    = :func:`prepare_python_execution` + :func:`run_prepared_execution`。
    拆开之后，"没开始执行"与"执行了但失败"是两件可以分开说的事。
    """
    prepared, receipt, _stage = prepare_python_execution(request, arguments)
    if prepared is None:
        return receipt
    return run_prepared_execution(
        prepared, cancel_check=cancel_check, on_runner_start=on_runner_start
    )


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
    # 这一次隔离执行的**调用身份**：同一次分析重放仍是这一步，再跑一次是新的。
    # 原来写死 `python.completed`，一个任务里跑两次分析会合并成一个步骤 ——
    # 第二次的完成事件覆盖第一次，界面上看起来只分析了一次。
    run_id = f"py_{uuid.uuid4().hex[:12]}"
    # 代码工件的身份与执行身份**分开**：`run_id` 是这次调用的身份，
    # `code_id` 是这段代码的身份。它们同源但语义不同 —— 界面据此判定
    # "已运行的就是屏幕上这一版"，而不是靠文案。
    code_id = f"code_{run_id}"
    # 这段代码的生命周期只占**一个** stepId：同一 stepId 的事件在原地更新，
    # 于是代码块不会因为"校验了""跑了"而变成三块。
    step_id = f"python.analysis:{run_id}"
    code_text = arguments.get("code") if isinstance(arguments, dict) else None
    # 取回正文时必须同时对上任务身份（见 python_code_artifacts 的说明）。
    code_task_id = getattr(request, "turnId", None) if request is not None else None
    if not isinstance(code_task_id, str):
        code_task_id = ""

    async def emit(event: dict[str, Any]) -> None:
        if on_event is not None:
            await on_event(event)

    def code_stage(
        stage: str, *, version: int, **extra: Any
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "codeId": code_id,
            "attemptId": run_id,
            "stage": stage,
            "version": version,
            "complete": True,
        }
        payload.update(extra)
        return payload
    def guard() -> bool:
        return cancelled.is_set() or bool(cancel_check and cancel_check())

    show_code = bool(
        on_event is not None and isinstance(code_text, str) and code_text.strip()
    )

    if show_code:
        # 代码在这里是**完整**的（SDK 的 tool use 一次送达整段代码）。因此第一条
        # 状态是"代码已生成"，绝不说"正在生成" —— 那会把"已经到齐"演成"逐字在写"。
        #
        # **正文不进事件**：事件会被写进 planning_task_events（SQLite），而仓库
        # 的硬边界是"生成的代码不进入持久化现场"。正文登记在进程内登记表里，
        # 事件只带工件身份与规模，前端按 `codeId` 经本机回环端点取回。
        python_code_artifacts.register(
            task_id=code_task_id,
            code_id=code_id,
            attempt_id=run_id,
            code=code_text,
        )
        await emit({
            "phase": "planning",
            "stepId": step_id,
            "status": "started",
            "title": "代码已生成",
            "summary": "代码已完整到达，接下来做本地校验",
            "pythonCode": code_stage("generated", version=1, chars=len(code_text)),
        })

    # ---- 执行前的全部确定性检查 -------------------------------------------
    # 这一步**不碰执行器**。它返回的结论决定后面到底有没有"正在运行"这回事：
    # 功能关闭 / 运行器缺失 / 没有可信绑定 / 入参非法，都在这里停下。
    prepared, refusal, refusal_stage = prepare_python_execution(request, arguments)
    if prepared is None:
        if show_code:
            # 没能开始执行：如实说"未开始执行"（或"校验未通过"），并且**不带**
            # executedVersion —— 否则界面会显示"已运行"，而实际容器一次都没起。
            reason_text = str(refusal.get("message") or "")
            await emit({
                "phase": "planning",
                "stepId": step_id,
                "status": "failed",
                "title": "代码没有开始执行"
                if refusal_stage == STAGE_NOT_STARTED
                else "代码校验未通过",
                "summary": reason_text
                or "这段代码没有进入隔离执行，表格没有被改动",
                "pythonCode": code_stage(
                    refusal_stage or STAGE_NOT_STARTED,
                    version=1,
                    chars=len(code_text or ""),
                    **({"reason": reason_text} if reason_text else {}),
                ),
            })
        return refusal

    if show_code:
        await emit({
            "phase": "planning",
            "stepId": step_id,
            "status": "progress",
            "title": "正在校验 Python 代码",
            "summary": "代码已生成，正在做本地校验",
            "pythonCode": code_stage("validating", version=1, chars=len(code_text or "")),
        })

    loop = asyncio.get_running_loop()
    # running 的证据只可能来自执行器自己：`on_runner_start` 在收到执行确认时由
    # 工作线程同步回调，这里把它转成一个事件并**登记下来**（先于终态等待），
    # 保证"正在运行"一定出现在终态之前，而不是靠调度顺序碰运气。
    started = {"yes": False}
    pending_emits: list[Any] = []

    async def emit_running() -> None:
        if not show_code:
            return
        await emit({
            "phase": "planning",
            "stepId": step_id,
            "status": "progress",
            "title": "正在隔离环境分析",
            "summary": "执行器已启动，正在隔离容器里运行",
            # 执行器确实启动了，才允许说"版本 1 正在运行"。
            "pythonCode": code_stage("running", version=1, executedVersion=1),
        })

    def on_runner_start() -> None:
        started["yes"] = True
        try:
            pending_emits.append(asyncio.run_coroutine_threadsafe(emit_running(), loop))
        except RuntimeError:
            # 事件循环已经关了（服务正在收尾）：不改事实，只是发不出进度。
            pass

    async def flush_start_events() -> None:
        for pending in pending_emits:
            try:
                await asyncio.wrap_future(pending)
            except Exception:
                pass
        pending_emits.clear()

    future = loop.run_in_executor(
        None,
        lambda: run_prepared_execution(
            prepared, cancel_check=guard, on_runner_start=on_runner_start
        ),
    )
    try:
        result = await asyncio.shield(future)
        await flush_start_events()
        status = result.get("status")
        if show_code:
            executed = {"executedVersion": 1} if started["yes"] else {}
            if status == RECEIPT_STATUS_COMPLETED:
                await emit({
                    "phase": "planning",
                    "stepId": step_id,
                    "status": "completed",
                    "title": "隔离 Python 分析完成",
                    "summary": "结果已登记为本地结果",
                    "pythonCode": code_stage("analysis_completed", version=1, **executed),
                    **({"pythonResult": result["resultCard"]} if result.get("resultCard") else {}),
                })
            elif status == "cancelled":
                await emit({
                    "phase": "planning",
                    "stepId": step_id,
                    "status": "cancelled",
                    "title": "隔离 Python 分析已暂停" if started["yes"] else "代码执行已取消",
                    "summary": "这次执行被中止，结果没有发布",
                    "pythonCode": code_stage("paused" if started["yes"] else STAGE_NOT_STARTED,
                                             version=1, reason=result.get("message"), **executed),
                })
            else:
                # 检查通过后仍可能在启动时失败。没收到执行确认时不能声称跑过代码。
                await emit({
                    "phase": "planning",
                    "stepId": step_id,
                    "status": "failed",
                    "title": "隔离 Python 分析未完成",
                    "summary": "这次执行没有产出结果，表格没有被改动",
                    "pythonCode": code_stage("execution_failed" if started["yes"] else STAGE_NOT_STARTED,
                                             version=1, reason=result.get("message"), **executed),
                })
        # 如果执行成功且有 resultHandle，发送一个进度事件通知前端。
        if on_event and not show_code and status == RECEIPT_STATUS_COMPLETED:
            result_card = result.get("resultCard")
            if result_card:
                await on_event({
                    "phase": "planning",
                    "stepId": f"python.completed:{run_id}",
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
            except Exception:
                # 取消后的执行器异常也必须收尾；不能跳过下面的取消终态。
                break
        await flush_start_events()
        if show_code:
            # 取消也可能发生在"执行器还没启动"的时候：那就说"未开始执行"，
            # 不能说"已暂停"（暂停意味着它跑过一段）。
            stage = "paused" if started["yes"] else STAGE_NOT_STARTED
            await emit({
                "phase": "planning",
                "stepId": step_id,
                "status": "cancelled",
                "title": "隔离 Python 分析已暂停" if started["yes"] else "代码执行已取消",
                "summary": "这次执行被中止，结果没有发布"
                if started["yes"]
                else "执行已取消，未收到代码开始执行的确认",
                "pythonCode": code_stage(
                    stage,
                    version=1,
                    **({"executedVersion": 1} if started["yes"] else {}),
                ),
            })
        raise
    except Exception:
        # 执行器抛异常也必须有终态：否则代码块永远停在"正在运行"。
        await flush_start_events()
        if show_code:
            await emit({
                "phase": "planning",
                "stepId": step_id,
                "status": "failed",
                "title": "隔离 Python 分析未完成",
                "summary": "这次执行没有产出结果，表格没有被改动",
                "pythonCode": code_stage(
                    "execution_failed" if started["yes"] else STAGE_NOT_STARTED,
                    version=1,
                    **({"executedVersion": 1} if started["yes"] else {}),
                ),
            })
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

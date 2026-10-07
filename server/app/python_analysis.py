"""隔离 Python 分析的可信协调器（失败关闭）。

设计依据：``docs/SDK_ISOLATED_PYTHON_DESIGN_2026-10-06.md``。

本模块的内部验收入口已接入真实 Docker 执行器和易失授权数据登记。
内部易失结果登记和本地预览编译已有实现；产品工具/UI 尚未接线，
仍只返回固定「不可用」回执。失败关闭要求：
运行器缺失、配置错误或健康检查不通过时，工具不得在任何宿主进程内继续，
也不得退回同用户子进程、venv 或临时目录来制造"已经实现"。

边界（不可越过）：
- 模型生成的代码只能进入**经过验证的隔离执行器**；本模块不运行代码。
- 本模块不读写工作簿、不访问宿主文件系统、不发起网络请求。
- 回执是**有类型的**：字段集合、取值枚举、状态组合、公开文案全部由本模块
  生成与校验，不信任调用方传入的自由文本；非法回执失败关闭。
- 原始 Python 代码只在任务受控内存中短期存在；本模块提供
  :func:`withheld_python_arguments` 供消息历史/检查点做确定性字段投影，
  代码字面量不进入任何持久化现场或模型历史。

启用前置（缺一不可，见 ``execution_status``）：
1. ``config/capabilities.json`` 的 ``pythonExecution.enabled = true``；
2. 已配置并验证通过的隔离执行器；
3. 产品授权、结果库和预览链路全部接线（尚未完成，开启配置也不执行）。
"""

from __future__ import annotations

import json
from typing import Any

from .capabilities import capability_bool, capability_int

# 固定公开文案。执行失败只允许安全错误类别，不允许原始 stderr/traceback 出站。
PYTHON_ANALYSIS_UNAVAILABLE_MESSAGE = "隔离 Python 不可用：尚未配置已验证的隔离执行器，本轮未运行任何代码。"
PYTHON_ANALYSIS_DISABLED_MESSAGE = "隔离 Python 未启用：功能保持关闭，本轮未运行任何代码。"
PYTHON_ANALYSIS_REJECTED_MESSAGE = "分析请求未通过校验：代码或数据句柄格式不正确，本轮未运行任何代码。"
PYTHON_ANALYSIS_CODE_NOT_PRESERVED_MESSAGE = (
    "原始代码未随会话保存，无法继续上次的 Python 调用；请重新提供完整分析代码。"
)
# 执行入口异常（参数解析/校验/协调器/分发/投影/序列化任一环节抛错）时的
# 统一安全文案：不携带 str(error)、traceback 或任何原始内容。
PYTHON_ANALYSIS_EXECUTION_ERROR_MESSAGE = (
    "隔离 Python 分析在执行环节发生内部错误；本轮未运行任何代码，请改用只读工具或直接提交计划。"
)

# 工具回执允许出站给模型的字段白名单。投影在可信侧完成，代码不可修改。
RECEIPT_KEYS = ("status", "runId", "reasonCategory", "message", "resultHandle")

# 回执是有类型的：状态与原因都是封闭枚举，组合白名单唯一决定公开文案。
RECEIPT_STATUS_UNAVAILABLE = "unavailable"
RECEIPT_STATUS_REJECTED = "rejected"
RECEIPT_STATUSES = (RECEIPT_STATUS_UNAVAILABLE, RECEIPT_STATUS_REJECTED)

REASON_DISABLED = "disabled"
REASON_RUNNER_MISSING = "runner_missing"
REASON_PRODUCT_NOT_READY = "product_not_ready"
REASON_INVALID_REQUEST = "invalid_request"
REASON_CODE_NOT_PRESERVED = "code_not_preserved"
REASON_EXECUTION_ERROR = "execution_error"
# 第八批新增：端点非本机 / 镜像非不可变 / 清理无法核实——都失败关闭、不外泄细节。
REASON_ENDPOINT_NOT_LOCAL = "endpoint_not_local"
REASON_IMAGE_NOT_IMMUTABLE = "image_not_immutable"
REASON_CLEANUP_UNVERIFIED = "cleanup_unverified"

# （status, reasonCategory）→ 固定公开文案。没有登记的结合即非法回执。
_ALLOWED_RECEIPT_COMBINATIONS: dict[tuple[str, str], str] = {
    (RECEIPT_STATUS_UNAVAILABLE, REASON_DISABLED): PYTHON_ANALYSIS_DISABLED_MESSAGE,
    (RECEIPT_STATUS_UNAVAILABLE, REASON_RUNNER_MISSING): PYTHON_ANALYSIS_UNAVAILABLE_MESSAGE,
    (RECEIPT_STATUS_UNAVAILABLE, REASON_PRODUCT_NOT_READY): "隔离 Python 运行器已就绪，但产品授权与结果预览尚未接通，本轮未运行代码。",
    (RECEIPT_STATUS_UNAVAILABLE, REASON_CODE_NOT_PRESERVED): PYTHON_ANALYSIS_CODE_NOT_PRESERVED_MESSAGE,
    (RECEIPT_STATUS_UNAVAILABLE, REASON_EXECUTION_ERROR): PYTHON_ANALYSIS_EXECUTION_ERROR_MESSAGE,
    (RECEIPT_STATUS_UNAVAILABLE, REASON_ENDPOINT_NOT_LOCAL): PYTHON_ANALYSIS_UNAVAILABLE_MESSAGE,
    (RECEIPT_STATUS_UNAVAILABLE, REASON_IMAGE_NOT_IMMUTABLE): PYTHON_ANALYSIS_UNAVAILABLE_MESSAGE,
    (RECEIPT_STATUS_UNAVAILABLE, REASON_CLEANUP_UNVERIFIED): PYTHON_ANALYSIS_EXECUTION_ERROR_MESSAGE,
    (RECEIPT_STATUS_REJECTED, REASON_INVALID_REQUEST): PYTHON_ANALYSIS_REJECTED_MESSAGE,
}

# 安全历史投影占位：代码字面量在消息历史/检查点中的唯一合法形态。
# 恢复轮执行到该占位时，协调器返回 code_not_preserved（明确"重新提供"，
# 不冒充断点续跑），绝不把它当作代码交给执行器。
PYTHON_CODE_WITHHELD_MARKER = "[原始代码未随会话保存：如需重试请重新提交完整代码]"

# 旧检查点里的 run_python_analysis 工具回执（修复前版本可能携带 error/message/
# stdout/traceback 自由文本）恢复时统一重建为这个固定回执，不读取旧内容。
PYTHON_ANALYSIS_OLD_RECEIPT_RECONSTRUCTED_CONTENT = json.dumps(
    {
        "status": RECEIPT_STATUS_UNAVAILABLE,
        "reasonCategory": REASON_CODE_NOT_PRESERVED,
        "message": PYTHON_ANALYSIS_CODE_NOT_PRESERVED_MESSAGE,
    },
    ensure_ascii=False,
)

# 旧检查点里存在无法可靠确认归属的自由文本（含疑似 Python 敏感输出）时，
# 恢复入口失败关闭的固定提示——明确要求重新生成预览，绝不猜测安全后重放。
PYTHON_ANALYSIS_UNATTRIBUTABLE_RESUME_MESSAGE = (
    "这段历史包含无法确认归属的 Python 分析上下文，出于隐私无法安全续跑；"
    "请重新生成预览，本轮不会重放旧内容。"
)

# 入参只允许这两个键；多余键（路径、命令行、密钥、宿主对象……）在可信侧拒绝。
_ARGUMENT_KEYS = ("code", "datasetHandles")


def python_execution_enabled() -> bool:
    """总开关。默认 False：未显式开启时工具不出现、调用即失败关闭。"""
    return capability_bool("pythonExecution", "enabled", default=False)


def python_execution_limits() -> dict[str, int]:
    """资源限额集中在 capabilities 配置（架构原则：限额不进代码）。"""
    section = "pythonExecution"
    return {
        "maxCodeChars": capability_int(section, "maxCodeChars"),
        "maxCallsPerTask": capability_int(section, "maxCallsPerTask"),
        "maxWallTimeSeconds": capability_int(section, "maxWallTimeSeconds"),
        "maxCpuTimeSeconds": capability_int(section, "maxCpuTimeSeconds"),
        "maxMemoryMb": capability_int(section, "maxMemoryMb"),
        "maxProcesses": capability_int(section, "maxProcesses"),
        "maxScratchMb": capability_int(section, "maxScratchMb"),
        "maxResultBytes": capability_int(section, "maxResultBytes"),
        "maxStdoutBytes": capability_int(section, "maxStdoutBytes"),
        "maxStderrBytes": capability_int(section, "maxStderrBytes"),
        "maxResultDepth": capability_int(section, "maxResultDepth"),
        "maxResultElements": capability_int(section, "maxResultElements"),
        "maxDatasetHandles": capability_int(section, "maxDatasetHandles"),
        "maxDatasetHandleChars": capability_int(section, "maxDatasetHandleChars"),
    }


def _fixed_receipt_message(status: str, reason_category: str) -> str:
    """公开文案只由可信状态组合决定，绝不使用调用方传入的自由文本。"""
    message = _ALLOWED_RECEIPT_COMBINATIONS.get((status, reason_category))
    if message is None:
        raise ValueError("回执状态组合非法")
    return message


def _unavailable_receipt(reason_category: str) -> dict[str, Any]:
    return {
        "status": RECEIPT_STATUS_UNAVAILABLE,
        "reasonCategory": reason_category,
        "message": _fixed_receipt_message(RECEIPT_STATUS_UNAVAILABLE, reason_category),
    }


def _runner_status() -> str | None:
    """已验证隔离执行器的健康检查。

    本批接入 Docker 运行时适配器（第一种实现，非唯一前提）。健康检查做
    **实测**（Docker 引擎可用 + 本机 Linux 容器 + 镜像就绪），不能只看配置
    文字。返回非空字符串表示"本机隔离运行时就绪"，否则 ``None`` 失败关闭。

    注意：此检查只判断运行时是否就绪，**不代表授权执行**；产品入口
    :func:`execute_python_analysis` 仍受 ``python_execution_enabled`` 总开关
    约束（保持 false），只有内部测试入口 :func:`run_isolated_analysis_internal`
    会真实执行。
    """
    try:
        from .python_docker_runner import (
            verify_engine, image_present, verified_endpoint, _docker_config,
        )
        config = _docker_config()
        endpoint = verified_endpoint()
        if not verify_engine(endpoint):
            return None
        if not image_present(config, endpoint):
            return None
        return config.image_ref()
    except Exception:
        # 运行时适配器导入/检查失败一律视为不可用（失败关闭），不回退裸执行。
        return None


def run_isolated_analysis_internal(
    code: str,
    input_payload: dict[str, Any],
    *, cancel_check=None,
) -> dict[str, Any]:
    """内部执行入口：真实调用 Docker 隔离运行时（仅供测试程序使用）。

    此入口**不经过**产品工具链（``_execute_tool``），不受
    ``pythonExecution.enabled`` 约束——它是"真实内部执行器"的测试入口，
    用于验证隔离/限额/清理证据。产品工具仍然保持 enabled=false，绝不开放。

    返回有类型结果：
    - 成功：``{"status": "completed", "result": <有类型值>}``
    - 失败：``{"status": <固定状态>, "reasonCategory": <固定类别>}``（不回显异常/代码/输出）
    """
    try:
        from .python_docker_runner import (
            DockerRuntimeError,
            run_python_isolated,
            verify_engine,
            image_present,
            validate_isolation_config,
            verified_endpoint,
            _docker_config,
        )
    except Exception:
        return {"status": "unavailable", "reasonCategory": "runner_missing"}

    limits = python_execution_limits()
    try:
        config = _docker_config()
    except DockerRuntimeError as exc:
        return {"status": "unavailable", "reasonCategory": exc.reason_category}

    try:
        # 隔离配置逐项校验：缺任一项即拒绝（不能只信默认值）
        isolation_ok, isolation_reason = validate_isolation_config(config)
        if not isolation_ok:
            # 区分「镜像缺不可变身份」与「其它隔离参数未落实」——两者都失败关闭，
            # 但原因不同，便于运行环境定位（都不回显原始细节）。
            if not config.is_immutable():
                return {"status": "unavailable",
                        "reasonCategory": REASON_IMAGE_NOT_IMMUTABLE}
            return {"status": "unavailable",
                    "reasonCategory": "isolation_not_applied"}
        # 锁定本机端点（拒绝远端/未知端点）
        endpoint = verified_endpoint()
        if not verify_engine(endpoint):
            return {"status": "unavailable", "reasonCategory": "runner_missing"}
        if not image_present(config, endpoint):
            return {"status": "unavailable", "reasonCategory": "image_missing"}
        result = run_python_isolated(
            config,
            code=code,
            input_payload=input_payload,
            cancel_check=cancel_check,
            cpu_quota=1.0,  # CPU 速率上限；累计 CPU 时间另由 max_cpu_seconds 表达
            cpu_seconds=limits["maxCpuTimeSeconds"],
            memory_mb=limits["maxMemoryMb"],
            # pids-limit：进程/线程数上限。模型代码在工作程序内执行（同进程），
            # 不是另一个必需进程——因此直接接线 maxProcesses，不再 +1。
            pids_limit=limits["maxProcesses"],
            scratch_mb=limits["maxScratchMb"],
            max_output_bytes=limits["maxResultBytes"],
            max_stderr_bytes=limits["maxStderrBytes"],
            max_result_depth=limits["maxResultDepth"],
            max_result_elements=limits["maxResultElements"],
            wall_timeout_seconds=limits["maxWallTimeSeconds"],
            endpoint=endpoint,
        )
    except DockerRuntimeError as exc:
        return {"status": "unavailable", "reasonCategory": exc.reason_category}
    except Exception:
        return {"status": "unavailable", "reasonCategory": "execution_error"}

    if result.is_success():
        return {"status": "completed", "result": result.result_payload.get("result")}
    return {"status": result.status, "reasonCategory": result.reason_category}


def run_registered_analysis_internal(*, data_registry, scope_token, scope, origin,
                                     receipts, code, cancel_check=None):
    """Synthetic/integration entry only; no HTTP/model product execution route.

    Resolves server-issued handles, passes values only to the locked container,
    cancels on grant invalidation and discards results if a grant changed while
    running. The host still must revalidate its live source before invocation.
    """
    from .python_datasets import DatasetRejected
    if not isinstance(code, str) or not code.strip() or len(code) > python_execution_limits()["maxCodeChars"]:
        return {"status": "rejected", "reasonCategory": "invalid_request"}
    try:
        values = data_registry.resolve(scope_token, scope, origin, receipts)
    except DatasetRejected:
        return {"status": "rejected", "reasonCategory": "authorization_changed"}

    def cancelled():
        return bool(cancel_check and cancel_check()) or not data_registry.valid(scope_token, scope, origin, receipts)

    result = run_isolated_analysis_internal(code, values, cancel_check=cancelled)
    # Even a runner success cannot publish a stale/late result.
    if cancelled():
        return {"status": "cancelled", "reasonCategory": "authorization_changed"}
    return result


def run_registered_result_internal(*, data_registry, scope_token, scope, origin,
                                   receipts, code, cancel_check=None):
    """Internal synthetic chain: validated matrix -> volatile local receipt.

    No model/HTTP execution route. The product tool remains fail-closed. Raw
    result bytes never become an assistant message, event or checkpoint here.
    """
    from .python_datasets import DatasetRejected
    response = run_registered_analysis_internal(data_registry=data_registry,
        scope_token=scope_token, scope=scope, origin=origin, receipts=receipts,
        code=code, cancel_check=cancel_check)
    if response.get("status") != "completed":
        return response
    try:
        receipt = data_registry.publish_result(scope_token, scope, origin, receipts, response.get("result"))
    except DatasetRejected:
        return {"status": "rejected", "reasonCategory": "invalid_result"}
    # Cancellation can race validation/publication; remove the exact result,
    # never a broad set, and don't report a late successful receipt.
    if (cancel_check and cancel_check()) or not data_registry.valid(scope_token, scope, origin, receipts):
        try:
            data_registry.discard_result(scope_token, scope, origin, receipts, receipt)
        except DatasetRejected:
            pass
        return {"status": "cancelled", "reasonCategory": "authorization_changed"}
    return {"status": "completed", "receipt": receipt.model_dump()}


def _validate_arguments(arguments: dict[str, Any]) -> bool:
    """可信侧严格入参校验（不依赖模型遵守 schema 或框架兜底）。返回 False 表示非法。

    拒绝类别：
    - 未知字段（path、apiKey、command 等一概拒绝，不做隐式转换）；
    - code 缺失、非字符串、空白、超过 ``maxCodeChars``；
    - datasetHandles 缺省视为空列表；非列表、数量超过 ``maxDatasetHandles``、
      单项非字符串、单项空白或超过 ``maxDatasetHandleChars``。

    形状合法不等于授权合法：datasetHandle 的登记核对尚未实现，不能把
    字符串校验冒充授权检查；无执行器时这里之后的路径仍然绝不运行代码。
    """
    if not isinstance(arguments, dict):
        return False
    if set(arguments.keys()) - set(_ARGUMENT_KEYS):
        return False
    limits = python_execution_limits()
    code = arguments.get("code")
    if not isinstance(code, str) or not code.strip():
        return False
    if len(code) > limits["maxCodeChars"]:
        return False
    handles = arguments.get("datasetHandles", [])
    if not isinstance(handles, list):
        return False
    if len(handles) > limits["maxDatasetHandles"]:
        return False
    for item in handles:
        if not isinstance(item, str) or not item.strip():
            return False
        if len(item) > limits["maxDatasetHandleChars"]:
            return False
    return True


def execute_python_analysis(request: Any, arguments: dict[str, Any]) -> dict[str, Any]:
    """``run_python_analysis`` 工具入口。**失败关闭**：没有执行器就没有执行。

    返回值是模型可见回执的**原始来源**，出站前还必须经过
    :func:`project_python_analysis_receipt` 的有类型校验与投影。
    """
    if not python_execution_enabled():
        return _unavailable_receipt(REASON_DISABLED)
    code = arguments.get("code") if isinstance(arguments, dict) else None
    if code == PYTHON_CODE_WITHHELD_MARKER:
        # 恢复轮：历史/检查点里的代码是投影占位，原始代码不可重建。
        # 明确要求模型重新提供完整代码，不冒充断点续跑。
        return _unavailable_receipt(REASON_CODE_NOT_PRESERVED)
    if not _validate_arguments(arguments):
        return _rejected_receipt()
    if _runner_status() is None:
        return _unavailable_receipt(REASON_RUNNER_MISSING)
    # Do not misreport an available runtime as missing or silently enable the
    # model tool before the product's grant/result/preview pipeline exists.
    return _unavailable_receipt(REASON_PRODUCT_NOT_READY)


def _rejected_receipt() -> dict[str, Any]:
    return {
        "status": RECEIPT_STATUS_REJECTED,
        "reasonCategory": REASON_INVALID_REQUEST,
        "message": _fixed_receipt_message(RECEIPT_STATUS_REJECTED, REASON_INVALID_REQUEST),
    }


def execution_error_receipt() -> dict[str, Any]:
    """执行路径异常的**统一安全回执**（缺陷 3）。

    参数解析、校验、协调器、工具分发、结果投影或序列化任一环节抛出异常时，
    模型可见内容只允许这个固定回执——不携带 ``str(error)``、traceback 或任何
    来自异常的文本。调用方（普通后端工具循环 / SDK handle）在捕获到
    ``run_python_analysis`` 相关异常时必须改用本回执，而不是把异常文本
    拼进工具消息、``last_error_message`` 或事件 trace。

    ``asyncio.CancelledError`` 继承自 ``BaseException``，不在常规
    ``except ValueError`` 捕获范围内，会自然向上传播——调用方**不得**用
    宽泛的 ``except Exception`` 吞掉取消信号。
    """
    return _unavailable_receipt(REASON_EXECUTION_ERROR)


def python_error_response_content() -> str:
    """执行路径异常的固定安全工具内容（供工具循环 / SDK 直接使用）。

    与 :func:`execution_error_receipt` 同一文案来源，但返回可直接作为
    ``tool`` 消息 content 的 JSON 串。不含任何来自异常的内容。
    """
    return json.dumps(execution_error_receipt(), ensure_ascii=False)


def project_python_analysis_receipt(receipt: dict[str, Any]) -> dict[str, Any]:
    """有类型回执校验与投影：模型可见内容的唯一出口。

    规则（全部失败关闭，非法即抛 ValueError，调用出口必须转固定安全文案）：
    - 回执必须是对象；字段集合 ⊆ :data:`RECEIPT_KEYS`，白名单外字段拒绝；
    - ``status``/``reasonCategory`` 必须是封闭枚举，且组合在登记表内；
    - 产品工具尚无可信结果发布通道（内部登记不等于模型授权）：
      ``runId``/``resultHandle`` 出现即拒绝；
      将来支持句柄时必须由可信协调器生成并核对任务归属，不能只验证
      字符串前缀或正则；
    - ``message`` 不读取输入——按可信状态组合重新生成固定公开文案。
    """
    if not isinstance(receipt, dict):
        raise ValueError("Python 分析回执必须是对象")
    if set(receipt.keys()) - set(RECEIPT_KEYS):
        raise ValueError("回执包含未登记字段")
    status = receipt.get("status")
    reason_category = receipt.get("reasonCategory")
    if not isinstance(status, str) or status not in RECEIPT_STATUSES:
        raise ValueError("回执状态非法")
    if not isinstance(reason_category, str):
        raise ValueError("回执原因类别非法")
    if (status, reason_category) not in _ALLOWED_RECEIPT_COMBINATIONS:
        raise ValueError("回执状态组合非法")
    if receipt.get("runId") is not None or receipt.get("resultHandle") is not None:
        raise ValueError("回执携带未登记的运行标识或结果句柄")
    return {
        "status": status,
        "reasonCategory": reason_category,
        "message": _fixed_receipt_message(status, reason_category),
    }


def withheld_python_arguments(arguments_json: str | None) -> str:
    """安全历史投影：**可信侧最小白名单**，与入参校验是两道独立边界。

    入参校验决定"这次调用要不要执行"；本函数决定"哪些内容允许进入消息历史、
    检查点与恢复现场"。被拒绝的参数**不得因为"没有执行"就原样保留**——未知字段、
    畸形对象、任意自由文本一律不进入历史。

    投影规则（确定性字段级处理，不做哨兵扫描、不做全局字符串替换）：
    - 无法解析为 JSON 对象 → 只保留 ``{"code": marker}``；
    - 解析为对象时，只保留两个可信键：
      ``code`` 固定替换为占位；
      ``datasetHandles`` 一律投影为**空列表**（固定、不含原值的兼容表示）。
      句柄登记核对尚未实现，任何模型自报的句柄字符串——哪怕形状合法、甚至
      长得像正常 ``dataset://`` 句柄——都无法证明是可信侧签发的；因此历史/
      检查点投影**不保留原值**，用空列表清楚表达「未验证」，绝不冒充已授权。
      形状校验仍是 ``_validate_arguments``（入参那道边界）的职责，两道互不替代。
    - 其余键（path / apiKey / command / 任意自由文本）全部丢弃。

    工具调用 ID、工具名由调用方在消息层保留，本函数只负责参数内容投影。
    幂等：已投影的现场再次投影结果不变。
    """
    try:
        parsed = json.loads(arguments_json or "{}")
    except (ValueError, TypeError):
        parsed = None
    if not isinstance(parsed, dict):
        # 畸形 JSON 或非对象：不保留任何原始内容，只留占位。
        return json.dumps({"code": PYTHON_CODE_WITHHELD_MARKER}, ensure_ascii=False)
    projected: dict[str, Any] = {}
    projected["datasetHandles"] = _project_safe_dataset_handles(
        parsed.get("datasetHandles")
    )
    projected["code"] = PYTHON_CODE_WITHHELD_MARKER
    return json.dumps(projected, ensure_ascii=False)


def _project_safe_dataset_handles(handles: Any) -> list[str]:
    """历史投影对 datasetHandles 的唯一合法形态：**空列表**。

    句柄登记核对尚未实现，因此模型自报的任何句柄字符串都无法证明是可信侧
    签发的。本函数不读取 ``handles`` 的内容（也不做正则、前缀、UUID、长度
    之类的形状猜测来冒充授权），统一返回空列表——固定、不含原值的兼容表示，
    清楚表达「未验证」。入参的形状校验在 ``_validate_arguments`` 独立完成，
    那道边界决定本轮要不要执行，这道边界决定哪些内容进入历史，互不替代。
    """
    return []


def python_analysis_tool_definition() -> dict[str, Any]:
    """``run_python_analysis`` 的工具 schema（与其他工具同形）。

    入参不含路径、命令行、包安装指令、宿主对象或密钥；数据只能通过
    任务内 ``datasetHandles`` 不透明句柄引用，句柄由可信协调器核对
    （登记核对尚未实现；形状校验不冒充授权检查）。
    """
    return {
        "type": "function",
        "function": {
            "name": "run_python_analysis",
            "description": (
                "在已验证的隔离执行器中运行 Python 分析代码。只能使用预装分析库"
                "与本任务授权数据的句柄；不能读写工作簿、宿主文件或网络。"
                "执行器未就绪时返回不可用回执，此时请改用只读工具或直接提交计划。"
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "code": {
                        "type": "string",
                        "description": "要运行的 Python 分析代码",
                    },
                    "datasetHandles": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "本任务已获准的数据句柄（不透明引用，不是路径）",
                    },
                },
                "required": ["code"],
                "additionalProperties": False,
            },
        },
    }

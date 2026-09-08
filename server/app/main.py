from __future__ import annotations

import asyncio
import hashlib
import json
import logging
import uuid
import time
from collections import deque
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, AsyncIterator

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Request
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel

from .capabilities import capabilities, capability_int
from .capabilities import model_timeout_seconds
from .intent import check_intent
from .folder_workbooks import (
    FolderCatalog,
    FolderExecuteRequest,
    FolderQueryRequest,
    FolderRefreshRequest,
    FolderScanRequest,
    FolderSnapshotRequest,
    create_folder_snapshot,
    cleanup_folder_backups,
    delete_folder_backup,
    execute_folder_plan,
    folder_session_root,
    inspect_managed_folder_recovery,
    export_folder_backup,
    list_folder_backups,
    open_folder_backup_directory,
    set_folder_backup_pinned,
    query_folder_data,
    refresh_folder,
    restore_managed_folder_backup,
    scan_folder_by_path,
    select_and_scan_folder,
    collect_legacy_folder_backups,
    undo_folder_execution,
)
from .models import (
    AnalysisPlan,
    AssistantResponse,
    DataToolResult,
    FolderBackupCleanupRequest,
    FolderBackupCleanupResponse,
    FolderBackupDeleteRequest,
    FolderBackupDeleteResponse,
    FolderBackupExportRequest,
    FolderBackupExportResponse,
    FolderBackupOpenRequest,
    FolderBackupOpenResponse,
    FolderBackupPinRequest,
    FolderBackupPinResponse,
    FolderBackupsResponse,
    FolderExecuteResponse,
    FolderLegacyCollectRequest,
    FolderLegacyCollectResponse,
    FolderManagedRecoveryRequest,
    FolderManagedRecoveryResponse,
    FolderUndoRequest,
    FolderUndoResponse,
    IntentCheckRequest,
    IntentCheckResponse,
    ManagedFolderRecoveryDirectoriesResponse,
    ManagedFolderRecoveryDirectory,
    ModelSettingsResponse,
    PlanRequest,
    PlanValidateRequest,
    TestModelConnectionResponse,
    SetFormulaModelRequest,
    TaskProgressEvent,
    TurnRequest,
    TurnResponse,
    UpdateModelSettingsRequest,
    UpsertModelConnectionRequest,
    WorkbookSnapshot,
)
from .llm import (
    delete_managed_connection,
    draft_model_connection,
    DuplicateModelConnectionError,
    LLMConnectError,
    LLMHTTPStatusError,
    LLMResponseError,
    LLMTimeoutError,
    LLMTransportError,
    model_catalog,
    model_settings_view,
    model_status,
    ModelConnectionNotFoundError,
    ModelConnectionStoreError,
    selected_model_config,
    set_formula_model_id,
    test_model_connection,
    upsert_managed_connection,
    update_model_api_key,
)
from .folder_mutex import FolderWriteBusyError
from .task_ledger import (
    ExpiredOperationError,
    FolderOperationLedger,
    FolderOperationRecord,
    ManagedDirectoryRecord,
)
from .planner import create_plan
from .plan_chat import chat_modify_plan
from .rule_generator import (
    GenerateFormulaRequest,
    GenerateFormulaResponse,
    generate_formula,
)
from .turn_state import TurnStateError, turn_registry

load_dotenv("server/.env")
logger = logging.getLogger(__name__)

@asynccontextmanager
async def _app_lifespan(_app: FastAPI):
    try:
        await run_in_threadpool(
            _folder_operation_ledger.compact_terminal_operations,
            capability_int("taskRecovery", "operationRetentionDays"),
        )
    except Exception:
        logger.exception("启动时无法压缩已过期的文件夹任务记录")
    try:
        await run_in_threadpool(_scan_registered_work_directories)
    except Exception:
        logger.exception("启动时无法扫描已登记工作目录的恢复元数据")
    yield


app = FastAPI(
    title="Excel Bro Local Service",
    version="0.1.0",
    description="Local planning and analysis service for the Excel Bro add-in.",
    lifespan=_app_lifespan,
)

ALLOWED_ORIGINS = [
    "https://localhost:3000",
    "https://127.0.0.1:3000",
    "http://localhost:3000",
    "http://127.0.0.1:3000",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=False,
    allow_methods=["DELETE", "GET", "POST", "PUT"],
    allow_headers=["Content-Type"],
)


@app.middleware("http")
async def guard_state_changing_origins(request: Request, call_next):
    """拦截跨站静默调用。

    本地 API 无鉴权，任何本机进程都能直接调用；CORS 只限制浏览器「读取」
    响应，并不阻止「发出」请求。恶意网页可用简单请求让服务改配置、删连接、
    弹文件夹对话框等。浏览器发起的跨源请求必定带 Origin，这里对状态变更
    方法校验 Origin 白名单；不带 Origin 的（curl、本机脚本、安装器健康检查）
    放行——它们不是浏览器跨站攻击面。
    """
    if request.method in {"POST", "PUT", "DELETE"}:
        origin = request.headers.get("origin")
        if origin is not None and origin not in ALLOWED_ORIGINS:
            return JSONResponse(
                status_code=403,
                content={
                    "detail": {
                        "code": "FORBIDDEN_ORIGIN",
                        "message": "请求来源不被允许",
                        "retryable": False,
                    }
                },
            )
    return await call_next(request)

@app.middleware("http")
async def allow_private_network_preflight(request: Request, call_next):
    """回应 Private Network Access 预检。

    任务窗格是 https://localhost:3000，本地服务是 http://127.0.0.1:8765，
    属于「安全上下文访问私有网络」。新版 Chromium/WebView2 会在预检里带
    Access-Control-Request-Private-Network，缺少对应的响应头就直接拦掉请求，
    并把失败报成「No 'Access-Control-Allow-Origin' header」——极具误导性，
    实际请求根本没发出去。

    PNA 的目的正是阻止公网页面访问本机服务，所以这里只对 Origin 白名单回应，
    不能无条件放开：不在白名单的来源拿不到这个头，保留原有防护。
    """
    response = await call_next(request)
    if (
        request.method == "OPTIONS"
        and request.headers.get("access-control-request-private-network") == "true"
        and request.headers.get("origin") in ALLOWED_ORIGINS
    ):
        response.headers["Access-Control-Allow-Private-Network"] = "true"
    return response


_diagnostic_events: deque[dict[str, object]] = deque(maxlen=500)


@app.middleware("http")
async def record_request_metrics(request, call_next):
    started = time.perf_counter()
    try:
        response = await call_next(request)
    except Exception as error:
        # 未捕获异常必须也进诊断日志。以前只在 call_next 正常返回后记录，500
        # 请求因此对 /api/diagnostics 完全隐形；而未捕获异常的响应又不带 CORS
        # 头，浏览器只会报「No Access-Control-Allow-Origin」+ Failed to fetch。
        # 两者叠加的结果是真正的错误在前后端都看不见，只能去翻 uvicorn 终端。
        _diagnostic_events.append(
            {
                "timestamp": datetime.now().astimezone().isoformat(),
                "method": request.method,
                "path": request.url.path,
                "status": 500,
                "durationMs": round((time.perf_counter() - started) * 1000, 2),
                "error": f"{type(error).__name__}: {error}"[:300],
            }
        )
        raise
    duration_ms = round((time.perf_counter() - started) * 1000, 2)
    response.headers["X-Excel-Bro-Duration-Ms"] = str(duration_ms)
    _diagnostic_events.append(
        {
            "timestamp": datetime.now().astimezone().isoformat(),
            "method": request.method,
            "path": request.url.path,
            "status": response.status_code,
            "durationMs": duration_ms,
        }
    )
    return response


@app.get("/api/diagnostics")
async def diagnostics() -> dict[str, object]:
    return {
        "generatedAt": datetime.now().astimezone().isoformat(),
        "events": list(_diagnostic_events),
    }


def service_error(
    status_code: int,
    code: str,
    message: str,
    *,
    retryable: bool,
) -> HTTPException:
    return HTTPException(
        status_code=status_code,
        detail={
            "code": code,
            "message": message,
            "retryable": retryable,
        },
    )


def _map_turn_error(error: Exception) -> tuple[int, str, str, bool]:
    """Map planning/model exceptions to (status, code, message, retryable).

    JSON 端点 raise service_error 用它，流式端点把它序列化成流内 error 事件，
    确保两条路径的错误语义完全一致。返回 None 表示不是可识别的轮次错误。
    """
    if isinstance(error, LLMTimeoutError):
        return (504, "MODEL_TIMEOUT", "模型响应超时，可以重试当前对话轮次。", True)
    if isinstance(error, LLMConnectError):
        return (
            502,
            "MODEL_UNREACHABLE",
            "无法连接模型服务，请检查网络和 AI_BASE_URL。",
            True,
        )
    if isinstance(error, LLMHTTPStatusError):
        status = error.status_code
        if status == 429:
            # 过载/限流是暂时的。后端已自动退避重试若干次仍未成功，这里给用户
            # 可理解的中文提示，不透传上游英文原始 body（如 engine_overloaded_error）。
            return (
                503,
                "MODEL_OVERLOADED",
                "模型服务当前繁忙（过载或限流），已自动重试多次仍未成功，请稍候再重试。",
                True,
            )
        if status >= 500:
            return (
                502,
                "MODEL_HTTP_ERROR",
                "模型服务暂时不可用（服务端错误），请稍候重试。",
                True,
            )
        # 401/402/403/404 等可行动的状态码由 LLMHTTPStatusError 自己给中文提示。
        return (502, "MODEL_HTTP_ERROR", str(error), False)
    if isinstance(error, (LLMTransportError, LLMResponseError)):
        return (502, "MODEL_CONNECTION_ERROR", str(error), True)
    if isinstance(error, TurnStateError):
        return (409, error.code, str(error), False)
    if isinstance(error, ValueError):
        return (422, "INVALID_REQUEST", str(error), False)
    raise error


@app.get("/health")
async def health() -> dict[str, str | bool | None]:
    try:
        return {
            "status": "ok",
            **model_status(),
        }
    except ModelConnectionStoreError as error:
        raise service_error(
            500,
            "MODEL_CONNECTION_STORE_INVALID",
            str(error),
            retryable=False,
        ) from error


@app.get("/api/models")
async def models() -> dict[str, object]:
    try:
        return model_catalog()
    except ModelConnectionStoreError as error:
        raise service_error(
            500,
            "MODEL_CONNECTION_STORE_INVALID",
            str(error),
            retryable=False,
        ) from error


@app.get("/api/settings/model", response_model=ModelSettingsResponse)
async def get_model_settings() -> ModelSettingsResponse:
    try:
        return ModelSettingsResponse.model_validate(model_settings_view())
    except ModelConnectionStoreError as error:
        raise service_error(
            500,
            "MODEL_CONNECTION_STORE_INVALID",
            str(error),
            retryable=False,
        ) from error


@app.put("/api/settings/model", response_model=ModelSettingsResponse)
async def save_model_settings(
    request: UpdateModelSettingsRequest,
) -> ModelSettingsResponse:
    try:
        return ModelSettingsResponse.model_validate(
            await run_in_threadpool(update_model_api_key, request.apiKey)
        )
    except OSError as error:
        raise service_error(
            500,
            "MODEL_SETTINGS_WRITE_FAILED",
            "无法保存 API Key，请检查本地配置文件是否可写。",
            retryable=False,
        ) from error


@app.post(
    "/api/settings/model/connections",
    response_model=ModelSettingsResponse,
)
async def save_model_connection(
    request: UpsertModelConnectionRequest,
) -> ModelSettingsResponse:
    try:
        result = await run_in_threadpool(
            upsert_managed_connection,
            connection_id=request.id,
            label=request.label,
            base_url=request.baseUrl,
            model=request.modelId,
            api_key=request.apiKey,
            clear_api_key=request.clearApiKey,
            supports_vision=request.supportsVision,
        )
        return ModelSettingsResponse.model_validate(result)
    except ModelConnectionNotFoundError as error:
        raise service_error(
            404,
            "MODEL_CONNECTION_NOT_FOUND",
            str(error),
            retryable=False,
        ) from error
    except DuplicateModelConnectionError as error:
        raise service_error(
            409,
            "MODEL_CONNECTION_DUPLICATE",
            str(error),
            retryable=False,
        ) from error
    except ModelConnectionStoreError as error:
        raise service_error(
            500,
            "MODEL_CONNECTION_STORE_INVALID",
            str(error),
            retryable=False,
        ) from error
    except OSError as error:
        raise service_error(
            500,
            "MODEL_SETTINGS_WRITE_FAILED",
            "无法保存模型连接，请检查本地配置目录是否可写。",
            retryable=False,
        ) from error


@app.put(
    "/api/settings/model/formula",
    response_model=ModelSettingsResponse,
)
async def save_formula_model(
    request: SetFormulaModelRequest,
) -> ModelSettingsResponse:
    """设置 /function 公式专用模型（空串=跟随全局选择）。"""
    try:
        result = await run_in_threadpool(set_formula_model_id, request.modelId)
        return ModelSettingsResponse.model_validate(result)
    except ValueError as error:
        raise service_error(
            400,
            "FORMULA_MODEL_INVALID",
            str(error),
            retryable=False,
        ) from error
    except ModelConnectionStoreError as error:
        raise service_error(
            500,
            "MODEL_CONNECTION_STORE_INVALID",
            str(error),
            retryable=False,
        ) from error
    except OSError as error:
        raise service_error(
            500,
            "MODEL_SETTINGS_WRITE_FAILED",
            "无法保存公式模型选择，请检查本地配置目录是否可写。",
            retryable=False,
        ) from error


@app.post(
    "/api/settings/model/connections/test",
    response_model=TestModelConnectionResponse,
)
async def check_model_connection(
    request: UpsertModelConnectionRequest,
) -> TestModelConnectionResponse:
    try:
        connection = draft_model_connection(
            connection_id=request.id,
            base_url=request.baseUrl,
            model=request.modelId,
            api_key=request.apiKey,
            clear_api_key=request.clearApiKey,
            supports_vision=request.supportsVision,
        )
        await test_model_connection(connection)
        return TestModelConnectionResponse(
            ok=True,
            message="连接成功，服务地址、模型 ID 和 API Key 可用。",
        )
    except ModelConnectionNotFoundError as error:
        raise service_error(
            404,
            "MODEL_CONNECTION_NOT_FOUND",
            str(error),
            retryable=False,
        ) from error
    except ModelConnectionStoreError as error:
        raise service_error(
            500,
            "MODEL_CONNECTION_STORE_INVALID",
            str(error),
            retryable=False,
        ) from error
    except LLMTimeoutError as error:
        raise service_error(
            504,
            "MODEL_CONNECTION_TEST_TIMEOUT",
            "测试连接超时，请检查服务地址或稍后重试。",
            retryable=True,
        ) from error
    except LLMConnectError as error:
        raise service_error(
            502,
            "MODEL_CONNECTION_TEST_UNREACHABLE",
            "无法连接模型服务，请检查服务地址和网络。",
            retryable=True,
        ) from error
    except LLMHTTPStatusError as error:
        message = (
            "模型服务拒绝了测试请求，请检查 API Key、模型 ID 和服务地址。"
        )
        raise service_error(
            502,
            "MODEL_CONNECTION_TEST_REJECTED",
            f"{message}（HTTP {error.status_code}）",
            retryable=error.retryable,
        ) from error
    except (LLMResponseError, LLMTransportError) as error:
        raise service_error(
            502,
            "MODEL_CONNECTION_TEST_FAILED",
            str(error),
            retryable=True,
        ) from error


@app.delete(
    "/api/settings/model/connections/{connection_id}",
    response_model=ModelSettingsResponse,
)
async def remove_model_connection(
    connection_id: str,
) -> ModelSettingsResponse:
    try:
        result = await run_in_threadpool(
            delete_managed_connection,
            connection_id,
        )
        return ModelSettingsResponse.model_validate(result)
    except ModelConnectionNotFoundError as error:
        raise service_error(
            404,
            "MODEL_CONNECTION_NOT_FOUND",
            str(error),
            retryable=False,
        ) from error
    except ModelConnectionStoreError as error:
        raise service_error(
            500,
            "MODEL_CONNECTION_STORE_INVALID",
            str(error),
            retryable=False,
        ) from error
    except OSError as error:
        raise service_error(
            500,
            "MODEL_SETTINGS_WRITE_FAILED",
            "无法删除模型连接，请检查本地配置目录是否可写。",
            retryable=False,
        ) from error


@app.get("/api/capabilities")
async def get_capabilities() -> dict[str, object]:
    return capabilities()


@app.post("/api/turn", response_model=TurnResponse)
async def turn(request: TurnRequest) -> TurnResponse:
    turn_id = request.turnId or f"turn-{uuid.uuid4().hex}"
    try:
        if isinstance(request, IntentCheckRequest):
            turn_registry.ensure_decision_allowed(turn_id)
            response = await check_intent(
                request,
                config=selected_model_config(request.modelId),
            )
            turn_registry.record_decision(turn_id, response)
        else:
            state = turn_registry.touch(turn_id)
            async with state.completion_lock:
                cached = turn_registry.cached_completion(turn_id, request)
                if cached is not None:
                    return cached.model_copy(update={"turnId": turn_id})
                response = await create_plan(
                    request, tool_cache=state.tool_cache
                )
                turn_registry.record_completion(turn_id, request, response)
        return response.model_copy(update={"turnId": turn_id})
    except (
        LLMTimeoutError,
        LLMConnectError,
        LLMHTTPStatusError,
        LLMTransportError,
        LLMResponseError,
        TurnStateError,
        ValueError,
    ) as error:
        status, code, message, retryable = _map_turn_error(error)
        raise service_error(status, code, message, retryable=retryable) from error


def _sse_frame(event: str, data: dict[str, Any]) -> str:
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def _log_task_progress(event: TaskProgressEvent) -> None:
    logger.info(
        "task_progress operation_id=%s execution_id=%s sequence=%s "
        "step_id=%s phase=%s status=%s tool=%s action=%s",
        event.operationId,
        event.executionId or "-",
        event.sequence,
        event.stepId,
        event.phase,
        event.status,
        event.toolName or "-",
        event.actionType or "-",
    )


@dataclass
class _FolderOperationState:
    request_fingerprint: str
    created_at: float
    condition: asyncio.Condition = field(default_factory=asyncio.Condition)
    events: list[dict[str, Any]] = field(default_factory=list)
    sequence: int = 0
    status: str = "running"
    result: dict[str, Any] | None = None
    error: dict[str, Any] | None = None
    task: asyncio.Task[None] | None = None


_folder_operations: dict[str, _FolderOperationState] = {}
_folder_operations_lock = asyncio.Lock()
_folder_operation_ledger = FolderOperationLedger()


def _scan_registered_work_directories() -> list[ManagedDirectoryRecord]:
    scanned: list[ManagedDirectoryRecord] = []
    for registered in _folder_operation_ledger.managed_directories():
        path = Path(registered.path)
        try:
            inspection = inspect_managed_folder_recovery(path)
        except (OSError, ValueError) as error:
            status = "unavailable" if not path.exists() else "unsafe"
            scanned.append(
                _folder_operation_ledger.record_directory_scan(
                    registered.directory_id,
                    status=status,
                    recovery_count=registered.recovery_count,
                    attention_count=max(registered.attention_count, 1),
                    operation_count=registered.operation_count,
                    error=str(error),
                )
            )
            continue
        scanned.append(
            _folder_operation_ledger.record_directory_scan(
                registered.directory_id,
                status="active",
                recovery_count=int(inspection["recoveryCount"]),
                attention_count=int(inspection["attentionCount"]),
                operation_count=int(inspection["operationCount"]),
                error=(
                    "；".join(str(item) for item in inspection["warnings"])[:2000]
                    if inspection["warnings"]
                    else None
                ),
            )
        )
    return scanned


def _register_work_directory(folder_path: str) -> ManagedDirectoryRecord:
    return _folder_operation_ledger.register_directory(
        folder_path,
        max_entries=capability_int("taskRecovery", "maxRegisteredDirectories"),
    )


def _managed_directory_view(
    record: ManagedDirectoryRecord,
) -> ManagedFolderRecoveryDirectory:
    return ManagedFolderRecoveryDirectory(
        directoryId=record.directory_id,
        folderName=Path(record.path).name or record.path,
        folderPath=record.path,
        status=record.status,
        recoveryCount=record.recovery_count,
        attentionCount=record.attention_count,
        operationCount=record.operation_count,
        lastScannedAt=(
            datetime.fromtimestamp(record.last_scanned_at).astimezone().isoformat()
            if record.last_scanned_at is not None
            else None
        ),
        statusReason=record.last_error,
    )


def _folder_state_from_record(record: FolderOperationRecord) -> _FolderOperationState:
    return _FolderOperationState(
        request_fingerprint=record.request_fingerprint,
        created_at=record.created_at,
        events=_folder_operation_ledger.events(record.operation_id),
        sequence=record.last_sequence,
        status=record.status,
        result=record.result,
        error=record.error,
    )


def _task_state_error(error: Exception) -> HTTPException:
    logger.exception("持久任务记录不可用")
    return service_error(
        503,
        "TASK_STATE_UNAVAILABLE",
        "无法安全保存或读取任务状态，已停止执行以避免重复写入。",
        retryable=False,
    )


def _folder_request_fingerprint(request: FolderExecuteRequest) -> str:
    payload = request.model_dump(
        mode="json", exclude={"operationId", "afterSequence"}
    )
    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _folder_stream_error(error: Exception) -> dict[str, Any]:
    return {
        "status": 422,
        "code": "FOLDER_EXECUTION_ERROR",
        "message": str(error),
        "retryable": True,
    }


async def _record_folder_progress(
    operation_id: str,
    state: _FolderOperationState,
    event: dict[str, object],
) -> None:
    next_sequence = state.sequence + 1
    payload = {
        "schemaVersion": 1,
        "operationId": operation_id,
        "sequence": next_sequence,
        "phase": event.get("phase", "executing"),
        "stepId": event.get("stepId", f"folder-{next_sequence}"),
        "status": event.get("status", "progress"),
        "title": event.get("title", "正在执行文件夹操作"),
        "summary": event.get("summary"),
        "detail": event.get("detail"),
        "current": event.get("current"),
        "total": event.get("total"),
        "unit": event.get("unit"),
        "actionIndex": event.get("actionIndex"),
        "actionCount": event.get("actionCount"),
        "actionType": event.get("actionType"),
        "toolName": event.get("toolName"),
        "facts": event.get("facts"),
        "target": event.get("target"),
        "trace": event.get("trace"),
        "executionId": event.get("executionId"),
        "timestamp": datetime.now().astimezone().isoformat(),
    }
    validated_event = TaskProgressEvent.model_validate(payload)
    _log_task_progress(validated_event)
    validated = validated_event.model_dump(mode="json")
    execution_id = event.get("executionId")
    _folder_operation_ledger.append_event(
        operation_id,
        next_sequence,
        validated,
        execution_id=execution_id if isinstance(execution_id, str) else None,
    )
    async with state.condition:
        state.sequence = next_sequence
        state.events.append(validated)
        state.condition.notify_all()


async def _run_folder_operation(
    operation_id: str,
    state: _FolderOperationState,
    request: FolderExecuteRequest,
) -> None:
    loop = asyncio.get_running_loop()

    def on_progress(event: dict[str, object]) -> None:
        future = asyncio.run_coroutine_threadsafe(
            _record_folder_progress(operation_id, state, event), loop
        )
        future.result(timeout=10)

    try:
        result = await run_in_threadpool(execute_folder_plan, request, on_progress)
    except Exception as error:
        expected_error = isinstance(error, (OSError, ValueError, RuntimeError))
        if not expected_error:
            logger.exception(
                "文件夹执行出现未预期异常 operation_id=%s", operation_id
            )
        public_message = (
            str(error)
            if expected_error
            else "文件夹执行发生未预期错误，请查看诊断日志。"
        )
        try:
            await _record_folder_progress(
                operation_id,
                state,
                {
                    "phase": "executing",
                    "stepId": "folder.failed",
                    "status": "failed",
                    "title": "文件夹操作执行失败",
                    "summary": public_message[:1000],
                },
            )
        except Exception:
            logger.exception(
                "文件夹执行失败后无法持久保存进度 operation_id=%s", operation_id
            )
        error_payload = {
            "status": 409 if expected_error else 500,
            "code": (
                "FOLDER_EXECUTION_REQUIRES_RECONCILIATION"
                if expected_error
                else "INTERNAL_ERROR"
            ),
            "message": (
                f"{public_message} 为避免重复写入，请先核对任务状态。"
                if expected_error
                else public_message
            ),
            "retryable": False,
        }
        try:
            _folder_operation_ledger.finish(
                operation_id,
                status="needs_reconciliation",
                error=error_payload,
            )
        except Exception:
            logger.exception(
                "文件夹执行失败后无法持久提交状态 operation_id=%s", operation_id
            )
        async with state.condition:
            state.error = error_payload
            state.status = "needs_reconciliation"
            state.condition.notify_all()
        return

    result_payload = result.model_dump(mode="json")
    execution_id = result_payload.get("executionId")
    undo = result_payload.get("undo")
    undo_id = undo.get("undoId") if isinstance(undo, dict) else None
    try:
        _folder_operation_ledger.finish(
            operation_id,
            status="completed",
            result=result_payload,
            execution_id=execution_id if isinstance(execution_id, str) else None,
            undo_id=undo_id if isinstance(undo_id, str) else None,
        )
    except Exception:
        logger.exception(
            "文件夹写入完成但无法持久提交完成状态 operation_id=%s", operation_id
        )
        persisted: FolderOperationRecord | None = None
        try:
            persisted = _folder_operation_ledger.get(operation_id)
        except Exception:
            logger.exception(
                "文件夹写入完成后无法复核持久状态 operation_id=%s", operation_id
            )
        if persisted is not None and persisted.status == "completed":
            async with state.condition:
                state.result = persisted.result or result_payload
                state.status = "completed"
                state.condition.notify_all()
            return
        error_payload = {
            "status": 503,
            "code": "TASK_STATE_UNAVAILABLE",
            "message": "文件可能已经写入，但完成状态未能安全保存。请先核对任务，系统不会自动重试。",
            "retryable": False,
        }
        try:
            _folder_operation_ledger.finish(
                operation_id,
                status="needs_reconciliation",
                error=error_payload,
            )
        except Exception:
            logger.exception(
                "无法把状态提交失败的文件夹任务标记为待核对 operation_id=%s",
                operation_id,
            )
        async with state.condition:
            state.error = error_payload
            state.status = "needs_reconciliation"
            state.condition.notify_all()
        return

    async with state.condition:
        state.result = result_payload
        state.status = "completed"
        state.condition.notify_all()


def _prune_folder_operations(now: float) -> None:
    ttl = capability_int("progress", "folderOperationTtlSeconds")
    expired = [
        operation_id
        for operation_id, state in _folder_operations.items()
        if state.status != "running" and now - state.created_at > ttl
    ]
    for operation_id in expired:
        _folder_operations.pop(operation_id, None)


@app.post("/api/turn/stream")
async def turn_stream(request: TurnRequest) -> StreamingResponse:
    """plan 分支的流式版本：实时推送分步进度，最后推送完整结果或错误事件。

    仅服务 plan 请求；intent 判断很快，不走流式。HTTP 已返回 200，模型错误
    通过流内 error 事件表达，字段与 /api/turn 的 service_error 完全一致。
    """
    if isinstance(request, IntentCheckRequest):
        raise service_error(
            422,
            "INVALID_REQUEST",
            "流式端点仅用于规划请求，意图判断请调用 /api/turn。",
            retryable=False,
        )

    turn_id = request.turnId or f"turn-{uuid.uuid4().hex}"

    async def generate() -> AsyncIterator[str]:
        try:
            state = turn_registry.touch(turn_id)
            async with state.completion_lock:
                cached = turn_registry.cached_completion(turn_id, request)
                if cached is not None:
                    result = cached.model_copy(update={"turnId": turn_id})
                    yield _sse_frame("result", result.model_dump(mode="json"))
                    return

                queue: asyncio.Queue[dict[str, Any]] = asyncio.Queue()
                event_sequence = 0

                async def on_event(event: dict[str, Any]) -> None:
                    nonlocal event_sequence
                    event_sequence += 1
                    payload = {
                        "schemaVersion": 1,
                        "operationId": turn_id,
                        "sequence": event_sequence,
                        "phase": event.get("phase", "planning"),
                        "stepId": event.get(
                            "stepId", f"planning-{event_sequence}"
                        ),
                        "status": event.get("status", "progress"),
                        "title": event.get("title", "正在处理"),
                        "summary": event.get("summary"),
                        "detail": event.get("detail"),
                        "current": event.get("current"),
                        "total": event.get("total"),
                        "unit": event.get("unit"),
                        "actionIndex": event.get("actionIndex"),
                        "actionCount": event.get("actionCount"),
                        "actionType": event.get("actionType"),
                        "toolName": event.get("toolName"),
                        "facts": event.get("facts"),
                        "target": event.get("target"),
                        "trace": event.get("trace"),
                        "timestamp": datetime.now().astimezone().isoformat(),
                    }
                    validated = TaskProgressEvent.model_validate(payload)
                    _log_task_progress(validated)
                    await queue.put(validated.model_dump(mode="json"))

                task = asyncio.create_task(
                    create_plan(
                        request,
                        tool_cache=state.tool_cache,
                        on_event=on_event,
                    )
                )
                while True:
                    drain = asyncio.ensure_future(queue.get())
                    done, _pending = await asyncio.wait(
                        {drain, task},
                        return_when=asyncio.FIRST_COMPLETED,
                    )
                    if drain in done:
                        yield _sse_frame("step", drain.result())
                        continue
                    drain.cancel()
                    break
                # task 已完成：先排空剩余 step，再产出结果。
                while not queue.empty():
                    yield _sse_frame("step", queue.get_nowait())
                response = await task
                turn_registry.record_completion(turn_id, request, response)
                result = response.model_copy(update={"turnId": turn_id})
                yield _sse_frame("result", result.model_dump(mode="json"))
        except (
            LLMTimeoutError,
            LLMConnectError,
            LLMHTTPStatusError,
            LLMTransportError,
            LLMResponseError,
            TurnStateError,
            ValueError,
        ) as error:
            status, code, message, retryable = _map_turn_error(error)
            yield _sse_frame(
                "error",
                {
                    "status": status,
                    "code": code,
                    "message": message,
                    "retryable": retryable,
                },
            )

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.post("/api/folders/select", response_model=FolderCatalog | None)
async def select_folder() -> FolderCatalog | None:
    catalog = await run_in_threadpool(select_and_scan_folder)
    if catalog is not None:
        try:
            await run_in_threadpool(_register_work_directory, catalog.folderPath)
        except Exception as error:
            raise _task_state_error(error) from error
    return catalog


@app.post("/api/folders/scan", response_model=FolderCatalog)
async def scan_folder_endpoint(request: FolderScanRequest) -> FolderCatalog:
    """按路径扫描文件夹（不弹对话框），供前端在会话丢失后静默重建。"""
    try:
        catalog = await run_in_threadpool(scan_folder_by_path, request.folderPath)
        await run_in_threadpool(_register_work_directory, catalog.folderPath)
        return catalog
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_DATA_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/folders/refresh", response_model=FolderCatalog)
async def refresh_folder_endpoint(
    request: FolderRefreshRequest,
) -> FolderCatalog:
    try:
        return await run_in_threadpool(refresh_folder, request.sessionId)
    except ValueError as error:
        raise service_error(
            422, "FOLDER_DATA_ERROR", str(error), retryable=True
        ) from error


@app.get(
    "/api/folders/recovery/directories",
    response_model=ManagedFolderRecoveryDirectoriesResponse,
)
async def managed_folder_recovery_directories(
) -> ManagedFolderRecoveryDirectoriesResponse:
    try:
        records = await run_in_threadpool(_scan_registered_work_directories)
    except Exception as error:
        raise _task_state_error(error) from error
    return ManagedFolderRecoveryDirectoriesResponse(
        directories=[
            _managed_directory_view(record)
            for record in records
            if record.status != "removed"
        ]
    )


@app.delete("/api/folders/recovery/directories/{directory_id}")
async def remove_managed_folder_recovery_directory(
    directory_id: str,
) -> dict[str, object]:
    if len(directory_id) != 64 or any(
        character not in "0123456789abcdef" for character in directory_id
    ):
        raise service_error(
            422, "INVALID_REQUEST", "工作目录标识无效。", retryable=False
        )
    try:
        removed = await run_in_threadpool(
            _folder_operation_ledger.remove_directory, directory_id
        )
    except Exception as error:
        raise _task_state_error(error) from error
    if not removed:
        raise service_error(
            404,
            "MANAGED_DIRECTORY_NOT_FOUND",
            "未找到已登记的工作目录。",
            retryable=False,
        )
    return {"directoryId": directory_id, "removed": True}


@app.post(
    "/api/folders/recovery/restore",
    response_model=FolderManagedRecoveryResponse,
)
async def restore_managed_folder_recovery(
    request: FolderManagedRecoveryRequest,
) -> FolderManagedRecoveryResponse:
    try:
        directory = await run_in_threadpool(
            _folder_operation_ledger.managed_directory, request.directoryId
        )
    except Exception as error:
        raise _task_state_error(error) from error
    if directory is None or directory.status == "removed":
        raise service_error(
            404,
            "MANAGED_DIRECTORY_NOT_FOUND",
            "未找到已登记的工作目录，请重新选择原目录。",
            retryable=False,
        )
    if directory.status != "active":
        raise service_error(
            409,
            "MANAGED_DIRECTORY_UNAVAILABLE",
            directory.last_error or "原工作目录当前不可用，请重新选择并核对。",
            retryable=False,
        )

    recovery_fingerprint = hashlib.sha256(
        json.dumps(
            {
                "directoryId": request.directoryId,
                "undoId": request.undoId,
                "expectedRevision": request.expectedRevision,
                "direction": "restore_before",
            },
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    ).hexdigest()
    try:
        operation, created = await run_in_threadpool(
            _folder_operation_ledger.register,
            request.operationId,
            recovery_fingerprint,
            work_directory=directory.path,
            plan_id=f"recovery:{request.undoId}",
        )
    except ExpiredOperationError as error:
        raise service_error(
            410,
            "OPERATION_EXPIRED",
            "这个恢复操作已超过保留期，不能用原标识重新执行。请刷新恢复列表后重新确认。",
            retryable=False,
        ) from error
    except Exception as error:
        raise _task_state_error(error) from error
    if operation.request_fingerprint != recovery_fingerprint:
        raise service_error(
            409,
            "OPERATION_ID_CONFLICT",
            "这个恢复操作标识已用于其他请求，请刷新恢复列表。",
            retryable=False,
        )
    if operation.status == "completed" and operation.result is not None:
        return FolderManagedRecoveryResponse.model_validate(operation.result)
    if not created and operation.status == "running":
        # 相同恢复身份允许可重入：若上一请求仍在执行，文件夹排他锁会返回 busy；
        # 若文件已恢复但完成事务未提交，效果核对会识别执行前指纹并补记完成状态。
        pass
    elif operation.status in {"failed", "needs_reconciliation"}:
        try:
            operation = await run_in_threadpool(
                _folder_operation_ledger.reopen_for_recovery,
                request.operationId,
                recovery_fingerprint,
            )
        except Exception as error:
            raise _task_state_error(error) from error
    elif operation.status != "running":
        raise service_error(
            409,
            "RECOVERY_OPERATION_CONFLICT",
            "恢复操作状态已变化，请刷新后核对。",
            retryable=False,
        )

    try:
        restored = await run_in_threadpool(
            restore_managed_folder_backup,
            Path(directory.path),
            undo_id=request.undoId,
            expected_revision=request.expectedRevision,
            recovery_operation_id=request.operationId,
        )
    except FolderWriteBusyError as error:
        raise service_error(
            409, "RECOVERY_OPERATION_BUSY", str(error), retryable=True
        ) from error
    except (OSError, ValueError, RuntimeError) as error:
        error_payload = {
            "status": 409,
            "code": "RECOVERY_REQUIRES_RECONCILIATION",
            "message": str(error),
            "retryable": False,
        }
        try:
            await run_in_threadpool(
                _folder_operation_ledger.finish,
                request.operationId,
                status="needs_reconciliation",
                error=error_payload,
                undo_id=request.undoId,
            )
        except Exception:
            logger.exception(
                "无法持久提交恢复待核对状态 operation_id=%s", request.operationId
            )
        raise service_error(
            409,
            "RECOVERY_REQUIRES_RECONCILIATION",
            str(error),
            retryable=False,
        ) from error

    response = FolderManagedRecoveryResponse(
        operationId=request.operationId,
        undoId=request.undoId,
        status="completed",
        filesRestored=restored.filesRestored,
        alreadyRestored=not restored.filesRestored,
    )
    payload = response.model_dump(mode="json")
    try:
        await run_in_threadpool(
            _folder_operation_ledger.finish,
            request.operationId,
            status="completed",
            result=payload,
            undo_id=request.undoId,
        )
    except Exception as error:
        try:
            persisted = await run_in_threadpool(
                _folder_operation_ledger.get, request.operationId
            )
        except Exception:
            persisted = None
        if (
            persisted is not None
            and persisted.status == "completed"
            and persisted.result is not None
        ):
            return FolderManagedRecoveryResponse.model_validate(persisted.result)
        raise _task_state_error(error) from error
    return response


@app.post("/api/folders/snapshot", response_model=WorkbookSnapshot)
async def folder_snapshot(request: FolderSnapshotRequest) -> WorkbookSnapshot:
    try:
        snapshot = await run_in_threadpool(create_folder_snapshot, request)
        return snapshot
    except ValueError as error:
        raise service_error(
            422, "FOLDER_DATA_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/folders/execute", response_model=FolderExecuteResponse)
async def folder_execute(request: FolderExecuteRequest) -> FolderExecuteResponse:
    raise service_error(
        409,
        "IDEMPOTENT_EXECUTION_REQUIRED",
        "此兼容入口不再接受写入；请使用带稳定 operationId 的流式执行入口。",
        retryable=False,
    )


@app.post("/api/folders/execute/stream")
async def folder_execute_stream(request: FolderExecuteRequest) -> StreamingResponse:
    """流式执行文件夹计划；同一 operationId 只执行一次并可重连追赶事件。"""
    operation_id = request.operationId
    if operation_id is None:
        raise service_error(
            422,
            "INVALID_REQUEST",
            "文件夹流式执行缺少 operationId。",
            retryable=False,
        )
    request_fingerprint = _folder_request_fingerprint(request)
    async with _folder_operations_lock:
        _prune_folder_operations(time.time())
        state = _folder_operations.get(operation_id)
        if state is not None and state.request_fingerprint != request_fingerprint:
            raise service_error(
                409,
                "OPERATION_ID_CONFLICT",
                "这个 operationId 已用于另一份执行请求，请重新生成操作标识。",
                retryable=False,
            )
        if state is None:
            try:
                existing_record = _folder_operation_ledger.get(operation_id)
            except Exception as error:
                raise _task_state_error(error) from error
            if existing_record is None:
                try:
                    expired_fingerprint = _folder_operation_ledger.expired_fingerprint(
                        operation_id
                    )
                except Exception as error:
                    raise _task_state_error(error) from error
                if expired_fingerprint is not None:
                    raise service_error(
                        410,
                        "OPERATION_EXPIRED",
                        "这个执行任务已超过结果保留期；为避免重复写入，不会重新创建执行。",
                        retryable=False,
                    )
            if existing_record is not None:
                if existing_record.request_fingerprint != request_fingerprint:
                    raise service_error(
                        409,
                        "OPERATION_ID_CONFLICT",
                        "这个 operationId 已用于另一份执行请求，请重新生成操作标识。",
                        retryable=False,
                    )
                if existing_record.status in {"completed", "failed"}:
                    try:
                        state = _folder_state_from_record(existing_record)
                    except Exception as error:
                        raise _task_state_error(error) from error
                    _folder_operations[operation_id] = state
                else:
                    raise service_error(
                        409,
                        "OPERATION_REQUIRES_RECONCILIATION",
                        "上次执行没有留下可确认的完成状态。为避免重复写入，已停止自动重试，请先核对任务。",
                        retryable=False,
                    )
        if state is None:
            max_operations = capability_int(
                "progress", "folderMaxPendingOperations"
            )
            if len(_folder_operations) >= max_operations:
                completed = sorted(
                    (
                        (key, value)
                        for key, value in _folder_operations.items()
                        if value.status != "running"
                    ),
                    key=lambda item: item[1].created_at,
                )
                for key, _value in completed:
                    if len(_folder_operations) < max_operations:
                        break
                    _folder_operations.pop(key, None)
            if len(_folder_operations) >= max_operations:
                raise service_error(
                    429,
                    "TOO_MANY_OPERATIONS",
                    "当前文件夹执行任务过多，请稍后重试。",
                    retryable=True,
                )
            try:
                work_directory = str(folder_session_root(request.sessionId))
            except ValueError as error:
                raise service_error(
                    422, "FOLDER_DATA_ERROR", str(error), retryable=True
                ) from error
            try:
                record, created = _folder_operation_ledger.register(
                    operation_id,
                    request_fingerprint,
                    work_directory=work_directory,
                    session_id=request.sessionId,
                    plan_id=request.plan.id,
                )
            except ExpiredOperationError as error:
                raise service_error(
                    410,
                    "OPERATION_EXPIRED",
                    "这个执行任务已超过结果保留期；为避免重复写入，不会重新创建执行。",
                    retryable=False,
                ) from error
            except Exception as error:
                raise _task_state_error(error) from error
            if record.request_fingerprint != request_fingerprint:
                raise service_error(
                    409,
                    "OPERATION_ID_CONFLICT",
                    "这个 operationId 已用于另一份执行请求，请重新生成操作标识。",
                    retryable=False,
                )
            if created:
                state = _FolderOperationState(
                    request_fingerprint=request_fingerprint,
                    created_at=record.created_at,
                )
                _folder_operations[operation_id] = state
                state.task = asyncio.create_task(
                    _run_folder_operation(operation_id, state, request)
                )
            elif record.status in {"completed", "failed"}:
                try:
                    state = _folder_state_from_record(record)
                except Exception as error:
                    raise _task_state_error(error) from error
                _folder_operations[operation_id] = state
            else:
                raise service_error(
                    409,
                    "OPERATION_REQUIRES_RECONCILIATION",
                    "上次执行没有留下可确认的完成状态。为避免重复写入，已停止自动重试，请先核对任务。",
                    retryable=False,
                )

    async def generate() -> AsyncIterator[str]:
        event_index = min(request.afterSequence, len(state.events))
        heartbeat_seconds = capability_int("progress", "heartbeatSeconds")
        while True:
            heartbeat = False
            async with state.condition:
                if event_index >= len(state.events) and state.status == "running":
                    try:
                        await asyncio.wait_for(
                            state.condition.wait(), timeout=heartbeat_seconds
                        )
                    except TimeoutError:
                        heartbeat = True
                pending = state.events[event_index:]
                event_index = len(state.events)
                status = state.status
                result = state.result
                error = state.error
            for event in pending:
                yield _sse_frame("progress", event)
            if status == "completed" and result is not None:
                yield _sse_frame("result", result)
                return
            if status in {"failed", "needs_reconciliation"} and error is not None:
                yield _sse_frame("error", error)
                return
            if heartbeat and not pending:
                yield ": heartbeat\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/folders/execute/status")
async def folder_execute_status(operationId: str) -> dict[str, Any]:
    """返回持久文件夹操作状态；查询本身绝不创建执行。"""
    async with _folder_operations_lock:
        _prune_folder_operations(time.time())
        state = _folder_operations.get(operationId)
        if state is None:
            try:
                record = _folder_operation_ledger.get(operationId)
            except Exception as error:
                raise _task_state_error(error) from error
            if record is None:
                try:
                    expired_fingerprint = _folder_operation_ledger.expired_fingerprint(
                        operationId
                    )
                except Exception as error:
                    raise _task_state_error(error) from error
                if expired_fingerprint is not None:
                    raise service_error(
                        410,
                        "OPERATION_EXPIRED",
                        "这个执行任务已超过结果保留期；查询不会创建新的执行。",
                        retryable=False,
                    )
                raise service_error(
                    404,
                    "OPERATION_NOT_FOUND",
                    "未找到该文件夹执行任务；查询不会创建新的执行。",
                    retryable=False,
                )
            try:
                state = _folder_state_from_record(record)
            except Exception as error:
                raise _task_state_error(error) from error
            if state.status == "running":
                state.status = "needs_reconciliation"
                state.error = {
                    "status": 409,
                    "code": "OPERATION_REQUIRES_RECONCILIATION",
                    "message": "执行现场已不在当前进程中，需要核对后才能继续。",
                    "retryable": False,
                }
        return {
            "operationId": operationId,
            "status": state.status,
            "lastSequence": state.sequence,
            "result": state.result,
            "error": state.error,
        }


@app.post("/api/folders/undo", response_model=FolderUndoResponse)
async def folder_undo(request: FolderUndoRequest) -> FolderUndoResponse:
    try:
        return await run_in_threadpool(undo_folder_execution, request)
    except (OSError, ValueError, RuntimeError) as error:
        raise service_error(
            422, "FOLDER_UNDO_ERROR", str(error), retryable=True
        ) from error


@app.get("/api/folders/backups", response_model=FolderBackupsResponse)
async def folder_backups(
    sessionId: str | None = None, folderPath: str | None = None
) -> FolderBackupsResponse:
    """列备份。sessionId 与 folderPath 二选一：前者带撤销能力，后者只读。"""
    try:
        response = await run_in_threadpool(list_folder_backups, sessionId, folderPath)
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_BACKUP_ERROR", str(error), retryable=True
        ) from error
    try:
        await run_in_threadpool(_register_work_directory, response.folderPath)
    except Exception:
        # 目录索引只负责重启发现；索引暂时不可写时不应阻断普通备份查看和导出。
        logger.exception("打开备份管理时无法登记工作目录 path=%s", response.folderPath)
    return response


@app.post("/api/folders/backups/open", response_model=FolderBackupOpenResponse)
async def folder_backup_open(
    request: FolderBackupOpenRequest,
) -> FolderBackupOpenResponse:
    try:
        return await run_in_threadpool(
            open_folder_backup_directory, request.directoryToken
        )
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_BACKUP_OPEN_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/folders/backups/export", response_model=FolderBackupExportResponse)
async def folder_backup_export(
    request: FolderBackupExportRequest,
) -> FolderBackupExportResponse:
    try:
        return await run_in_threadpool(
            export_folder_backup,
            request.sessionId,
            request.undoId,
            backup_management_token=request.backupManagementToken,
            operation_id=request.operationId,
        )
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_BACKUP_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/folders/backups/delete", response_model=FolderBackupDeleteResponse)
async def folder_backup_delete(
    request: FolderBackupDeleteRequest,
) -> FolderBackupDeleteResponse:
    try:
        return await run_in_threadpool(
            delete_folder_backup,
            request.backupManagementToken,
            request.undoId,
            confirmed=request.confirmed,
            expected_revision=request.expectedRevision,
            expected_invalidated_undo_ids=request.expectedInvalidatedUndoIds,
        )
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_BACKUP_DELETE_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/folders/backups/pin", response_model=FolderBackupPinResponse)
async def folder_backup_pin(request: FolderBackupPinRequest) -> FolderBackupPinResponse:
    try:
        return await run_in_threadpool(
            set_folder_backup_pinned,
            request.backupManagementToken,
            request.undoId,
            pinned=request.pinned,
        )
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_BACKUP_PIN_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/folders/backups/cleanup", response_model=FolderBackupCleanupResponse)
async def folder_backup_cleanup(
    request: FolderBackupCleanupRequest,
) -> FolderBackupCleanupResponse:
    try:
        return await run_in_threadpool(
            cleanup_folder_backups,
            request.sessionId,
            backup_management_token=request.backupManagementToken,
        )
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_BACKUP_ERROR", str(error), retryable=True
        ) from error


@app.post(
    "/api/folders/backups/legacy/collect",
    response_model=FolderLegacyCollectResponse,
)
async def folder_backup_collect_legacy(
    request: FolderLegacyCollectRequest,
) -> FolderLegacyCollectResponse:
    try:
        return await run_in_threadpool(
            collect_legacy_folder_backups,
            request.sessionId,
            confirmed=request.confirmed,
            backup_management_token=request.backupManagementToken,
        )
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_BACKUP_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/folders/query", response_model=DataToolResult)
async def folder_query(request: FolderQueryRequest) -> DataToolResult:
    try:
        return await run_in_threadpool(query_folder_data, request)
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_DATA_TOOL_ERROR", str(error), retryable=True
        ) from error


@app.post("/api/formulas/generate", response_model=GenerateFormulaResponse)
async def generate_native_formula(request: GenerateFormulaRequest) -> GenerateFormulaResponse:
    """/function 短链：根据描述和表格上下文生成原生 Excel 公式"""
    try:
        return await generate_formula(request)
    except ValueError as error:
        raise service_error(
            422, "FORMULA_GENERATION_ERROR", str(error), retryable=False
        ) from error
    except (
        LLMConnectError,
        LLMHTTPStatusError,
        LLMResponseError,
        LLMTimeoutError,
        LLMTransportError,
    ) as error:
        raise service_error(
            502, "MODEL_UNAVAILABLE", f"模型服务连接失败：{error}", retryable=True
        ) from error


@app.post(
    "/api/plan",
    response_model=AssistantResponse,
    response_model_exclude_none=True,
)
async def plan(request: PlanRequest) -> AssistantResponse:
    try:
        return await create_plan(request)
    except (
        LLMConnectError,
        LLMHTTPStatusError,
        LLMResponseError,
        LLMTimeoutError,
        LLMTransportError,
    ) as error:
        raise HTTPException(status_code=502, detail=f"模型服务连接失败：{error}") from error
    except ValueError as error:
        raise service_error(
            422, "INVALID_REQUEST", str(error), retryable=False
        ) from error


@app.post("/api/plans/validate", response_model=AnalysisPlan)
async def validate_plan(request: PlanValidateRequest) -> AnalysisPlan:
    """校验并规范化一份手工编辑过的计划。

    保存编辑后的计划必须过这一关，返回的是**规范化之后**的计划（验收条件、
    去重后的 warnings 等都由 AnalysisPlan 的校验器重算）。前端不允许自己拼
    acceptanceCriteria —— 计划内容、liveness、指纹、验收条件这四样只要有两样
    对不上，就会出现「保存成功但执行时验证失败」这种最难查的沉默错。

    FastAPI 会在请求体不合法时直接返回 422 + Pydantic 的逐字段报错，正好是
    编辑器需要的反馈，所以这里不额外包装。
    """
    return request.plan


class PlanChatRequest(BaseModel):
    plan: AnalysisPlan
    message: str
    modelId: str | None = None


class PlanChatResponse(BaseModel):
    explanation: str
    modifiedPlan: AnalysisPlan


@app.post("/api/plans/chat-modify", response_model=PlanChatResponse)
async def plan_chat_modify(request: PlanChatRequest) -> PlanChatResponse:
    """AI 辅助修改计划。

    用户提供当前计划和自然语言描述，AI 返回修改建议和完整的新计划。
    """
    try:
        config = selected_model_config(request.modelId)
        explanation, modified_plan = await chat_modify_plan(
            current_plan=request.plan,
            user_message=request.message,
            connection=config,
            timeout=model_timeout_seconds(),
        )
        return PlanChatResponse(
            explanation=explanation,
            modifiedPlan=modified_plan
        )
    except (
        LLMTimeoutError,
        LLMConnectError,
        LLMHTTPStatusError,
        LLMTransportError,
        LLMResponseError,
    ) as error:
        status, code, message, retryable = _map_turn_error(error)
        raise service_error(status, code, message, retryable=retryable) from error
    except ValueError as error:
        raise service_error(
            422, "INVALID_PLAN_MODIFICATION", str(error), retryable=False
        ) from error


@app.post(
    "/api/intent",
    response_model=IntentCheckResponse,
    response_model_exclude_none=True,
)
async def intent(request: IntentCheckRequest) -> IntentCheckResponse:
    try:
        return await check_intent(
            request,
            config=selected_model_config(request.modelId),
        )
    except LLMTimeoutError as error:
        raise HTTPException(
            status_code=504,
            detail=(
                "模型意图识别超时。可以重试，或在 server/.env 中提高 "
                "AI_INTENT_TIMEOUT_SECONDS。"
            ),
        ) from error
    except LLMConnectError as error:
        raise HTTPException(
            status_code=502,
            detail="无法连接模型服务，请检查网络和 AI_BASE_URL。",
        ) from error
    except LLMHTTPStatusError as error:
        status = error.status_code
        if status == 429:
            raise HTTPException(
                status_code=503,
                detail=(
                    "模型服务当前繁忙（过载或限流），已自动重试多次仍未成功，"
                    "请稍候再重试。"
                ),
            ) from error
        if status >= 500:
            raise HTTPException(
                status_code=502,
                detail="模型服务暂时不可用（服务端错误），请稍候重试。",
            ) from error
        raise HTTPException(status_code=502, detail=str(error)) from error
    except (LLMTransportError, LLMResponseError) as error:
        raise HTTPException(
            status_code=502,
            detail=f"模型意图识别连接失败：{error}",
        ) from error
    except ValueError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error
    ManagedFolderRecoveryDirectoriesResponse,
    ManagedFolderRecoveryDirectory,

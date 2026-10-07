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

from .capabilities import capabilities, capability_float, capability_int
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
    cleanup_folder_task_artifacts,
    delete_folder_backup,
    execute_folder_plan,
    folder_execution_context,
    folder_resume_checkpoint_summary,
    folder_session_root,
    inspect_managed_folder_recovery,
    export_folder_backup,
    list_folder_backups,
    open_folder_backup_directory,
    resume_folder_plan,
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
    AnswerResponse,
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
    FolderExecutionResumeRequest,
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
    OfficeActionCompleteRequest,
    OfficeActionReconcileRequest,
    OfficeActionStartRequest,
    OfficeExecutionFinishRequest,
    OfficeExecutionRegisterRequest,
    OfficeExecutionResumeRequest,
    OfficeExecutionStateResponse,
    OfficeExecutionSuspendRequest,
    WpsActionCompleteRequest,
    WpsActionReconcileRequest,
    WpsActionStartRequest,
    WpsExecutionFinishRequest,
    WpsExecutionRegisterRequest,
    WpsExecutionResumeRequest,
    WpsExecutionStateResponse,
    WpsExecutionSuspendRequest,
    WpsRecoveryDescriptor,
    PlanRequest,
    PlanningContinueRequest,
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
    ModelNotAvailableError,
    ModelNotConfiguredError,
    resolve_model_config,
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
    OfficeExecutionRecord,
    PlanningTaskRecord,
)
from .excel_agent import AgentBudgetPause
from .planner import create_plan
from .plan_chat import chat_modify_plan
from .rule_generator import (
    GenerateFormulaRequest,
    GenerateFormulaResponse,
    generate_formula,
)
from .turn_state import TurnStateError, turn_registry
from .excel_window_layout import excel_window_layout, WindowLayoutError
from .models import ExcelReviewLayoutPrepareRequest, ExcelReviewLayoutTicket, ExcelReviewLayoutResult, ExcelReviewLayoutArrangeRequest, ExcelReviewPaneMeasurement

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
    "http://localhost:3889",
    "http://127.0.0.1:3889",
    # 兼容修复前由 apps/wps-addin/package.json 启动的 3890 开发页。
    # 新的 WPS 调试入口统一使用 3889；这里保留精确回环来源，避免已打开
    # 的任务窗格在热更新后仍因 CORS 预检失败而只能显示 Failed to fetch。
    "http://localhost:3890",
    "http://127.0.0.1:3890",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=False,
    allow_methods=["DELETE", "GET", "POST", "PUT"],
    allow_headers=["Content-Type"],
)

# Deliberately separate from /plan: volatile authorized rows may only reach a
# local service and are never included in the model/history/checkpoint DTOs.
from .python_data_api import create_python_data_router
app.include_router(create_python_data_router(ALLOWED_ORIGINS))


def _review_layout_origin(request: Request) -> str:
    # This desktop action is narrower than the general planning API: localhost
    # only, explicit Excel task-pane Origin, no no-Origin CLI / folder / WPS use.
    origin = request.headers.get("origin", "")
    if (not request.client or request.client.host not in {"127.0.0.1", "::1"}
        or request.url.hostname not in {"127.0.0.1", "localhost", "::1"}
        or origin not in ALLOWED_ORIGINS[:4]
        or request.headers.get("content-type", "").split(";", 1)[0].strip().lower() != "application/json"):
        raise HTTPException(403, detail={"code": "LAYOUT_FORBIDDEN", "message": "仅允许本机 Excel 任务窗格请求窗口排列。", "retryable": False})
    return origin


@app.post("/api/office/review-layout/prepare", response_model=ExcelReviewLayoutTicket)
async def prepare_excel_review_layout(payload: ExcelReviewLayoutPrepareRequest, request: Request):
    origin = _review_layout_origin(request)
    try:
        token = await run_in_threadpool(excel_window_layout.prepare, payload, origin)
        return ExcelReviewLayoutTicket(token=token, layoutVersion=3)
    except Exception as exc:
        code = exc.code if isinstance(exc, WindowLayoutError) else "LAYOUT_UNAVAILABLE"
        raise HTTPException(409, detail={"code": code, "message": "无法确认原表窗口；请先激活原表，或手动并排查看。", "retryable": False}) from None


@app.post("/api/office/review-layout/measure", response_model=ExcelReviewPaneMeasurement)
async def measure_excel_review_pane(payload: ExcelReviewLayoutPrepareRequest, request: Request):
    _review_layout_origin(request)
    try:
        return await run_in_threadpool(excel_window_layout.measure, payload)
    except Exception as exc:
        code = exc.code if isinstance(exc, WindowLayoutError) else "LAYOUT_UNAVAILABLE"
        raise HTTPException(409, detail={"code": code, "message": "无法确认本次审阅窗口宽度，请手动调整 AI 分隔线。", "retryable": False}) from None


@app.post("/api/office/review-layout/arrange", response_model=ExcelReviewLayoutResult)
async def arrange_excel_review_layout(payload: ExcelReviewLayoutArrangeRequest, request: Request):
    origin = _review_layout_origin(request)
    try:
        return await run_in_threadpool(excel_window_layout.arrange, payload.token, origin)
    except Exception as exc:
        code = exc.code if isinstance(exc, WindowLayoutError) else "LAYOUT_UNAVAILABLE"
        raise HTTPException(409, detail={"code": code, "message": "自动窗口排列未完成；请手动并排查看。原表未因此写入。", "retryable": False}) from None


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
    if isinstance(error, ModelNotConfiguredError):
        return (409, "MODEL_NOT_CONFIGURED", str(error), False)
    if isinstance(error, ModelNotAvailableError):
        return (409, "MODEL_NOT_AVAILABLE", str(error), False)
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
                config=resolve_model_config(request.modelId),
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
        ModelNotConfiguredError,
        ModelNotAvailableError,
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
_active_planning_tasks: set[str] = set()
_active_planning_tasks_lock = asyncio.Lock()


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

    def on_checkpoint(checkpoint: dict[str, object]) -> None:
        _folder_operation_ledger.commit_checkpoint(
            operation_id,
            kind=str(checkpoint.get("kind") or "folder_publish"),
            payload=checkpoint,
        )

    try:
        checkpoint_supported = (
            len(request.plan.actions) == 1
            and request.plan.actions[0].type == "splitGroupAggregate"
        )
        if checkpoint_supported:
            result = await run_in_threadpool(
                execute_folder_plan, request, on_progress, on_checkpoint
            )
        else:
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

    try:
        await run_in_threadpool(
            cleanup_folder_task_artifacts,
            folder_session_root(request.sessionId),
            operation_id,
        )
    except Exception:
        logger.warning(
            "已完成任务的暂存产物仍待清理 operation_id=%s",
            operation_id,
            exc_info=True,
        )

    async with state.condition:
        state.result = result_payload
        state.status = "completed"
        state.condition.notify_all()


async def _run_folder_resume(
    operation_id: str,
    state: _FolderOperationState,
    record: FolderOperationRecord,
    checkpoint_payload: dict[str, Any],
) -> None:
    loop = asyncio.get_running_loop()

    def on_progress(event: dict[str, object]) -> None:
        future = asyncio.run_coroutine_threadsafe(
            _record_folder_progress(operation_id, state, event), loop
        )
        future.result(timeout=10)

    def on_checkpoint(checkpoint: dict[str, object]) -> None:
        _folder_operation_ledger.commit_checkpoint(
            operation_id,
            kind=str(checkpoint.get("kind") or "folder_publish"),
            payload=checkpoint,
        )

    try:
        if record.authorization is None:
            raise ValueError("任务缺少原授权现场，不能续跑")
        result = await run_in_threadpool(
            resume_folder_plan,
            operation_id,
            record.authorization,
            checkpoint_payload,
            on_progress,
            on_checkpoint,
        )
        result_payload = result.model_dump(mode="json")
        undo = result_payload.get("undo")
        undo_id = undo.get("undoId") if isinstance(undo, dict) else None
        _folder_operation_ledger.finish(
            operation_id,
            status="completed",
            result=result_payload,
            execution_id=result.executionId,
            undo_id=undo_id if isinstance(undo_id, str) else None,
        )
    except Exception as error:
        expected_error = isinstance(error, (OSError, ValueError, RuntimeError))
        if not expected_error:
            logger.exception("断点续跑出现未预期异常 operation_id=%s", operation_id)
        message = str(error) if expected_error else "断点续跑发生未预期错误，请查看诊断日志。"
        try:
            await _record_folder_progress(
                operation_id,
                state,
                {
                    "phase": "saving",
                    "stepId": "folder.resume.failed",
                    "status": "failed",
                    "title": "断点续跑已安全停止",
                    "summary": message[:1000],
                },
            )
        except Exception:
            logger.exception("断点续跑失败后无法保存进度 operation_id=%s", operation_id)
        error_payload = {
            "status": 409 if expected_error else 500,
            "code": "FOLDER_RESUME_CONFLICT" if expected_error else "INTERNAL_ERROR",
            "message": message,
            "retryable": False,
        }
        try:
            _folder_operation_ledger.finish(
                operation_id, status="needs_reconciliation", error=error_payload
            )
        except Exception:
            logger.exception("断点续跑失败后无法提交状态 operation_id=%s", operation_id)
        async with state.condition:
            state.error = error_payload
            state.status = "needs_reconciliation"
            state.condition.notify_all()
        return


    try:
        if record.authorization is not None and isinstance(
            record.authorization.get("root"), str
        ):
            await run_in_threadpool(
                cleanup_folder_task_artifacts,
                str(record.authorization["root"]),
                operation_id,
            )
    except Exception:
        logger.warning(
            "已完成续跑任务的暂存产物仍待清理 operation_id=%s",
            operation_id,
            exc_info=True,
        )

    async with state.condition:
        state.result = result_payload
        state.error = None
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


def _planning_request_fingerprint(request: PlanRequest) -> str:
    encoded = request.model_dump_json(
        exclude={"turnId"}, exclude_none=True
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


_OFFICE_RECOVERABLE_ACTION_TYPES = {
    "writeValues",
    "writeFormulas",
    "setFill",
    "setFont",
    "setNumberFormat",
}

# WPS 当前只开放这三类，且 clearRange 只认 contents；R1C1 与结构动作一律拒绝。
_WPS_RECOVERABLE_ACTION_TYPES = {
    "writeValues",
    "writeFormulas",
    "clearRange",
}


def _office_request_fingerprint(request: OfficeExecutionRegisterRequest) -> str:
    payload = request.model_dump(mode="json", exclude={"operationId", "confirmed"})
    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _validate_office_recovery_plan(request: OfficeExecutionRegisterRequest) -> None:
    authorized = {name.casefold() for name in request.authorizedSheetNames}
    source_sheets = {name.casefold() for name in request.plan.sourceFingerprintSheets}
    for index, action in enumerate(request.plan.actions):
        if action.type not in _OFFICE_RECOVERABLE_ACTION_TYPES:
            raise service_error(
                422,
                "OFFICE_RECOVERY_UNSUPPORTED_ACTION",
                f"第 {index + 1} 步 {action.type} 暂不支持跨重启自动续跑。该计划仍可按普通模式执行，但不会声称能够断点恢复。",
                retryable=False,
            )
        if action.sheet.casefold() not in authorized:
            raise service_error(
                403,
                "OFFICE_RECOVERY_SCOPE_MISMATCH",
                f"第 {index + 1} 步目标工作表「{action.sheet}」不在本次授权范围内。",
                retryable=False,
            )
        if action.sheet.casefold() in source_sheets:
            raise service_error(
                422,
                "OFFICE_RECOVERY_SOURCE_TARGET_OVERLAP",
                f"第 {index + 1} 步会修改来源工作表，暂不能安全区分任务自身修改与恢复期间的外部修改。",
                retryable=False,
            )
        if action.type == "writeFormulas" and action.formulaR1C1:
            raise service_error(
                422,
                "OFFICE_RECOVERY_UNSUPPORTED_ACTION",
                "R1C1 铺满公式暂不支持跨重启效果核对。",
                retryable=False,
            )


def _office_state_payload(
    record: OfficeExecutionRecord, *, include_owner: bool = False
) -> OfficeExecutionStateResponse:
    authorized = record.authorization.get("authorizedSheetNames", [])
    return OfficeExecutionStateResponse(
        operationId=record.operation_id,
        status=record.status,
        revision=record.revision,
        workbookIdentity=record.workbook_identity,
        plan=AnalysisPlan.model_validate(record.plan),
        authorizedSheetNames=[str(value) for value in authorized],
        allowedExternalFiles=[
            str(value) for value in record.authorization.get("allowedExternalFiles", [])
        ],
        nextActionIndex=record.next_action_index,
        inFlightActionIndex=record.in_flight_action_index,
        beforeSnapshot=record.before_snapshot,
        checkpoint=record.checkpoint,
        result=record.result,
        error=record.error,
        canResume=record.status in {"running", "failed", "needs_reconciliation"},
        requiresReconciliation=record.in_flight_action_index is not None,
        ownerToken=record.owner_token if include_owner else None,
    )


def _wps_request_fingerprint(request: WpsExecutionRegisterRequest) -> str:
    payload = request.model_dump(mode="json", exclude={"operationId", "confirmed"})
    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _validate_wps_recovery_descriptor(request: WpsExecutionRegisterRequest) -> None:
    authorized = {name.casefold() for name in request.authorizedSheetNames}
    if sum(action.cellCount for action in request.descriptor.actions) > capability_int("wpsWrite", "maxCells"):
        raise service_error(422, "WPS_RECOVERY_INVALID", "恢复目标超过 WPS 写入单元格限额。", retryable=False)
    for index, action in enumerate(request.descriptor.actions):
        if action.actionIndex != index or action.cellCount != action.rowCount * action.columnCount:
            raise service_error(422, "WPS_RECOVERY_INVALID", "恢复动作序号或矩形尺寸不一致。", retryable=False)
        if action.type not in _WPS_RECOVERABLE_ACTION_TYPES:
            raise service_error(
                422,
                "WPS_RECOVERY_UNSUPPORTED_ACTION",
                f"第 {index + 1} 步 {action.type} 暂不支持 WPS 跨重启自动续跑。",
                retryable=False,
            )
        if action.sheet.casefold() not in authorized:
            raise service_error(
                403,
                "WPS_RECOVERY_SCOPE_MISMATCH",
                f"第 {index + 1} 步目标工作表「{action.sheet}」不在本次授权范围内。",
                retryable=False,
            )
        if action.type == "clearRange" and action.applyTo != "contents":
            raise service_error(
                422,
                "WPS_RECOVERY_UNSUPPORTED_ACTION",
                "WPS 只支持清空明确区域的内容（applyTo=contents）。",
                retryable=False,
            )
        if action.type == "writeFormulas" and action.formulaR1C1:
            raise service_error(
                422,
                "WPS_RECOVERY_UNSUPPORTED_ACTION",
                "R1C1 铺满公式暂不支持 WPS 跨重启效果核对。",
                retryable=False,
            )


def _wps_state_payload(
    record: OfficeExecutionRecord, *, include_owner: bool = False
) -> WpsExecutionStateResponse:
    authorized = record.authorization.get("authorizedSheetNames", [])
    descriptor = WpsRecoveryDescriptor.model_validate(record.plan)
    owner_active = record.status == "running" and record.updated_at + capability_int(
        "taskRecovery", "wpsOwnerLeaseSeconds"
    ) > time.time()
    return WpsExecutionStateResponse(
        operationId=record.operation_id,
        status=record.status,
        revision=record.revision,
        workbookIdentity=record.workbook_identity,
        taskId=record.task_id or "",
        # 账本里存的是恢复描述，不是完整计划；下发状态因此也不含单元格矩阵。
        descriptor=descriptor,
        authorizedSheetNames=[str(value) for value in authorized],
        nextActionIndex=record.next_action_index,
        totalActions=len(descriptor.actions),
        inFlightActionIndex=record.in_flight_action_index,
        beforeSnapshot=record.before_snapshot,
        checkpoint=record.checkpoint,
        result=record.result,
        error={"code": "WPS_WRITE_INTERRUPTED", "message": "WPS 写入中断，需要核对现场。"} if record.error else None,
        canResume=not owner_active and record.status in {"running", "failed", "needs_reconciliation"},
        requiresReconciliation=record.in_flight_action_index is not None,
        ownerActive=owner_active,
        ownerToken=record.owner_token if include_owner else None,
    )


def _wps_conflict(error: Exception) -> HTTPException:
    if isinstance(error, ValueError):
        return service_error(422, "WPS_RECOVERY_INVALID", str(error), retryable=False)
    if isinstance(error, RuntimeError):
        return service_error(409, "WPS_RECOVERY_CONFLICT", str(error), retryable=True)
    return _task_state_error(error)


def _wps_record(operation_id: str) -> OfficeExecutionRecord:
    record = _folder_operation_ledger.office_execution(operation_id)
    if record is None or not record.workbook_identity.startswith("wps-path:"):
        raise RuntimeError("未找到这次 WPS 执行记录")
    # Old full-plan records cannot be resumed or disclosed via the new protocol.
    try:
        WpsRecoveryDescriptor.model_validate(record.plan)
        _wps_state_payload(record)
    except (ValueError, TypeError):
        # Never stringify validation errors containing a legacy plan/matrix.
        raise ValueError("旧执行记录不符合无矩阵恢复协议，不能下发或接管；请重新核对并预览。") from None
    return record


def _wps_evidence(operation_id: str, action_index: int, evidence: Any, *, after: bool = False) -> dict[str, Any]:
    record = _wps_record(operation_id)
    actions = record.plan["actions"]
    if action_index >= len(actions):
        raise ValueError("WPS 动作位置超出登记范围")
    action = actions[action_index]
    payload = evidence.model_dump(mode="json", exclude_none=True)
    if (payload["sheetName"], payload["range"], payload["type"], payload["beforeFingerprint"]) != (
        action["sheet"], action["range"], action["type"], action["previewFingerprint"]
    ):
        raise ValueError("WPS 动作证据与登记的范围或写前状态不一致")
    if after and (not payload.get("afterFingerprint") or (
        action.get("expectedFingerprint") and payload["afterFingerprint"] != action["expectedFingerprint"]
    )):
        raise ValueError("WPS 写后证据与登记的目标指纹不一致")
    return payload


def _planning_pause_payload(
    record: PlanningTaskRecord,
    *,
    interrupted: bool = False,
) -> dict[str, Any]:
    maximum_extensions = capability_int("taskRecovery", "planningMaxExtensions")
    hard_seconds = capability_float("taskRecovery", "planningHardWallTimeSeconds")
    extension_seconds = capability_float("taskRecovery", "planningExtensionSeconds")
    within_hard_limit = record.active_elapsed_seconds < hard_seconds
    model_error = bool(
        record.error
        and str(record.error.get("code", "")).startswith("MODEL_")
        and bool(record.error.get("retryable", False))
    )
    requires_extension = record.status == "waiting_user" and not model_error
    settings = model_settings_view()
    raw_connections = settings.get("connections", [])
    requires_vision = bool(record.request.get("images"))
    available_connections = [
        {
            "catalogModelId": str(connection.get("catalogModelId", "")),
            "label": str(connection.get("label", "")),
            "modelId": str(connection.get("modelId", "")),
            "supportsVision": connection.get("supportsVision") is True,
        }
        for connection in raw_connections
        if isinstance(connection, dict)
        and str(connection.get("catalogModelId", ""))
        and str(connection.get("catalogModelId", "")) != record.current_model_id
        and (not requires_vision or connection.get("supportsVision") is True)
    ]
    can_continue = (
        record.checkpoint is not None
        and within_hard_limit
        and (
            record.status == "running"
            or (
                record.status == "waiting_user"
                and (
                    model_error
                    or (
                        record.extension_count < maximum_extensions
                        and record.approved_wall_time_seconds < hard_seconds
                    )
                )
            )
        )
    )
    workbook = record.request.get("workbook")
    workbook_snapshot = workbook if isinstance(workbook, dict) else {}
    return {
        "kind": "planning_pause",
        "turnId": record.turn_id,
        "status": "interrupted" if interrupted else "waiting_user",
        "revision": record.revision,
        "lastSequence": record.last_sequence,
        "reason": record.pause_reason
        or (
            "服务重启后已找回已保存的规划现场。"
            if interrupted
            else "本轮处理时间已用完，当前进度已安全保存。"
        ),
        "activeElapsedSeconds": record.active_elapsed_seconds,
        "approvedWallTimeSeconds": record.approved_wall_time_seconds,
        "extensionSeconds": min(
            extension_seconds,
            max(hard_seconds - record.approved_wall_time_seconds, 0.0),
        ),
        "extensionCount": record.extension_count,
        "maximumExtensions": maximum_extensions,
        "canContinue": can_continue,
        "requiresExtension": requires_extension,
        "pauseType": "model_error" if model_error else ("interrupted" if interrupted else "budget"),
        "currentModelId": record.current_model_id,
        "modelSwitchCount": len(record.model_switches),
        "canSwitchModel": (model_error or interrupted)
        and record.checkpoint is not None
        and bool(available_connections),
        "availableModelConnections": available_connections,
        "sourceFingerprint": workbook_snapshot.get("sourceFingerprint"),
        "sourceFingerprintSheets": workbook_snapshot.get("sourceFingerprintSheets", []),
    }


async def _planning_task_stream(
    request: PlanRequest,
    record: PlanningTaskRecord,
    *,
    after_sequence: int = 0,
    activity_reserved: bool = False,
) -> AsyncIterator[str]:
    turn_id = record.turn_id
    state = turn_registry.touch(turn_id)
    task: asyncio.Task[AssistantResponse] | None = None
    async with state.completion_lock:
        if not activity_reserved:
            async with _active_planning_tasks_lock:
                if turn_id in _active_planning_tasks:
                    yield _sse_frame(
                        "error",
                        {
                            "status": 409,
                            "code": "PLANNING_TASK_STILL_RUNNING",
                            "message": "该规划任务仍在处理中，请勿重复继续。",
                            "retryable": True,
                        },
                    )
                    return
                _active_planning_tasks.add(turn_id)
        try:
            persisted_events = _folder_operation_ledger.planning_events(turn_id)
            for event in persisted_events:
                if int(event.get("sequence", 0) or 0) > after_sequence:
                    yield _sse_frame("step", event)

            queue: asyncio.Queue[dict[str, Any]] = asyncio.Queue()
            event_sequence = record.last_sequence
            current_record = record

            async def on_event(event: dict[str, Any]) -> None:
                nonlocal event_sequence, current_record
                event_sequence += 1
                payload = {
                    "schemaVersion": 1,
                    "operationId": turn_id,
                    "sequence": event_sequence,
                    "phase": event.get("phase", "planning"),
                    "stepId": event.get("stepId", f"planning-{event_sequence}"),
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
                serialized = validated.model_dump(mode="json")
                current_record = _folder_operation_ledger.append_planning_event(
                    turn_id, event_sequence, serialized
                )
                await queue.put(serialized)

            async def on_checkpoint(checkpoint: dict[str, Any]) -> None:
                nonlocal current_record
                loop_state = checkpoint.get("loopState")
                elapsed = (
                    float(loop_state.get("activeElapsedSeconds", 0.0) or 0.0)
                    if isinstance(loop_state, dict)
                    else current_record.active_elapsed_seconds
                )
                current_record = _folder_operation_ledger.save_planning_checkpoint(
                    turn_id,
                    checkpoint,
                    active_elapsed_seconds=elapsed,
                )

            task = asyncio.create_task(
                create_plan(
                    request,
                    tool_cache=state.tool_cache,
                    on_event=on_event,
                    on_checkpoint=on_checkpoint,
                    resume_checkpoint=current_record.checkpoint,
                    approved_wall_time_seconds=current_record.approved_wall_time_seconds,
                )
            )
            while True:
                drain = asyncio.ensure_future(queue.get())
                done, _pending = await asyncio.wait(
                    {drain, task}, return_when=asyncio.FIRST_COMPLETED
                )
                if drain in done:
                    yield _sse_frame("step", drain.result())
                    continue
                drain.cancel()
                break
            while not queue.empty():
                yield _sse_frame("step", queue.get_nowait())
            try:
                response = await task
            except AgentBudgetPause as pause:
                loop_state = pause.checkpoint.get("loopState")
                elapsed = (
                    float(loop_state.get("activeElapsedSeconds", 0.0) or 0.0)
                    if isinstance(loop_state, dict)
                    else current_record.active_elapsed_seconds
                )
                current_record = _folder_operation_ledger.pause_planning_task(
                    turn_id,
                    pause.checkpoint,
                    active_elapsed_seconds=elapsed,
                    reason="本轮处理时间已用完，当前进度已安全保存。",
                )
                yield _sse_frame("paused", _planning_pause_payload(current_record))
                return
            result = response.model_copy(update={"turnId": turn_id})
            current_record = _folder_operation_ledger.finish_planning_task(
                turn_id,
                status="failed" if isinstance(result, AnswerResponse) and result.outcome == "failed" else "completed",
                result=result.model_dump(mode="json"),
                error=({"status": 422, "code": "PLANNING_STOPPED", "message": result.message, "retryable": False}
                       if isinstance(result, AnswerResponse) and result.outcome == "failed" else None),
            )
            turn_registry.record_completion(turn_id, request, response)
            yield _sse_frame("result", current_record.result or result.model_dump(mode="json"))
        except (
            LLMTimeoutError,
            LLMConnectError,
            LLMHTTPStatusError,
            LLMTransportError,
            LLMResponseError,
            ModelNotConfiguredError,
            ModelNotAvailableError,
            TurnStateError,
            ValueError,
        ) as error:
            status, code, message, retryable = _map_turn_error(error)
            failure = {
                "status": status,
                "code": code,
                "message": message,
                "retryable": retryable,
            }
            try:
                current = _folder_operation_ledger.planning_task(turn_id)
                if current is not None and current.status == "running":
                    if current.checkpoint is not None and retryable and code.startswith("MODEL_"):
                        current_record = _folder_operation_ledger.pause_planning_task(
                            turn_id,
                            current.checkpoint,
                            active_elapsed_seconds=current.active_elapsed_seconds,
                            reason=(
                                f"模型连接中断：{message} 已保存完整规划现场；"
                                "可以重试当前连接，或明确选择一个已配置的备用模型。"
                            ),
                            error=failure,
                        )
                    else:
                        current_record = _folder_operation_ledger.finish_planning_task(
                            turn_id,
                            status="failed",
                            error=failure,
                        )
            except Exception:
                logger.exception("无法保存规划任务失败状态 turn_id=%s", turn_id)
            if current_record.status == "waiting_user":
                yield _sse_frame("paused", _planning_pause_payload(current_record))
                return
            yield _sse_frame(
                "error",
                failure,
            )
        except Exception:
            # SSE headers have already been sent, so an unexpected failure must be
            # represented in both the durable ledger and the event stream.  Without
            # this branch the client sees a broken connection while the task remains
            # incorrectly marked as running forever.
            logger.exception("规划任务发生未处理异常 turn_id=%s", turn_id)
            failure = {
                "status": 500,
                "code": "PLANNING_TASK_FAILED",
                "message": "规划任务未能安全保存当前状态，请重新发起任务。",
                "retryable": False,
            }
            try:
                current = _folder_operation_ledger.planning_task(turn_id)
                if current is not None and current.status == "running":
                    _folder_operation_ledger.finish_planning_task(
                        turn_id,
                        status="failed",
                        error=failure,
                    )
            except Exception:
                logger.exception("无法保存规划任务意外失败状态 turn_id=%s", turn_id)
            yield _sse_frame("error", failure)
        finally:
            # Closing an SSE response can finalize this async generator while the
            # model request is still running.  Do not release the active-task fence
            # until that child task has actually stopped; otherwise a reconnect can
            # resume the same checkpoint concurrently with the orphaned request.
            if task is not None and not task.done():
                task.cancel()
                try:
                    await task
                except asyncio.CancelledError:
                    pass
                except Exception:
                    logger.exception(
                        "断开规划流时等待后台任务停止失败 turn_id=%s", turn_id
                    )
            async with _active_planning_tasks_lock:
                _active_planning_tasks.discard(turn_id)


@app.post("/api/turn/stream")
async def turn_stream(request: TurnRequest) -> StreamingResponse:
    """Persisted planning stream with safe budget-pause boundaries."""
    if isinstance(request, IntentCheckRequest):
        raise service_error(
            422,
            "INVALID_REQUEST",
            "流式端点仅用于规划请求，意图判断请调用 /api/turn。",
            retryable=False,
        )
    turn_id = request.turnId or f"turn-{uuid.uuid4().hex}"
    normalized_request = request.model_copy(update={"turnId": turn_id})
    try:
        record, created = _folder_operation_ledger.register_planning_task(
            turn_id,
            _planning_request_fingerprint(normalized_request),
            normalized_request.model_dump(mode="json", exclude_none=True),
            approved_wall_time_seconds=capability_float(
                "taskRecovery", "planningReminderSeconds"
            ),
        )
    except Exception as error:
        raise _task_state_error(error) from error

    async def generate() -> AsyncIterator[str]:
        if not created:
            if record.status == "completed" and record.result is not None:
                yield _sse_frame("result", record.result)
                return
            if record.status == "waiting_user":
                yield _sse_frame("paused", _planning_pause_payload(record))
                return
            if record.status == "running":
                async with _active_planning_tasks_lock:
                    active = turn_id in _active_planning_tasks
                if active:
                    yield _sse_frame(
                        "error",
                        {
                            "status": 409,
                            "code": "PLANNING_TASK_STILL_RUNNING",
                            "message": "该规划任务仍在处理中。",
                            "retryable": True,
                        },
                    )
                else:
                    yield _sse_frame(
                        "paused", _planning_pause_payload(record, interrupted=True)
                    )
                return
            error = record.error or {}
            yield _sse_frame(
                "error",
                {
                    "status": error.get("status", 409),
                    "code": error.get("code", "PLANNING_TASK_STOPPED"),
                    "message": error.get("message", "规划任务已经停止。"),
                    "retryable": bool(error.get("retryable", False)),
                },
            )
            return
        async for item in _planning_task_stream(normalized_request, record):
            yield item

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/turn/{turn_id}/status")
async def planning_task_status(turn_id: str) -> dict[str, Any]:
    record = _folder_operation_ledger.planning_task(turn_id)
    if record is None:
        raise service_error(
            404, "PLANNING_TASK_NOT_FOUND", "未找到该规划任务。", retryable=False
        )
    async with _active_planning_tasks_lock:
        active = turn_id in _active_planning_tasks
    if record.status == "completed":
        return {
            "turnId": turn_id,
            "status": "completed",
            "revision": record.revision,
            "lastSequence": record.last_sequence,
            "result": record.result,
            "canContinue": False,
        }
    if record.status in {"running", "waiting_user"}:
        interrupted = record.status == "running" and not active
        return {
            **_planning_pause_payload(record, interrupted=interrupted),
            "status": "running" if active else ("interrupted" if interrupted else "waiting_user"),
            "canContinue": False if active else _planning_pause_payload(record, interrupted=interrupted)["canContinue"],
        }
    return {
        "turnId": turn_id,
        "status": record.status,
        "revision": record.revision,
        "lastSequence": record.last_sequence,
        "error": record.error,
        "canContinue": False,
    }


@app.post("/api/turn/continue/stream")
async def planning_task_continue_stream(
    request: PlanningContinueRequest,
) -> StreamingResponse:
    async with _active_planning_tasks_lock:
        if request.turnId in _active_planning_tasks:
            raise service_error(
                409,
                "PLANNING_TASK_STILL_RUNNING",
                "该规划任务仍在处理中，不能重复继续。",
                retryable=True,
            )
    existing = _folder_operation_ledger.planning_task(request.turnId)
    if existing is None or existing.checkpoint is None:
        raise service_error(
            409,
            "PLANNING_CHECKPOINT_NOT_AVAILABLE",
            "该任务没有可验证的完整规划现场，不能继续。",
            retryable=False,
        )
    extension_seconds = (
        capability_float("taskRecovery", "planningExtensionSeconds")
        if existing.status == "waiting_user"
        and not (
            existing.error
            and str(existing.error.get("code", "")).startswith("MODEL_")
            and bool(existing.error.get("retryable", False))
        )
        else 0.0
    )
    interrupted = existing.status == "running"
    selected_model_id = request.modelId
    if selected_model_id is not None:
        pause_payload = _planning_pause_payload(existing, interrupted=interrupted)
        if not pause_payload["canSwitchModel"]:
            raise service_error(
                409,
                "MODEL_SWITCH_NOT_ALLOWED",
                "只有因模型连接故障暂停，或执行器中断且现场完整的任务，才能切换备用模型。",
                retryable=False,
            )
        allowed_ids = {
            str(item.get("catalogModelId"))
            for item in pause_payload.get("availableModelConnections", [])
            if isinstance(item, dict)
        }
        if selected_model_id not in allowed_ids:
            raise service_error(
                409,
                "MODEL_SWITCH_NOT_ALLOWED",
                "所选备用模型未配置、能力不匹配或与当前模型相同。",
                retryable=False,
            )
        try:
            fallback_config = selected_model_config(selected_model_id)
        except ModelNotConfiguredError as error:
            raise service_error(
                409,
                "MODEL_NOT_CONFIGURED",
                str(error),
                retryable=False,
            ) from error
        except ModelNotAvailableError as error:
            raise service_error(
                409, "MODEL_SWITCH_NOT_ALLOWED", str(error), retryable=False
            ) from error
        if existing.request.get("images") and not fallback_config.supports_vision:
            raise service_error(
                409,
                "MODEL_SWITCH_CAPABILITY_MISMATCH",
                "所选备用模型不支持本任务所需的图片输入。",
                retryable=False,
            )
    try:
        async with _active_planning_tasks_lock:
            if request.turnId in _active_planning_tasks:
                raise service_error(
                    409,
                    "PLANNING_TASK_STILL_RUNNING",
                    "该规划任务仍在处理中，不能重复继续。",
                    retryable=True,
                )
            claimed, should_run = _folder_operation_ledger.claim_planning_continuation(
                request.turnId,
                expected_revision=request.expectedRevision,
                request_id=request.requestId,
                extension_seconds=extension_seconds,
                maximum_extensions=capability_int(
                    "taskRecovery", "planningMaxExtensions"
                ),
                hard_wall_time_seconds=capability_float(
                    "taskRecovery", "planningHardWallTimeSeconds"
                ),
                model_id=selected_model_id,
                model_switch_reason=(
                    (
                        "用户在执行器中断后明确选择备用模型接续断点"
                        if interrupted
                        else "用户在模型连接故障后明确选择备用模型"
                    )
                    if selected_model_id is not None
                    else None
                ),
            )
            if should_run:
                if selected_model_id is not None:
                    switch_event = TaskProgressEvent.model_validate(
                        {
                            "schemaVersion": 1,
                            "operationId": claimed.turn_id,
                            "sequence": claimed.last_sequence + 1,
                            "phase": "planning",
                            "stepId": f"planning.model.switch.{len(claimed.model_switches)}",
                            "status": "completed",
                            "title": "已切换到用户选择的备用模型",
                            "summary": "保留原任务范围、已完成工具结果和累计预算，从完整检查点继续。",
                            "facts": [
                                {
                                    "label": "备用连接",
                                    "value": selected_model_id,
                                    "style": "code",
                                }
                            ],
                            "timestamp": datetime.now().astimezone().isoformat(),
                        }
                    )
                    claimed = _folder_operation_ledger.append_planning_event(
                        claimed.turn_id,
                        claimed.last_sequence + 1,
                        switch_event.model_dump(mode="json"),
                    )
                plan_request = PlanRequest.model_validate(claimed.request).model_copy(
                    update={"modelId": claimed.current_model_id}
                )
                _active_planning_tasks.add(request.turnId)
    except RuntimeError as error:
        raise service_error(
            409, "PLANNING_CONTINUE_CONFLICT", str(error), retryable=True
        ) from error
    if not should_run:
        if claimed.status == "completed" and claimed.result is not None:
            async def completed() -> AsyncIterator[str]:
                yield _sse_frame("result", claimed.result or {})
            return StreamingResponse(completed(), media_type="text/event-stream")
        raise service_error(
            409,
            "PLANNING_TASK_STOPPED",
            "该规划任务已经停止，不能继续。",
            retryable=False,
        )
    return StreamingResponse(
        _planning_task_stream(
            plan_request,
            claimed,
            after_sequence=request.afterSequence,
            activity_reserved=True,
        ),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


def _office_conflict(error: Exception) -> HTTPException:
    if isinstance(error, ValueError):
        return service_error(422, "OFFICE_RECOVERY_INVALID", str(error), retryable=False)
    if isinstance(error, RuntimeError):
        return service_error(409, "OFFICE_RECOVERY_CONFLICT", str(error), retryable=True)
    return _task_state_error(error)


def _require_office_namespace(operation_id: str) -> None:
    record = _folder_operation_ledger.office_execution(operation_id)
    if record is None or not record.workbook_identity.startswith("url:"):
        raise RuntimeError("未找到这次 Office 执行记录，不能跨宿主接管")


@app.post(
    "/api/office/executions/register",
    response_model=OfficeExecutionStateResponse,
)
async def register_office_execution(
    request: OfficeExecutionRegisterRequest,
) -> OfficeExecutionStateResponse:
    _validate_office_recovery_plan(request)
    try:
        record, _created = await run_in_threadpool(
            _folder_operation_ledger.register_office_execution,
            request.operationId,
            _office_request_fingerprint(request),
            plan=request.plan.model_dump(mode="json"),
            authorization={
                "authorizedSheetNames": request.authorizedSheetNames,
                "allowedExternalFiles": request.allowedExternalFiles,
            },
            workbook_identity=request.workbookIdentity,
        )
    except Exception as error:
        raise _office_conflict(error) from error
    return _office_state_payload(record, include_owner=record.status == "running")


@app.get(
    "/api/office/executions/{operation_id}",
    response_model=OfficeExecutionStateResponse,
)
async def office_execution_status(operation_id: str) -> OfficeExecutionStateResponse:
    try:
        record = await run_in_threadpool(
            _folder_operation_ledger.office_execution, operation_id
        )
    except Exception as error:
        raise _task_state_error(error) from error
    if record is None or not record.workbook_identity.startswith("url:"):
        raise service_error(
            404,
            "OFFICE_EXECUTION_NOT_FOUND",
            "未找到这次 Office 执行记录；查询不会创建新的写入任务。",
            retryable=False,
        )
    return _office_state_payload(record)


@app.post(
    "/api/office/executions/{operation_id}/actions/start",
    response_model=OfficeExecutionStateResponse,
)
async def start_office_action(
    operation_id: str, request: OfficeActionStartRequest
) -> OfficeExecutionStateResponse:
    try:
        await run_in_threadpool(_require_office_namespace, operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.begin_office_action,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            action_index=request.actionIndex,
            before_snapshot=request.beforeSnapshot,
        )
    except Exception as error:
        raise _office_conflict(error) from error
    return _office_state_payload(record, include_owner=True)


@app.post(
    "/api/office/executions/{operation_id}/actions/complete",
    response_model=OfficeExecutionStateResponse,
)
async def complete_office_action(
    operation_id: str, request: OfficeActionCompleteRequest
) -> OfficeExecutionStateResponse:
    try:
        await run_in_threadpool(_require_office_namespace, operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.complete_office_action,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            action_index=request.actionIndex,
            checkpoint=request.checkpoint,
        )
    except Exception as error:
        raise _office_conflict(error) from error
    return _office_state_payload(record, include_owner=True)


@app.post(
    "/api/office/executions/{operation_id}/suspend",
    response_model=OfficeExecutionStateResponse,
)
async def suspend_office_execution(
    operation_id: str, request: OfficeExecutionSuspendRequest
) -> OfficeExecutionStateResponse:
    try:
        await run_in_threadpool(_require_office_namespace, operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.suspend_office_execution,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            error={"code": request.code, "message": request.message},
        )
    except Exception as error:
        raise _office_conflict(error) from error
    return _office_state_payload(record)


@app.post(
    "/api/office/executions/{operation_id}/resume",
    response_model=OfficeExecutionStateResponse,
)
async def resume_office_execution(
    operation_id: str, request: OfficeExecutionResumeRequest
) -> OfficeExecutionStateResponse:
    try:
        await run_in_threadpool(_require_office_namespace, operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.claim_office_resume,
            operation_id,
            expected_revision=request.expectedRevision,
        )
    except Exception as error:
        raise _office_conflict(error) from error
    return _office_state_payload(record, include_owner=True)


@app.post(
    "/api/office/executions/{operation_id}/actions/reconcile",
    response_model=OfficeExecutionStateResponse,
)
async def reconcile_office_action(
    operation_id: str, request: OfficeActionReconcileRequest
) -> OfficeExecutionStateResponse:
    try:
        await run_in_threadpool(_require_office_namespace, operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.reconcile_office_action,
            operation_id,
            expected_revision=request.expectedRevision,
            action_index=request.actionIndex,
            outcome=request.outcome,
            checkpoint=request.checkpoint,
        )
    except Exception as error:
        raise _office_conflict(error) from error
    return _office_state_payload(record, include_owner=True)


@app.post(
    "/api/office/executions/{operation_id}/complete",
    response_model=OfficeExecutionStateResponse,
)
async def finish_office_execution(
    operation_id: str, request: OfficeExecutionFinishRequest
) -> OfficeExecutionStateResponse:
    try:
        await run_in_threadpool(_require_office_namespace, operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.finish_office_execution,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            result=request.result,
        )
    except Exception as error:
        raise _office_conflict(error) from error
    return _office_state_payload(record)


@app.post(
    "/api/wps/executions/register",
    response_model=WpsExecutionStateResponse,
)
async def register_wps_execution(
    request: WpsExecutionRegisterRequest,
) -> WpsExecutionStateResponse:
    _validate_wps_recovery_descriptor(request)
    try:
        record, _created = await run_in_threadpool(
            _folder_operation_ledger.register_office_execution,
            request.operationId,
            _wps_request_fingerprint(request),
            plan=request.descriptor.model_dump(mode="json"),
            authorization={
                "authorizedSheetNames": request.authorizedSheetNames,
                "allowedExternalFiles": [],
            },
            workbook_identity=request.workbookIdentity,
            identity_prefix="wps-path:",
            task_id=request.taskId,
        )
    except Exception as error:
        raise _wps_conflict(error) from error
    return _wps_state_payload(record, include_owner=_created and record.status == "running")


@app.get(
    "/api/wps/executions/{operation_id}",
    response_model=WpsExecutionStateResponse,
)
async def wps_execution_status(operation_id: str) -> WpsExecutionStateResponse:
    try:
        record = await run_in_threadpool(
            _folder_operation_ledger.office_execution, operation_id
        )
    except Exception as error:
        raise _task_state_error(error) from error
    if record is None or not record.workbook_identity.startswith("wps-path:"):
        raise service_error(
            404,
            "WPS_EXECUTION_NOT_FOUND",
            "未找到这次 WPS 执行记录；查询不会创建新的写入任务。",
            retryable=False,
        )
    try:
        return _wps_state_payload(await run_in_threadpool(_wps_record, operation_id))
    except Exception as error:
        raise _wps_conflict(error) from error


@app.post(
    "/api/wps/executions/{operation_id}/actions/start",
    response_model=WpsExecutionStateResponse,
)
async def start_wps_action(
    operation_id: str, request: WpsActionStartRequest
) -> WpsExecutionStateResponse:
    try:
        evidence = _wps_evidence(operation_id, request.actionIndex, request.beforeSnapshot)
        record = await run_in_threadpool(
            _folder_operation_ledger.begin_office_action,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            action_index=request.actionIndex,
            before_snapshot=evidence,
        )
    except Exception as error:
        raise _wps_conflict(error) from error
    return _wps_state_payload(record, include_owner=True)


@app.post(
    "/api/wps/executions/{operation_id}/actions/complete",
    response_model=WpsExecutionStateResponse,
)
async def complete_wps_action(
    operation_id: str, request: WpsActionCompleteRequest
) -> WpsExecutionStateResponse:
    try:
        evidence = _wps_evidence(operation_id, request.actionIndex, request.checkpoint, after=True)
        record = await run_in_threadpool(
            _folder_operation_ledger.complete_office_action,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            action_index=request.actionIndex,
            checkpoint=evidence,
        )
    except Exception as error:
        raise _wps_conflict(error) from error
    return _wps_state_payload(record, include_owner=True)


@app.post(
    "/api/wps/executions/{operation_id}/suspend",
    response_model=WpsExecutionStateResponse,
)
async def suspend_wps_execution(
    operation_id: str, request: WpsExecutionSuspendRequest
) -> WpsExecutionStateResponse:
    try:
        _wps_record(operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.suspend_office_execution,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            error={"code": "WPS_WRITE_INTERRUPTED", "message": "WPS 写入中断，需要核对现场。"},
        )
    except Exception as error:
        raise _wps_conflict(error) from error
    return _wps_state_payload(record)


@app.post(
    "/api/wps/executions/{operation_id}/resume",
    response_model=WpsExecutionStateResponse,
)
async def resume_wps_execution(
    operation_id: str, request: WpsExecutionResumeRequest
) -> WpsExecutionStateResponse:
    try:
        await run_in_threadpool(_wps_record, operation_id)
        record = await run_in_threadpool(
            _folder_operation_ledger.claim_office_resume,
            operation_id,
            expected_revision=request.expectedRevision,
            require_idle=True,
        )
    except Exception as error:
        raise _wps_conflict(error) from error
    return _wps_state_payload(record, include_owner=True)


@app.post(
    "/api/wps/executions/{operation_id}/actions/reconcile",
    response_model=WpsExecutionStateResponse,
)
async def reconcile_wps_action(
    operation_id: str, request: WpsActionReconcileRequest
) -> WpsExecutionStateResponse:
    try:
        _wps_record(operation_id)
        evidence = _wps_evidence(operation_id, request.actionIndex, request.checkpoint, after=True) if request.checkpoint else None
        record = await run_in_threadpool(
            _folder_operation_ledger.reconcile_office_action,
            operation_id,
            expected_revision=request.expectedRevision,
            action_index=request.actionIndex,
            outcome=request.outcome,
            checkpoint=evidence,
            require_idle=True,
        )
    except Exception as error:
        raise _wps_conflict(error) from error
    return _wps_state_payload(record, include_owner=True)


@app.post(
    "/api/wps/executions/{operation_id}/complete",
    response_model=WpsExecutionStateResponse,
)
async def finish_wps_execution(
    operation_id: str, request: WpsExecutionFinishRequest
) -> WpsExecutionStateResponse:
    try:
        previous = _wps_record(operation_id)
        if request.result.actionCount != len(previous.plan["actions"]) or (
            request.result.cellCount is not None and request.result.cellCount != sum(
                action["cellCount"] for action in previous.plan["actions"]
            )
        ):
            raise ValueError("完成回执的动作或单元格计数与登记不一致")
        record = await run_in_threadpool(
            _folder_operation_ledger.finish_office_execution,
            operation_id,
            expected_revision=request.expectedRevision,
            owner_token=request.ownerToken,
            result=request.result.model_dump(mode="json", exclude_none=True),
        )
    except Exception as error:
        raise _wps_conflict(error) from error
    return _wps_state_payload(record)


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
                checkpoint_supported = (
                    len(request.plan.actions) == 1
                    and request.plan.actions[0].type == "splitGroupAggregate"
                )
                if checkpoint_supported:
                    try:
                        authorization = folder_execution_context(request)
                        record = _folder_operation_ledger.save_execution_context(
                            operation_id,
                            request=request.model_dump(
                                mode="json", exclude={"afterSequence"}, exclude_none=True
                            ),
                            authorization=authorization,
                        )
                    except Exception as error:
                        raise _task_state_error(error) from error
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
        try:
            persisted = _folder_operation_ledger.get(operationId)
            checkpoint = _folder_operation_ledger.checkpoint(operationId)
        except Exception as error:
            raise _task_state_error(error) from error
        checkpoint_stage = (
            str(checkpoint.payload.get("stage")) if checkpoint is not None else None
        )
        resume_summary: dict[str, object] = {
            "canResume": False,
            "reason": None,
        }
        resume_candidate = (
            state.task is None
            and state.status in {"running", "needs_reconciliation"}
            and checkpoint_stage in {"publish_prepared", "publishing", "published"}
        )
        if resume_candidate and persisted is not None and checkpoint is not None:
            if persisted.authorization is None:
                resume_summary["reason"] = "任务缺少原授权现场"
            else:
                try:
                    resume_summary = folder_resume_checkpoint_summary(
                        operationId,
                        persisted.authorization,
                        checkpoint.payload,
                    )
                except Exception as error:
                    logger.warning(
                        "核对文件夹任务续跑条件失败 operation_id=%s: %s",
                        operationId,
                        error,
                    )
                    resume_summary = {
                        "canResume": False,
                        "reason": "恢复资料当前无法安全核对",
                    }
        return {
            "operationId": operationId,
            "status": state.status,
            "lastSequence": state.sequence,
            "result": state.result,
            "error": state.error,
            "revision": persisted.revision if persisted is not None else None,
            "canResume": bool(resume_summary.get("canResume")),
            "checkpointStage": checkpoint_stage,
            "resumeReason": resume_summary.get("reason"),
            "resumeFilesApplied": resume_summary.get("filesApplied"),
            "resumeFilesPending": resume_summary.get("filesPending"),
        }


@app.post("/api/folders/execute/resume/stream")
async def folder_execute_resume_stream(
    request: FolderExecutionResumeRequest,
) -> StreamingResponse:
    """Explicitly continue a durable folder publish from its last safe boundary."""

    operation_id = request.operationId
    async with _folder_operations_lock:
        active = _folder_operations.get(operation_id)
        if active is not None and active.task is not None and not active.task.done():
            raise service_error(
                409,
                "OPERATION_STILL_RUNNING",
                "该任务仍在当前进程中执行，不能同时发起续跑。",
                retryable=True,
            )
        try:
            existing = _folder_operation_ledger.get(operation_id)
            checkpoint = _folder_operation_ledger.checkpoint(operation_id)
        except Exception as error:
            raise _task_state_error(error) from error
        if existing is None:
            raise service_error(
                404, "OPERATION_NOT_FOUND", "未找到可续跑的任务。", retryable=False
            )
        if existing.status == "completed":
            state = _folder_state_from_record(existing)
            _folder_operations[operation_id] = state
        else:
            if existing.revision != request.expectedRevision:
                raise service_error(
                    409,
                    "RESUME_CLAIM_CONFLICT",
                    "任务状态已变化，请刷新后再决定是否继续。",
                    retryable=True,
                )
            if checkpoint is None:
                raise service_error(
                    409,
                    "CHECKPOINT_NOT_AVAILABLE",
                    "该任务没有留下可验证的安全断点，不能续跑。",
                    retryable=False,
                )
            if existing.authorization is None:
                raise service_error(
                    409,
                    "RESUME_CONTEXT_NOT_AVAILABLE",
                    "该任务缺少原授权现场，不能安全续跑。",
                    retryable=False,
                )
            try:
                resume_summary = folder_resume_checkpoint_summary(
                    operation_id,
                    existing.authorization,
                    checkpoint.payload,
                )
            except Exception as error:
                raise service_error(
                    409,
                    "RESUME_RECONCILIATION_FAILED",
                    f"恢复资料当前无法安全核对：{error}",
                    retryable=False,
                ) from error
            if not bool(resume_summary.get("canResume")):
                raise service_error(
                    409,
                    "RESUME_NOT_SAFE",
                    str(resume_summary.get("reason") or "当前文件状态不允许继续任务。"),
                    retryable=False,
                )
            try:
                claimed = _folder_operation_ledger.claim_for_resume(
                    operation_id, expected_revision=request.expectedRevision
                )
            except RuntimeError as error:
                raise service_error(
                    409, "RESUME_CLAIM_CONFLICT", str(error), retryable=True
                ) from error
            state = _folder_state_from_record(claimed)
            _folder_operations[operation_id] = state
            state.task = asyncio.create_task(
                _run_folder_resume(operation_id, state, claimed, checkpoint.payload)
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
    except (ModelNotConfiguredError, ModelNotAvailableError) as error:
        # 与 /api/turn、/api/turn/stream 走同一套映射：
        # 「没有可用模型」必须是可识别的稳定错误码，而不是泛化的 INVALID_REQUEST。
        # retryable 取映射结果（当前为 False —— 需要用户先去配置模型，
        # 原样重试不会成功；文案里已给出配置路径）。
        status, code, message, retryable = _map_turn_error(error)
        raise service_error(
            status, code, message, retryable=retryable
        ) from error
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

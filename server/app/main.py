from __future__ import annotations

import asyncio
import json
import uuid
import time
from collections import deque
from datetime import datetime
from typing import Any, AsyncIterator

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

from .capabilities import capabilities
from .intent import check_intent
from .folder_workbooks import (
    FolderCatalog,
    FolderExecuteRequest,
    FolderExecuteResponse,
    FolderQueryRequest,
    FolderSnapshotRequest,
    create_folder_snapshot,
    execute_folder_plan,
    query_folder_data,
    select_and_scan_folder,
)
from .models import (
    AssistantResponse,
    DataToolResult,
    IntentCheckRequest,
    IntentCheckResponse,
    ModelSettingsResponse,
    PlanRequest,
    TestModelConnectionResponse,
    SetFormulaModelRequest,
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
from .planner import create_plan
from .rule_generator import (
    GenerateFormulaRequest,
    GenerateFormulaResponse,
    generate_formula,
)
from .turn_state import TurnStateError, turn_registry

load_dotenv("server/.env")

app = FastAPI(
    title="Excel Bro Local Service",
    version="0.1.0",
    description="Local planning and analysis service for the Excel Bro add-in.",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://localhost:3000",
        "https://127.0.0.1:3000",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
    ],
    allow_credentials=False,
    allow_methods=["DELETE", "GET", "POST", "PUT"],
    allow_headers=["Content-Type"],
)

_diagnostic_events: deque[dict[str, object]] = deque(maxlen=500)


@app.middleware("http")
async def record_request_metrics(request, call_next):
    started = time.perf_counter()
    response = await call_next(request)
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
        body = error.body
        return (
            502,
            "MODEL_HTTP_ERROR",
            f"模型服务返回 HTTP {status}{f'：{body}' if body else ''}",
            False,
        )
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

                async def on_event(event: dict[str, Any]) -> None:
                    await queue.put(event)

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
    return await run_in_threadpool(select_and_scan_folder)


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
    try:
        return await run_in_threadpool(execute_folder_plan, request)
    except (OSError, ValueError) as error:
        raise service_error(
            422, "FOLDER_EXECUTION_ERROR", str(error), retryable=True
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
        body = error.body
        raise HTTPException(
            status_code=502,
            detail=f"模型服务返回 HTTP {status}{f'：{body}' if body else ''}",
        ) from error
    except (LLMTransportError, LLMResponseError) as error:
        raise HTTPException(
            status_code=502,
            detail=f"模型意图识别连接失败：{error}",
        ) from error
    except ValueError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error

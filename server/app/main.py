from __future__ import annotations

import uuid

from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.concurrency import run_in_threadpool
from fastapi.middleware.cors import CORSMiddleware

from .capabilities import capabilities
from .intent import check_intent
from .folder_workbooks import (
    FolderCatalog,
    FolderExecuteRequest,
    FolderExecuteResponse,
    FolderSnapshotRequest,
    create_folder_snapshot,
    execute_folder_plan,
    select_and_scan_folder,
)
from .models import (
    AssistantResponse,
    IntentCheckRequest,
    IntentCheckResponse,
    PlanRequest,
    TurnRequest,
    TurnResponse,
    WorkbookSnapshot,
)
from .llm import (
    LLMConnectError,
    LLMHTTPStatusError,
    LLMResponseError,
    LLMTimeoutError,
    LLMTransportError,
    model_catalog,
    model_status,
    selected_model_config,
)
from .planner import create_plan
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
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)


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


@app.get("/health")
async def health() -> dict[str, str | bool | None]:
    return {
        "status": "ok",
        **model_status(),
    }


@app.get("/api/models")
async def models() -> dict[str, object]:
    return model_catalog()


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
                response = await create_plan(request)
                turn_registry.record_completion(turn_id, request, response)
        return response.model_copy(update={"turnId": turn_id})
    except LLMTimeoutError as error:
        raise service_error(
            504,
            "MODEL_TIMEOUT",
            "模型响应超时，可以重试当前对话轮次。",
            retryable=True,
        ) from error
    except LLMConnectError as error:
        raise service_error(
            502,
            "MODEL_UNREACHABLE",
            "无法连接模型服务，请检查网络和 AI_BASE_URL。",
            retryable=True,
        ) from error
    except LLMHTTPStatusError as error:
        status = error.status_code
        body = error.body
        raise service_error(
            502,
            "MODEL_HTTP_ERROR",
            f"模型服务返回 HTTP {status}{f'：{body}' if body else ''}",
            retryable=status == 429 or status >= 500,
        ) from error
    except (LLMTransportError, LLMResponseError) as error:
        raise service_error(
            502,
            "MODEL_CONNECTION_ERROR",
            str(error),
            retryable=True,
        ) from error
    except TurnStateError as error:
        raise service_error(
            409,
            error.code,
            str(error),
            retryable=False,
        ) from error
    except ValueError as error:
        raise service_error(
            422,
            "INVALID_REQUEST",
            str(error),
            retryable=False,
        ) from error


@app.post("/api/folders/select", response_model=FolderCatalog | None)
async def select_folder() -> FolderCatalog | None:
    return await run_in_threadpool(select_and_scan_folder)


@app.post("/api/folders/snapshot", response_model=WorkbookSnapshot)
async def folder_snapshot(request: FolderSnapshotRequest) -> WorkbookSnapshot:
    try:
        snapshot = await run_in_threadpool(create_folder_snapshot, request)
        return snapshot
    except ValueError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error


@app.post("/api/folders/execute", response_model=FolderExecuteResponse)
async def folder_execute(request: FolderExecuteRequest) -> FolderExecuteResponse:
    try:
        return await run_in_threadpool(execute_folder_plan, request)
    except (OSError, ValueError) as error:
        raise HTTPException(status_code=422, detail=str(error)) from error


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
        raise HTTPException(status_code=422, detail=str(error)) from error


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

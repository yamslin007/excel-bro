from fastapi import APIRouter, HTTPException, Request
from fastapi.concurrency import run_in_threadpool
from .models import (WorksheetCheckpointBackupConsent, WorksheetCheckpointListRequest,
                     WorksheetCheckpointListResult, WorksheetCheckpointOpenResult,
                     WorksheetCheckpointReference, WorksheetDeleteCheckpointRequest,
                     WorksheetDeleteCheckpointResult, WorksheetRestoreResult)
from .worksheet_checkpoint import CheckpointPersistenceError, WorksheetCheckpointService


def create_worksheet_checkpoint_router(allowed_origins, service=None):
    service = service or WorksheetCheckpointService()
    router = APIRouter(prefix="/api/worksheet-checkpoints")

    def guard(request):
        origin = request.headers.get("origin", "")
        if (not request.client or request.client.host not in {"127.0.0.1", "::1"}
            or request.url.hostname not in {"127.0.0.1", "localhost", "::1"}
            or origin not in allowed_origins
            or request.headers.get("content-type", "").split(";", 1)[0].strip().lower() != "application/json"):
            raise HTTPException(403, detail="仅允许本机加载项访问工作表检查点。")
        return origin

    async def call(method, *args):
        try:
            return await run_in_threadpool(method, *args)
        except CheckpointPersistenceError as exc:
            raise HTTPException(409, detail=str(exc)) from None
        except ValueError as exc:
            raise HTTPException(409, detail=str(exc)) from None
        except Exception:
            raise HTTPException(409, detail="检查点操作尚未确认；请查询状态，不要重复执行删除。") from None

    @router.post("/execute", response_model=WorksheetDeleteCheckpointResult)
    async def execute(payload: WorksheetDeleteCheckpointRequest, request: Request):
        return await call(service.execute, payload, guard(request))

    @router.post("/confirm-backup", response_model=WorksheetDeleteCheckpointResult)
    async def confirm_backup(payload: WorksheetCheckpointBackupConsent, request: Request):
        return await call(service.confirm_backup, payload, guard(request))

    @router.post("/list", response_model=WorksheetCheckpointListResult)
    async def list_for_workbook(payload: WorksheetCheckpointListRequest, request: Request):
        return await call(service.list_for_workbook, payload, guard(request))

    @router.post("/status", response_model=WorksheetDeleteCheckpointResult)
    async def status(payload: WorksheetCheckpointReference, request: Request):
        return await call(service.status, payload.operationId, guard(request))

    @router.post("/open", response_model=WorksheetCheckpointOpenResult)
    async def open_copy(payload: WorksheetCheckpointReference, request: Request):
        return await call(service.open, payload.operationId, guard(request))

    @router.post("/restore", response_model=WorksheetRestoreResult)
    async def restore_in_place(payload: WorksheetCheckpointReference, request: Request):
        return await call(service.restore, payload.operationId, guard(request))

    return router

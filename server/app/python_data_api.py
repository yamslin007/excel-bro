"""Dedicated local-only data channel; never reuses planning payloads."""
import json
from urllib.parse import urlsplit

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse

from .capabilities import capability_int
from .models import PythonDataScopeRequest, PythonDataSnapshotRequest, PythonDataRevokeRequest, PythonResultReadRequest
from .python_analysis import python_execution_enabled
from .python_datasets import registry


def create_python_data_router(allowed_origins: list[str]) -> APIRouter:
    router = APIRouter(prefix="/api/python/data")

    async def read(request: Request, model):
        origin = request.headers.get("origin", "")
        if (not request.client or request.client.host not in {"127.0.0.1", "::1"}
            or request.url.hostname not in {"localhost", "127.0.0.1", "::1"}
            or origin not in allowed_origins
            or request.headers.get("content-type", "").split(";", 1)[0].strip().lower() != "application/json"):
            raise HTTPException(403, detail="仅允许本机加载项传递授权数据。")
        if not python_execution_enabled():
            raise HTTPException(409, detail="隔离 Python 未启用，未读取数据。")
        maximum = capability_int("pythonData", "maxRequestBytes")
        body = bytearray()
        try:
            async for chunk in request.stream():
                if len(body) + len(chunk) > maximum:
                    raise ValueError()
                body.extend(chunk)
            def invalid_constant(_):
                raise ValueError()
            payload = model.model_validate(json.loads(body, parse_constant=invalid_constant))
            expected_host = "excel-office" if urlsplit(origin).port == 3000 else "wps-et"
            if payload.scope.hostId != expected_host:
                raise ValueError()
            return payload, origin
        except Exception:
            # Never FastAPI's default 422 input echo, exception repr or raw rows.
            raise HTTPException(409, detail="授权数据请求未通过校验。") from None

    def response(value):
        return JSONResponse(value, headers={"Cache-Control": "no-store"})

    @router.post("/scope")
    async def scope(request: Request):
        payload, origin = await read(request, PythonDataScopeRequest)
        try:
            return response({"scopeToken": registry.create(payload.scope, origin)})
        except Exception:
            raise HTTPException(409, detail="无法登记当前授权范围。") from None

    @router.post("/snapshot")
    async def snapshot(request: Request):
        payload, origin = await read(request, PythonDataSnapshotRequest)
        try:
            receipt = registry.register(payload.scopeToken, payload.scope, origin,
                                        payload.sourceIndex, payload.values)
            return response(receipt.model_dump())
        except Exception:
            raise HTTPException(409, detail="授权数据已失效或超过限额，请重新确认。") from None

    @router.post("/revoke")
    async def revoke(request: Request):
        payload, origin = await read(request, PythonDataRevokeRequest)
        try:
            registry.revoke(payload.scopeToken, payload.scope, origin)
            return response({"revoked": True})
        except Exception:
            raise HTTPException(409, detail="授权数据已失效。") from None

    @router.post("/result")
    async def result(request: Request):
        payload, origin = await read(request, PythonResultReadRequest)
        try:
            return response(registry.read_result(payload.scopeToken, payload.scope,
                origin, payload.datasets, payload.result))
        except Exception:
            raise HTTPException(409, detail="分析结果已失效，请重新读取并分析。") from None

    @router.post("/result/discard")
    async def discard_result(request: Request):
        payload, origin = await read(request, PythonResultReadRequest)
        try:
            registry.discard_result(payload.scopeToken, payload.scope, origin, payload.datasets, payload.result)
            return response({"discarded": True})
        except Exception:
            raise HTTPException(409, detail="分析结果已失效。") from None

    return router

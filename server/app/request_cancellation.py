"""Cancellation for request-bound, read-only model calls before a task exists."""
from __future__ import annotations

import asyncio
from typing import Awaitable, TypeVar

from fastapi import HTTPException, Request

T = TypeVar("T")


async def until_disconnected(work: Awaitable[T], request: Request | None) -> T:
    # Direct internal callers have no HTTP lifetime. Never apply this to writes
    # or persistent planners, whose own ledger/cleanup owns cancellation.
    if request is None:
        return await work

    async def disconnected() -> None:
        # FastAPI has consumed the parsed request body before entering the route.
        while (await request.receive())["type"] != "http.disconnect":
            pass

    task = asyncio.ensure_future(work)
    watcher = asyncio.create_task(disconnected())
    try:
        done, _ = await asyncio.wait({task, watcher}, return_when=asyncio.FIRST_COMPLETED)
        if watcher in done:
            # Also observe receive errors; never publish a decision after disconnect.
            watcher.result()
            raise HTTPException(status_code=499, detail={
                "code": "REQUEST_CANCELLED", "message": "请求已暂停。", "retryable": False,
            })
        return task.result()
    finally:
        for pending in (task, watcher):
            if not pending.done():
                pending.cancel()
        await asyncio.gather(task, watcher, return_exceptions=True)

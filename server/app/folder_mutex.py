from __future__ import annotations

import hashlib
import os
import threading
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from .app_data import task_lock_root


class FolderWriteBusyError(RuntimeError):
    pass


_registry_guard = threading.Lock()
_thread_locks: dict[str, threading.Lock] = {}


def _identity(root: Path) -> str:
    normalized = os.path.normcase(str(root.resolve()))
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def _thread_lock(key: str) -> threading.Lock:
    with _registry_guard:
        return _thread_locks.setdefault(key, threading.Lock())


def _lock_file(file_handle) -> None:
    file_handle.seek(0)
    if os.name == "nt":
        import msvcrt

        try:
            msvcrt.locking(file_handle.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as error:
            raise FolderWriteBusyError(
                "这个工作目录正在被另一项 Excel Bro 写入或恢复任务使用，请稍后重试"
            ) from error
        return
    import fcntl

    try:
        fcntl.flock(file_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError as error:
        raise FolderWriteBusyError(
            "这个工作目录正在被另一项 Excel Bro 写入或恢复任务使用，请稍后重试"
        ) from error


def _unlock_file(file_handle) -> None:
    file_handle.seek(0)
    if os.name == "nt":
        import msvcrt

        msvcrt.locking(file_handle.fileno(), msvcrt.LK_UNLCK, 1)
        return
    import fcntl

    fcntl.flock(file_handle.fileno(), fcntl.LOCK_UN)


@contextmanager
def folder_write_guard(root: Path) -> Iterator[None]:
    """Serialize writes for one registered work directory."""

    key = _identity(root)
    thread_lock = _thread_lock(key)
    if not thread_lock.acquire(blocking=False):
        raise FolderWriteBusyError(
            "这个工作目录正在被另一项 Excel Bro 写入或恢复任务使用，请稍后重试"
        )
    lock_root = task_lock_root()
    lock_root.mkdir(parents=True, exist_ok=True)
    lock_path = lock_root / f"folder-{key}.lock"
    file_handle = None
    locked = False
    try:
        file_handle = lock_path.open("a+b")
        file_handle.seek(0, os.SEEK_END)
        if file_handle.tell() == 0:
            file_handle.write(b"0")
            file_handle.flush()
        _lock_file(file_handle)
        locked = True
        yield
    finally:
        if locked and file_handle is not None:
            _unlock_file(file_handle)
        if file_handle is not None:
            file_handle.close()
        thread_lock.release()

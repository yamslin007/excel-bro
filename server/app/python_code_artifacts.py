"""模型生成的 Python 代码工件的**短期内存**登记（不落盘、不出站）。

边界（与仓库既有纪律一致，不得放宽）：

* 代码只在任务受控内存里短期存在。``planning_task_events`` 之类的账本一律只留
  ``codeId`` / 阶段 / 版本 / 字符数，**没有正文**。已有测试
  （``test_python_analysis.py::test_unverified_handles_never_enter_sqlite`` 等）
  把这条钉成硬判据：正文一旦进入 ``on_event`` 就会落进 SQLite。
* 取回只走本机回环端点，按 ``codeId`` 精确命中，且必须同时给出所属任务 ——
  只有 ``codeId`` 拿得到的话，任何能读到一个 id 的进程都能读到别的任务的代码。
* 容量与存活时间都有上限，过期即弃（这与"不保存代码"是同一条边界的两种表达）。
"""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass

# 最多同时保留多少段代码、每段活多久。
_MAX_ENTRIES = 24
_TTL_SECONDS = 30 * 60


@dataclass(frozen=True)
class PythonCodeArtifact:
    task_id: str
    code_id: str
    attempt_id: str
    code: str
    created_at: float


class PythonCodeRegistry:
    """按 ``codeId`` 存取代码正文的进程内登记表。"""

    def __init__(self, *, max_entries: int = _MAX_ENTRIES, ttl_seconds: float = _TTL_SECONDS) -> None:
        self._lock = threading.Lock()
        self._entries: dict[str, PythonCodeArtifact] = {}
        self._max_entries = max_entries
        self._ttl = ttl_seconds

    def _purge(self, now: float) -> None:
        expired = [
            key for key, item in self._entries.items() if now - item.created_at > self._ttl
        ]
        for key in expired:
            self._entries.pop(key, None)
        # 超出容量时按最旧优先淘汰：代码是"临时工件"，不是账本。
        while len(self._entries) > self._max_entries:
            oldest = min(self._entries.items(), key=lambda kv: kv[1].created_at)[0]
            self._entries.pop(oldest, None)

    def register(self, *, task_id: str, code_id: str, attempt_id: str, code: str) -> None:
        now = time.time()
        with self._lock:
            self._entries[code_id] = PythonCodeArtifact(
                task_id=task_id,
                code_id=code_id,
                attempt_id=attempt_id,
                code=code,
                created_at=now,
            )
            self._purge(now)

    def get(self, code_id: str, *, task_id: str | None = None) -> PythonCodeArtifact | None:
        """按 ``codeId`` 取回；给了 ``task_id`` 就必须与登记时的一致。

        **任务不一致按"不存在"处理**：不区分"没有这段代码"与"这段代码不属于你"，
        否则端点会变成一个跨任务探测工具。
        """
        now = time.time()
        with self._lock:
            self._purge(now)
            item = self._entries.get(code_id)
        if item is None:
            return None
        if task_id is not None and item.task_id != task_id:
            return None
        return item

    def clear(self) -> None:
        with self._lock:
            self._entries.clear()


python_code_artifacts = PythonCodeRegistry()

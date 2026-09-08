from __future__ import annotations

import os
import sys
from pathlib import Path


def application_data_root() -> Path:
    """Return the single local data root used by dev and packaged runtimes."""

    configured = os.getenv("EXCEL_BRO_CONFIG_DIR", "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "Excel Bro"
    if os.name == "nt":
        local_app_data = os.getenv("LOCALAPPDATA", "").strip()
        if not local_app_data:
            raise RuntimeError("Windows 环境缺少 LOCALAPPDATA，无法定位 Excel Bro 数据目录")
        return Path(local_app_data) / "Excel Bro"
    xdg_data_home = os.getenv("XDG_DATA_HOME", "").strip()
    base = (
        Path(xdg_data_home).expanduser()
        if xdg_data_home
        else Path.home() / ".local" / "share"
    )
    return base / "excel-bro"


def task_state_path() -> Path:
    return application_data_root() / "tasks" / "state.sqlite3"


def task_lock_root() -> Path:
    return application_data_root() / "tasks" / "locks"

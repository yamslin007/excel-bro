from __future__ import annotations

import os
import json
from dataclasses import dataclass
from pathlib import Path
import uuid

from dotenv import set_key


class ModelConnectionStoreError(RuntimeError):
    pass


class ModelConnectionNotFoundError(ValueError):
    pass


class DuplicateModelConnectionError(ValueError):
    pass


def _first_env(*names: str) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    return ""


def _csv_env(name: str) -> list[str]:
    return [
        value.strip()
        for value in os.getenv(name, "").split(",")
        if value.strip()
    ]


@dataclass(frozen=True, slots=True)
class ModelConnection:
    """A selected model and the credentials needed by its adapter."""

    base_url: str
    model: str
    api_key: str
    supports_vision: bool
    adapter: str = "openai_compatible"


@dataclass(frozen=True, slots=True)
class ModelSettings:
    base_url: str
    default_model: str
    api_key: str
    models: tuple[str, ...]
    vision_models: frozenset[str]

    def supports_vision(self, model: str) -> bool:
        normalized = model.casefold()
        return "*" in self.vision_models or normalized in self.vision_models


@dataclass(frozen=True, slots=True)
class ManagedModelConnection:
    id: str
    label: str
    base_url: str
    model: str
    api_key: str
    supports_vision: bool

    @property
    def catalog_id(self) -> str:
        return f"connection:{self.id}"


def _config_path() -> Path:
    configured_directory = os.getenv("EXCEL_BRO_CONFIG_DIR", "").strip()
    if configured_directory:
        return Path(configured_directory) / ".env"
    return Path("server/.env")


def _connection_store_path() -> Path:
    return _config_path().parent / "model-connections.json"


def load_managed_connections() -> tuple[ManagedModelConnection, ...]:
    path = _connection_store_path()
    if not path.exists():
        return ()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except OSError as error:
        raise ModelConnectionStoreError(
            "无法读取模型连接配置文件"
        ) from error
    except ValueError as error:
        raise ModelConnectionStoreError(
            "模型连接配置文件已损坏，请先修复或备份该文件"
        ) from error
    if not isinstance(payload, dict) or not isinstance(
        payload.get("connections", []), list
    ):
        raise ModelConnectionStoreError("模型连接配置文件格式无效")
    items = payload.get("connections", [])
    connections: list[ManagedModelConnection] = []
    for index, item in enumerate(items):
        if not isinstance(item, dict):
            raise ModelConnectionStoreError(
                f"模型连接配置文件第 {index + 1} 项格式无效"
            )
        connection_id = str(item.get("id", "")).strip()
        label = str(item.get("label", "")).strip()
        base_url = str(item.get("baseUrl", "")).strip().rstrip("/")
        model = str(item.get("modelId", "")).strip()
        api_key = str(item.get("apiKey", "")).strip()
        if not connection_id or not label or not base_url or not model:
            raise ModelConnectionStoreError(
                f"模型连接配置文件第 {index + 1} 项缺少必要字段"
            )
        connections.append(
            ManagedModelConnection(
                id=connection_id,
                label=label,
                base_url=base_url,
                model=model,
                api_key=api_key,
                supports_vision=item.get("supportsVision") is True,
            )
        )
    return tuple(connections)


def load_formula_model_id() -> str:
    """读取顶层 formulaModelId：/function 专用模型选择。空串=跟随全局。"""
    path = _connection_store_path()
    if not path.exists():
        return ""
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return ""
    if not isinstance(payload, dict):
        return ""
    return str(payload.get("formulaModelId", "")).strip()


def _write_managed_connections(
    connections: tuple[ManagedModelConnection, ...],
    formula_model_id: str | None = None,
) -> None:
    path = _connection_store_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    # 保留已有的 formulaModelId（除非本次显式传入新值），避免写连接时丢掉公式模型选择。
    resolved_formula_id = (
        load_formula_model_id() if formula_model_id is None else formula_model_id
    )
    payload = {
        "version": 1,
        "formulaModelId": resolved_formula_id,
        "connections": [
            {
                "id": connection.id,
                "label": connection.label,
                "baseUrl": connection.base_url,
                "modelId": connection.model,
                "apiKey": connection.api_key,
                "supportsVision": connection.supports_vision,
            }
            for connection in connections
        ],
    }
    temporary = path.with_suffix(f"{path.suffix}.tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def upsert_managed_connection(
    *,
    connection_id: str | None,
    label: str,
    base_url: str,
    model: str,
    api_key: str | None,
    clear_api_key: bool,
    supports_vision: bool,
) -> dict[str, object]:
    current = list(load_managed_connections())
    existing = next(
        (
            connection
            for connection in current
            if connection.id == connection_id
        ),
        None,
    )
    if connection_id and existing is None:
        raise ModelConnectionNotFoundError("要编辑的模型连接不存在")
    normalized_base_url = base_url.strip().rstrip("/")
    normalized_model = model.strip()
    duplicate = next(
        (
            connection
            for connection in current
            if connection.id != connection_id
            and connection.base_url.casefold() == normalized_base_url.casefold()
            and connection.model.casefold() == normalized_model.casefold()
        ),
        None,
    )
    if duplicate is not None:
        raise DuplicateModelConnectionError(
            f"模型「{duplicate.label} · {duplicate.model}」已经使用相同的服务地址"
        )
    resolved_id = connection_id or f"model-{uuid.uuid4().hex}"
    resolved_api_key = (
        ""
        if clear_api_key
        else api_key.strip()
        if api_key is not None
        else existing.api_key
        if existing
        else ""
    )
    replacement = ManagedModelConnection(
        id=resolved_id,
        label=label.strip(),
        base_url=normalized_base_url,
        model=normalized_model,
        api_key=resolved_api_key,
        supports_vision=supports_vision,
    )
    next_connections = tuple(
        replacement if connection.id == resolved_id else connection
        for connection in current
    )
    if existing is None:
        next_connections = (*next_connections, replacement)
    _write_managed_connections(next_connections)
    return model_settings_view()


def draft_model_connection(
    *,
    connection_id: str | None,
    base_url: str,
    model: str,
    api_key: str | None,
    clear_api_key: bool,
    supports_vision: bool,
) -> ModelConnection:
    existing = next(
        (
            connection
            for connection in load_managed_connections()
            if connection.id == connection_id
        ),
        None,
    )
    if connection_id and existing is None:
        raise ModelConnectionNotFoundError("要测试的模型连接不存在")
    resolved_api_key = (
        ""
        if clear_api_key
        else api_key.strip()
        if api_key is not None
        else existing.api_key
        if existing
        else ""
    )
    return ModelConnection(
        base_url=base_url.strip().rstrip("/"),
        model=model.strip(),
        api_key=resolved_api_key,
        supports_vision=supports_vision,
    )


def delete_managed_connection(connection_id: str) -> dict[str, object]:
    current = load_managed_connections()
    remaining = tuple(
        connection
        for connection in current
        if connection.id != connection_id
    )
    if len(remaining) == len(current):
        raise ModelConnectionNotFoundError("要删除的模型连接不存在")
    # 删掉的正好是公式专用模型时，清空 formulaModelId，回退到跟随全局。
    formula_id = load_formula_model_id()
    deleted_catalog_id = f"connection:{connection_id}"
    next_formula_id = "" if formula_id == deleted_catalog_id else formula_id
    _write_managed_connections(remaining, formula_model_id=next_formula_id)
    return model_settings_view()


def set_formula_model_id(model_id: str) -> dict[str, object]:
    """设置 /function 公式专用模型。空串=跟随全局；否则须是已配置的模型。"""
    normalized = (model_id or "").strip()
    if normalized:
        catalog = model_catalog()
        valid_ids = {
            str(option["id"])
            for option in catalog["models"]  # type: ignore[index]
            if str(option["id"]) != "local"
        }
        if normalized not in valid_ids:
            raise ValueError("所选公式模型不存在或未配置")
    _write_managed_connections(
        load_managed_connections(), formula_model_id=normalized
    )
    return model_settings_view()


def load_model_settings() -> ModelSettings | None:
    base_url = _first_env("AI_BASE_URL", "OPENAI_BASE_URL").rstrip("/")
    default_model = _first_env("AI_MODEL", "OPENAI_MODEL")
    models = tuple(
        dict.fromkeys(
            model
            for model in [default_model, *_csv_env("AI_MODELS")]
            if model
        )
    )
    if not base_url or not models:
        return None
    if not default_model:
        default_model = models[0]
    vision_models = frozenset(
        model.casefold() for model in _csv_env("AI_VISION_MODELS")
    )
    return ModelSettings(
        base_url=base_url,
        default_model=default_model,
        api_key=_first_env("AI_API_KEY", "OPENAI_API_KEY"),
        models=models,
        vision_models=vision_models,
    )


def selected_model_config(
    model_id: str | None = None,
) -> ModelConnection | None:
    if model_id == "local":
        return None
    managed = load_managed_connections()
    if model_id and model_id.startswith("connection:"):
        selected = next(
            (
                connection
                for connection in managed
                if connection.catalog_id == model_id
            ),
            None,
        )
        if selected is None:
            raise ValueError("所选模型连接不存在")
        return ModelConnection(
            base_url=selected.base_url,
            model=selected.model,
            api_key=selected.api_key,
            supports_vision=selected.supports_vision,
        )
    settings = load_model_settings()
    if settings is None:
        if model_id:
            raise ValueError("所选模型尚未在本地服务中配置")
        if managed:
            selected = managed[0]
            return ModelConnection(
                base_url=selected.base_url,
                model=selected.model,
                api_key=selected.api_key,
                supports_vision=selected.supports_vision,
            )
        return None
    selected_model = model_id or settings.default_model
    if selected_model not in settings.models:
        raise ValueError(f"模型「{selected_model}」不在服务端允许列表中")
    return ModelConnection(
        base_url=settings.base_url,
        model=selected_model,
        api_key=settings.api_key,
        supports_vision=settings.supports_vision(selected_model),
    )


def formula_model_config() -> ModelConnection | None:
    """解析 /function 公式专用模型（顶层 formulaModelId）为可调用连接。
    未设置（空串）则返回 None，调用方据此回退到全局选择逻辑。"""
    formula_id = load_formula_model_id()
    if not formula_id:
        return None
    return selected_model_config(formula_id)


def model_catalog() -> dict[str, object]:
    settings = load_model_settings()
    managed = load_managed_connections()
    options: list[dict[str, str | bool]] = [
        {
            "id": "local",
            "label": "基础模式",
            "provider": "local",
            "available": True,
            "supportsVision": False,
        }
    ]
    if settings is not None:
        options.extend(
            {
                "id": model,
                "label": model,
                "provider": "model",
                "available": True,
                "supportsVision": settings.supports_vision(model),
            }
            for model in settings.models
        )
    options.extend(
        {
            "id": connection.catalog_id,
            "label": (
                connection.model
                if connection.label == connection.model
                else f"{connection.label} · {connection.model}"
            ),
            "provider": "model",
            "available": True,
            "supportsVision": connection.supports_vision,
        }
        for connection in managed
    )
    return {
        "defaultModelId": (
            settings.default_model
            if settings is not None
            else managed[0].catalog_id if managed else "local"
        ),
        "models": options,
    }


def model_status() -> dict[str, str | bool | None]:
    settings = load_model_settings()
    managed = load_managed_connections()
    if settings is None and not managed:
        return {
            "configured": False,
            "mode": "local",
            "model": None,
        }
    return {
        "configured": True,
        "mode": "model",
        "model": (
            settings.default_model
            if settings is not None
            else managed[0].model
        ),
    }


def model_settings_view() -> dict[str, object]:
    settings = load_model_settings()
    api_key = _first_env("AI_API_KEY", "OPENAI_API_KEY")
    connections = load_managed_connections()
    return {
        "baseUrl": settings.base_url if settings else None,
        "defaultModel": settings.default_model if settings else None,
        "apiKeyConfigured": bool(api_key),
        "apiKeyHint": f"••••{api_key[-4:]}" if api_key else None,
        "formulaModelId": load_formula_model_id(),
        "connections": [
            {
                "id": connection.id,
                "catalogModelId": connection.catalog_id,
                "label": connection.label,
                "baseUrl": connection.base_url,
                "modelId": connection.model,
                "supportsVision": connection.supports_vision,
                "apiKeyConfigured": bool(connection.api_key),
                "apiKeyHint": (
                    f"••••{connection.api_key[-4:]}"
                    if connection.api_key
                    else None
                ),
            }
            for connection in connections
        ],
    }


def update_model_api_key(api_key: str) -> dict[str, object]:
    config_path = _config_path()
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.touch(exist_ok=True)
    set_key(
        str(config_path),
        "AI_API_KEY",
        api_key,
        quote_mode="always",
    )
    os.environ["AI_API_KEY"] = api_key
    return model_settings_view()

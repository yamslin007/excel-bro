from __future__ import annotations

import os
import json
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlsplit
import uuid

from dotenv import set_key


class ModelConnectionStoreError(RuntimeError):
    pass


class ModelConnectionNotFoundError(ValueError):
    pass


class DuplicateModelConnectionError(ValueError):
    pass


class ModelNotConfiguredError(ValueError):
    """没有配置任何真实模型连接时，需要模型的入口必须显式失败。

    这是稳定错误含义：调用方据此返回 MODEL_NOT_CONFIGURED，不得静默
    降级到规则模式，也不得擅自改用其他供应商。
    """


class ModelNotAvailableError(ValueError):
    """请求指定的模型 ID 不存在、未配置或已失效。"""


# 旧版本前端会在模型选择里保存虚拟的 "local"，它代表“基础模式”而不是一个
# 真实模型连接。保留这个常量只用于识别和迁移旧记录，不再作为可选项返回。
LEGACY_LOCAL_MODEL_ID = "local"
CLAUDE_SDK_CATALOG_ID = "claude-agent-sdk"
CLAUDE_SDK_MANAGED_PREFIX = f"{CLAUDE_SDK_CATALOG_ID}:"
DEEPSEEK_ANTHROPIC_URL = "https://api.deepseek.com/anthropic"


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


def _claude_sdk_connection() -> ModelConnection | None:
    """The SDK is an explicit, environment-managed model choice.

    It never inherits an OpenAI-compatible connection's endpoint or key.
    """
    model = os.getenv("CLAUDE_AGENT_SDK_MODEL", "").strip()
    key = os.getenv("ANTHROPIC_API_KEY", "").strip()
    # The SDK stores local session transcripts. Never advertise it before the
    # operator has chosen a dedicated destination outside this repository.
    if not model or not key or not _claude_sdk_state_ready():
        return None
    return ModelConnection(
        base_url="https://api.anthropic.com",
        model=model,
        api_key=key,
        supports_vision=os.getenv("CLAUDE_AGENT_SDK_SUPPORTS_VISION", "").strip().lower() == "true",
        adapter="claude_agent_sdk",
    )


def _claude_sdk_state_ready() -> bool:
    state_dir = os.getenv("CLAUDE_AGENT_SDK_STATE_DIR", "").strip()
    if not state_dir:
        return False
    state_path = Path(state_dir)
    repository = Path(__file__).resolve().parents[3]
    if not state_path.is_absolute():
        return False
    state_path = state_path.resolve()
    if state_path == repository or repository in state_path.parents:
        return False
    return True


def _deepseek_anthropic_url(base_url: str) -> str | None:
    """Only translate DeepSeek's known OpenAI endpoint, never arbitrary URLs."""
    try:
        parsed = urlsplit(base_url.strip())
    except ValueError:
        return None
    if (
        parsed.scheme.lower() != "https"
        or parsed.netloc.lower() != "api.deepseek.com"
        or parsed.path.rstrip("/") not in ("", "/v1")
        or parsed.query or parsed.fragment
    ):
        return None
    return DEEPSEEK_ANTHROPIC_URL


def _managed_sdk_connection(
    model_id: str, managed: tuple[ManagedModelConnection, ...]
) -> ModelConnection | None:
    if not model_id.startswith(CLAUDE_SDK_MANAGED_PREFIX) or not _claude_sdk_state_ready():
        return None
    connection_id = model_id[len(CLAUDE_SDK_MANAGED_PREFIX):]
    selected = next((item for item in managed if item.id == connection_id), None)
    if selected is None or not selected.api_key:
        return None
    sdk_url = _deepseek_anthropic_url(selected.base_url)
    if sdk_url is None:
        return None
    return ModelConnection(
        base_url=sdk_url,
        model=selected.model,
        api_key=selected.api_key,
        supports_vision=selected.supports_vision,
        adapter="claude_agent_sdk",
    )


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
    """设置 /function 公式专用模型。空串=跟随全局；否则须是已配置的真实模型。

    旧记录里的 ``local`` 一律按“跟随全局”处理并写回空串，避免把已移除的
    基础模式入口又固化进配置文件。"""
    normalized = (model_id or "").strip()
    if normalized == LEGACY_LOCAL_MODEL_ID:
        normalized = ""
    if normalized:
        catalog = model_catalog()
        valid_ids = {
            str(option["id"])
            for option in catalog["models"]  # type: ignore[index]
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


def resolve_model_config(
    model_id: str | None = None,
) -> ModelConnection | None:
    """解析模型配置；无任何真实连接时返回 None（不抛错）。

    与 ``selected_model_config`` 的区别：这个函数只回答“当前能不能调用模型”，
    供健康检查、模型设置和已经明确允许无模型的固化工具使用。需要模型的入口
    应改用 ``selected_model_config``。
    """
    # 旧记录：优先按真实连接解析；没有任何真实连接时才是未配置。
    if model_id == CLAUDE_SDK_CATALOG_ID:
        return _claude_sdk_connection()
    managed = load_managed_connections()
    if model_id and model_id.startswith(CLAUDE_SDK_MANAGED_PREFIX):
        return _managed_sdk_connection(model_id, managed)
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
            return None
        return ModelConnection(
            base_url=selected.base_url,
            model=selected.model,
            api_key=selected.api_key,
            supports_vision=selected.supports_vision,
        )
    settings = load_model_settings()
    if settings is None:
        if model_id:
            return None
        if managed:
            selected = managed[0]
            return ModelConnection(
                base_url=selected.base_url,
                model=selected.model,
                api_key=selected.api_key,
                supports_vision=selected.supports_vision,
            )
        return _claude_sdk_connection()
    selected_model = model_id or settings.default_model
    if selected_model not in settings.models:
        return None
    return ModelConnection(
        base_url=settings.base_url,
        model=selected_model,
        api_key=settings.api_key,
        supports_vision=settings.supports_vision(selected_model),
    )


def selected_model_config(
    model_id: str | None = None,
) -> ModelConnection:
    """解析可调用的真实模型连接；不可用时抛出稳定错误。

    - 未配置任何真实模型（或旧记录里的 ``local``）：``ModelNotConfiguredError``
    - 指定了不存在/失效的模型 ID：``ModelNotAvailableError``

    绝不返回 None，也绝不回退到任何规则模式或另一个供应商。
    """
    if model_id == CLAUDE_SDK_CATALOG_ID:
        connection = _claude_sdk_connection()
        if connection is None:
            raise ModelNotAvailableError(
                "Claude Agent SDK 未配置；请设置模型、API Key，以及仓库外的绝对会话存储目录。"
            )
        return connection
    managed = load_managed_connections()
    if model_id and model_id.startswith(CLAUDE_SDK_MANAGED_PREFIX):
        connection = _managed_sdk_connection(model_id, managed)
        if connection is None:
            raise ModelNotAvailableError(
                "Agent SDK 对应的模型连接不存在、端点不兼容或会话目录未配置。"
            )
        return connection
    settings = load_model_settings()
    sdk_connection = _claude_sdk_connection()
    if not managed and settings is None and sdk_connection is None:
        raise ModelNotConfiguredError(
            "尚未配置任何模型连接。请填写服务地址、模型 ID 和 API Key 后重试。"
        )
    if model_id == LEGACY_LOCAL_MODEL_ID:
        # 旧版加载项可能仍提交 local。它不是一个真实连接，必须显式失败；
        # 新前端会在目录加载成功后迁移掉这条记录。
        raise ModelNotAvailableError(
            "“基础模式”已移除，请选择一个已配置的模型连接后重试。"
        )
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
            raise ModelNotAvailableError("所选模型连接不存在或已被删除。")
        return ModelConnection(
            base_url=selected.base_url,
            model=selected.model,
            api_key=selected.api_key,
            supports_vision=selected.supports_vision,
        )
    if settings is None:
        # 只有托管连接可用。空 model_id 时取第一个真实连接，显式指定的
        # 裸模型名（来自 .env 风格配置）在托管目录里找不到就是不可用。
        if not model_id:
            if not managed:
                assert sdk_connection is not None
                return sdk_connection
            selected = managed[0]
            return ModelConnection(
                base_url=selected.base_url,
                model=selected.model,
                api_key=selected.api_key,
                supports_vision=selected.supports_vision,
            )
        selected = next(
            (
                connection
                for connection in managed
                if connection.model.casefold() == model_id.casefold()
            ),
            None,
        )
        if selected is None:
            raise ModelNotAvailableError("所选模型尚未在本地服务中配置。")
        return ModelConnection(
            base_url=selected.base_url,
            model=selected.model,
            api_key=selected.api_key,
            supports_vision=selected.supports_vision,
        )
    selected_model = model_id or settings.default_model
    if selected_model not in settings.models:
        raise ModelNotAvailableError(
            f"模型「{selected_model}」不在服务端允许列表中。"
        )
    return ModelConnection(
        base_url=settings.base_url,
        model=selected_model,
        api_key=settings.api_key,
        supports_vision=settings.supports_vision(selected_model),
    )


def formula_model_config() -> ModelConnection | None:
    """解析 /function 公式专用模型（顶层 formulaModelId）为可调用连接。

    未设置（空串）返回 None，调用方据此回退到全局选择逻辑。旧记录指向
    已移除的 ``local`` 或已删除连接时同样返回 None（按“跟随全局”处理），
    调用方最终仍会用 ``selected_model_config`` 校验真实连接是否可用。"""
    formula_id = load_formula_model_id()
    if not formula_id or formula_id == LEGACY_LOCAL_MODEL_ID:
        return None
    managed = load_managed_connections()
    settings = load_model_settings()
    if not managed and settings is None:
        return None
    try:
        return selected_model_config(formula_id)
    except (ModelNotConfiguredError, ModelNotAvailableError):
        # 公式模型失效不应让 /function 直接失败：按跟随全局处理。
        return None


def model_catalog() -> dict[str, object]:
    """模型目录：只包含真实配置的模型连接。

    没有任何真实连接时返回 ``models=[]``、``defaultModelId=None``。不得用
    虚拟的 “基础模式” 选项占位来掩盖“未配置”或“目录读取失败”。
    """
    settings = load_model_settings()
    managed = load_managed_connections()
    options: list[dict[str, str | bool]] = []
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
    model_counts: dict[str, int] = {}
    for connection in managed:
        key = connection.model.casefold()
        model_counts[key] = model_counts.get(key, 0) + 1
    model_ordinals: dict[str, int] = {}
    for connection in managed:
        key = connection.model.casefold()
        model_ordinals[key] = model_ordinals.get(key, 0) + 1
        base_label = (
            connection.model
            if connection.label == connection.model
            else f"{connection.label} · {connection.model}"
        )
        suffix = f" ({model_ordinals[key]})" if model_counts[key] > 1 else ""
        options.append({
            "id": connection.catalog_id,
            "label": f"{base_label}{suffix}",
            "provider": "model",
            "available": True,
            "supportsVision": connection.supports_vision,
        })
        sdk_id = f"{CLAUDE_SDK_MANAGED_PREFIX}{connection.id}"
        if _managed_sdk_connection(sdk_id, managed) is not None:
            options.append({
                "id": sdk_id,
                "label": f"Agent SDK · {base_label}{suffix}",
                "provider": "model",
                "available": True,
                "supportsVision": connection.supports_vision,
            })
    sdk_connection = _claude_sdk_connection()
    if sdk_connection is not None:
        options.append({
            "id": CLAUDE_SDK_CATALOG_ID,
            "label": f"Claude Agent SDK · {sdk_connection.model}",
            "provider": "model",
            "available": True,
            "supportsVision": sdk_connection.supports_vision,
        })
    if settings is not None:
        default_model_id: str | None = settings.default_model
    elif managed:
        default_model_id = managed[0].catalog_id
    elif sdk_connection is not None:
        default_model_id = CLAUDE_SDK_CATALOG_ID
    else:
        default_model_id = None
    return {
        "defaultModelId": default_model_id,
        "models": options,
    }


def model_status() -> dict[str, str | bool | None]:
    settings = load_model_settings()
    managed = load_managed_connections()
    sdk_connection = _claude_sdk_connection()
    if settings is None and not managed and sdk_connection is None:
        return {
            "configured": False,
            "mode": "model",
            "model": None,
        }
    return {
        "configured": True,
        "mode": "model",
        "model": (
            settings.default_model
            if settings is not None
            else managed[0].model if managed else sdk_connection.model
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

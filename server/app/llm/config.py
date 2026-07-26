from __future__ import annotations

import os
from dataclasses import dataclass


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
    settings = load_model_settings()
    if settings is None:
        if model_id:
            raise ValueError("所选模型尚未在本地服务中配置")
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


def model_catalog() -> dict[str, object]:
    settings = load_model_settings()
    options: list[dict[str, str | bool]] = [
        {
            "id": "local",
            "label": "基础模式",
            "provider": "local",
            "available": True,
            "supportsVision": False,
        }
    ]
    if settings is None:
        return {"defaultModelId": "local", "models": options}
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
    return {
        "defaultModelId": settings.default_model,
        "models": options,
    }


def model_status() -> dict[str, str | bool | None]:
    settings = load_model_settings()
    if settings is None:
        return {
            "configured": False,
            "mode": "local",
            "model": None,
        }
    return {
        "configured": True,
        "mode": "model",
        "model": settings.default_model,
    }

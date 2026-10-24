from __future__ import annotations

import json

from ..models import ModelUsageErrorCode


class LLMError(RuntimeError):
    """Base exception for failures raised by an LLM adapter."""


class LLMTimeoutError(LLMError):
    pass


class LLMConnectError(LLMError):
    pass


class LLMTransportError(LLMError):
    pass


class LLMResponseError(LLMError):
    pass


MODEL_USAGE_MESSAGES: dict[ModelUsageErrorCode, str] = {
    "MODEL_QUOTA_EXHAUSTED": "模型 API 余额不足或调用额度已用尽。请在模型服务商处检查余额、调用额度或套餐，恢复后重试，或在模型设置中切换可用连接。",
    "MODEL_CONTEXT_LIMIT": "本次请求超过模型的上下文长度上限。请缩短对话或拆分任务后重试。",
    "MODEL_OUTPUT_LIMIT": "本次模型输出达到长度上限，未获得完整结果。请拆分任务或调整模型输出上限后重试。",
    "MODEL_BUDGET_LIMIT": "本次任务达到配置的模型费用上限。请检查任务预算设置，或拆分任务后重试。",
    "MODEL_TURN_LIMIT": "本次任务达到模型调用轮数上限。请拆分任务后重试。",
}


class LLMUsageLimitError(LLMResponseError):
    """An actionable limit; repeating the same request cannot resolve it."""

    def __init__(self, code: ModelUsageErrorCode) -> None:
        self.code = code
        super().__init__(MODEL_USAGE_MESSAGES[code])


def model_usage_error_code(body: str) -> ModelUsageErrorCode | None:
    """Classify only provider failure metadata, never successful model output.

    Inspect known error fields rather than echoed prompts/request data. Raw
    payloads are used only for classification and never in the usage notice.
    """
    bounded = body[:16384]
    try:
        payload = json.loads(bounded)
    except (ValueError, TypeError):
        payload = None
    if isinstance(payload, dict):
        error = payload.get("error", payload)
        if isinstance(error, dict):
            bounded = " ".join(str(error.get(key, "")) for key in ("code", "type", "message"))
        elif isinstance(error, str):
            bounded = error
        else:
            bounded = ""
    text = bounded.lower()
    if any(marker in text for marker in (
        "insufficient_quota", "insufficient_balance", "insufficient balance",
        "billing_error", "insufficient credits",
        "billing_hard_limit_reached", "credit balance is too low",
        "exceeded your current quota", "余额不足", "额度已用尽", "额度不足",
    )):
        return "MODEL_QUOTA_EXHAUSTED"
    if any(marker in text for marker in (
        "context_length_exceeded", "maximum context length", "prompt is too long",
        "上下文长度超限",
    )):
        return "MODEL_CONTEXT_LIMIT"
    return None


_STATUS_HINTS: dict[int, str] = {
    401: "模型服务拒绝了 API Key（未授权），请检查密钥是否正确或已过期。",
    402: "模型服务商余额不足，请充值或在设置里切换到其他模型连接。",
    403: "模型服务拒绝访问，请检查该 API Key 是否有权调用所选模型。",
    404: "模型服务找不到该接口或模型，请检查 baseUrl 和模型名。",
}


class LLMHTTPStatusError(LLMError):
    def __init__(self, status_code: int, body: str) -> None:
        self.status_code = status_code
        self.body = body.strip()[:300]
        self.usage_code: ModelUsageErrorCode | None = (
            "MODEL_QUOTA_EXHAUSTED" if status_code == 402 else
            model_usage_error_code(body) if status_code in {400, 403, 429} else None
        )
        self.retry_after: float | None = None
        super().__init__(self._compose_message())

    def _compose_message(self) -> str:
        """Known limits get safe actionable text, with no provider body attached."""
        hint = MODEL_USAGE_MESSAGES[self.usage_code] if self.usage_code else _STATUS_HINTS.get(self.status_code)
        if hint:
            return f"{hint}（模型服务返回 HTTP {self.status_code}）"
        return (
            f"模型服务返回 HTTP {self.status_code}"
            f"{f'：{self.body}' if self.body else ''}"
        )

    @property
    def retryable(self) -> bool:
        return self.usage_code is None and (self.status_code == 429 or self.status_code >= 500)

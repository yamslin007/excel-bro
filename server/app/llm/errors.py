from __future__ import annotations


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


class LLMHTTPStatusError(LLMError):
    def __init__(self, status_code: int, body: str) -> None:
        self.status_code = status_code
        self.body = body.strip()[:300]
        self.retry_after: float | None = None
        super().__init__(
            f"模型服务返回 HTTP {status_code}"
            f"{f'：{self.body}' if self.body else ''}"
        )

    @property
    def retryable(self) -> bool:
        return self.status_code == 429 or self.status_code >= 500

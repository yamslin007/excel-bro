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
        self.retry_after: float | None = None
        super().__init__(self._compose_message())

    def _compose_message(self) -> str:
        """可行动的状态码给中文提示，其余透传原始 body 便于排查。

        402 之类的错误上游只回一句英文 JSON（如 Insufficient Balance），直接透给
        用户看不出该干什么。提示放在前面，原始 body 附在括号里不丢排查信息。
        """
        hint = _STATUS_HINTS.get(self.status_code)
        if hint:
            return f"{hint}（模型服务返回 HTTP {self.status_code}）"
        return (
            f"模型服务返回 HTTP {self.status_code}"
            f"{f'：{self.body}' if self.body else ''}"
        )

    @property
    def retryable(self) -> bool:
        return self.status_code == 429 or self.status_code >= 500

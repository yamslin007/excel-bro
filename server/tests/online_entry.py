"""在线连续集成的**显式测试入口**（测试专用，不是生产模块，也不给生产加旁路开关）。

由 ``scripts/run_online_link.mjs`` 启动，只绑回环地址。它做三件事：

1. 打开隔离 Python 总开关 —— 走驱动写入的**临时能力表**
   （``EXCEL_BRO_CAPABILITIES_PATH``），不改生产 ``config/capabilities.json``；
2. 把 ``OpenAICompatibleClient.chat_completions`` 换成本文件里的**固定工具调用
   序列**。模型推理是本轮唯一被允许替代的环节，替身**只产出工具调用**，
   绝不调用任何内部函数去绕过分派、登记、编译或执行；
3. 用 uvicorn 起**真实** ASGI 应用（真实路由、真实工具处理、真实桥、真实
   协调器、真实 Docker 执行器、真实结果登记与白名单编译）。

环境变量
--------
``EB_ONLINE_PORT``       绑定的回环端口（必填）
``EB_ONLINE_SCENARIO``   场景 JSON 路径（默认 ``config/online_link_scenario.json``）
``EB_ONLINE_LOG``        相关日志落盘路径（可选；同时也写到 stdout）
``EXCEL_BRO_CAPABILITIES_PATH``  测试能力表（由驱动写入临时文件）
``AI_BASE_URL`` / ``AI_MODEL``   模型连接。真实端点**不会被访问**：推理已被
                          替身接住，配置只是为了让模型目录 / 健康检查可用。
"""

from __future__ import annotations

import json
import os
import pathlib
import sys
import time
from typing import Any

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

os.environ.setdefault("AI_BASE_URL", "http://127.0.0.1:1")
os.environ.setdefault("AI_MODEL", "online-test-model")

SCENARIO_PATH = pathlib.Path(
    os.environ.get(
        "EB_ONLINE_SCENARIO", str(REPO_ROOT / "config" / "online_link_scenario.json")
    )
)
SCENARIO = json.loads(SCENARIO_PATH.read_text(encoding="utf-8"))
CODE_PATH = REPO_ROOT / str(SCENARIO["analysisCodePath"])
ANALYSIS_CODE = CODE_PATH.read_text(encoding="utf-8")

_LOG_FILE = os.environ.get("EB_ONLINE_LOG")


def log(message: str) -> None:
    line = f"[online-server] {message}"
    print(line, flush=True)
    if _LOG_FILE:
        with open(_LOG_FILE, "a", encoding="utf-8") as handle:
            handle.write(f"{time.strftime('%H:%M:%S')} {line}\n")


# ---------------------------------------------------------------------------
# 固定的模型替身：只产出工具调用 / 结构化判定，不做任何内部调用
# ---------------------------------------------------------------------------


_USAGE = {"prompt_tokens": 1, "completion_tokens": 1, "total_tokens": 2}


def _completion(content: str | None, tool_calls: list[dict[str, Any]] | None = None):
    """``chat_completions`` 的返回契约是 ``(payload, usage)``，不是裸 payload。"""
    message: dict[str, Any] = {"role": "assistant", "content": content}
    if tool_calls:
        message["tool_calls"] = tool_calls
    payload = {
        "choices": [
            {"message": message, "finish_reason": "tool_calls" if tool_calls else "stop"}
        ],
        "usage": dict(_USAGE),
    }
    return payload, dict(_USAGE)


def _tool_call(name: str, arguments: dict[str, Any], call_id: str):
    return _completion(
        None,
        [
            {
                "id": call_id,
                "type": "function",
                "function": {
                    "name": name,
                    "arguments": json.dumps(arguments, ensure_ascii=False),
                },
            }
        ],
    )


def _system_prompt(messages: list[dict[str, Any]]) -> str:
    for message in messages:
        if message.get("role") == "system":
            return str(message.get("content") or "")
    return ""


def _last_tool_payload(messages: list[dict[str, Any]]) -> dict[str, Any] | None:
    for message in reversed(messages):
        if message.get("role") == "tool":
            try:
                payload = json.loads(message.get("content") or "{}")
            except ValueError:
                return {}
            return payload if isinstance(payload, dict) else {}
    return None


def _planning_call(messages: list[dict[str, Any]]):
    payload = _last_tool_payload(messages)
    if payload is None:
        sources = [
            {"sheet": SCENARIO["memberSheet"], "range": SCENARIO["memberRange"]},
            {"sheet": SCENARIO["orderSheet"], "range": SCENARIO["orderRange"]},
        ]
        log("model: prepare_python_data")
        return _tool_call("prepare_python_data", {"sources": sources}, "call-prepare-1")
    handles = payload.get("datasetHandles")
    if isinstance(handles, list) and handles:
        log(f"model: run_python_analysis handles={len(handles)}")
        return _tool_call(
            "run_python_analysis",
            {"code": ANALYSIS_CODE, "datasetHandles": handles},
            "call-run-1",
        )
    handle = payload.get("resultHandle")
    if isinstance(handle, str) and handle:
        log(f"model: submit_result_plan ref={handle}")
        return _tool_call(
            "submit_result_plan",
            {
                "resultRef": handle,
                "includeHeader": False,
                "title": "会员等级汇总",
                "target": {
                    "sheet": SCENARIO["targetSheet"],
                    "range": SCENARIO["targetRange"],
                },
            },
            "call-submit-1",
        )
    log(f"model: unexpected tool payload {sorted(payload)}")
    return _completion("工具回执无法继续，本轮结束。")


def _install_model_double() -> None:
    from server.app.llm import client as llm_client

    async def chat_completions(
        self,
        *,
        messages,
        max_tokens=None,
        temperature=None,
        tools=None,
        tool_choice=None,
    ):
        system = _system_prompt(list(messages))
        if "先理解消息" in system:
            return _completion(
                json.dumps(
                    {
                        "kind": "needs_workbook_context",
                        "summary": "按会员等级汇总订单金额",
                    },
                    ensure_ascii=False,
                )
            )
        if "判断用户是在普通交谈" in system:
            return _completion(
                json.dumps(
                    {
                        "kind": "proceed",
                        "summary": "按会员等级汇总订单金额",
                        "confirmedPrompt": SCENARIO["prompt"],
                    },
                    ensure_ascii=False,
                )
            )
        return _planning_call(list(messages))

    llm_client.OpenAICompatibleClient.chat_completions = chat_completions


def _watch_correlation(app: Any) -> None:
    """记录同一次任务的 taskId / operationId / 结果引用，便于事后对账。

    只落在测试入口的日志里（含 taskId、operationId、结果引用），不含任何源数据值。
    """
    watched = (
        "/api/sdk/tool-bridge/poll",
        "/api/sdk/tool-bridge/receipt",
        "/api/sdk/tool-bridge/result",
        "/api/turn",
    )

    @app.middleware("http")
    async def log_bridge_traffic(request, call_next):
        response = await call_next(request)
        path = request.url.path
        if path in watched:
            log(f"http {request.method} {path} -> {response.status_code}")
        return response

    async def log_scenario() -> None:
        log(f"scenario={SCENARIO_PATH.name} code={CODE_PATH.name}")

    app.add_event_handler("startup", log_scenario)


def main() -> int:
    port = os.environ.get("EB_ONLINE_PORT")
    if not port:
        log("EB_ONLINE_PORT 未设置，拒绝启动")
        return 2

    _install_model_double()

    from server.app.main import app  # noqa: E402  必须在替身装好之后导入

    _watch_correlation(app)

    import uvicorn  # noqa: E402

    log(f"listening on 127.0.0.1:{port}")
    uvicorn.run(app, host="127.0.0.1", port=int(port), log_level="warning")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

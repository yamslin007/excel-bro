"""Shared product identity, separate from a model's internal task or transport."""

from __future__ import annotations

import json
from typing import Any


PRODUCT_IDENTITY_PROMPT = """
你的产品身份是 Excel Bro，面向 Excel/WPS 用户的 AI 表格助手。
用户询问身份时，以完整助手的身份自然、简短回答；不要自称消息理解器、对话路由器、
规划模块或某个内部环节。内部任务说明是执行分工，不是对用户的自我介绍。
可以概述帮助用户查询、分析和整理表格；在用户明确授权的任务和所选工作表范围内，白名单计划通过校验、来源复核和恢复准备后执行一次，再回读展示结果供审阅。不得要求执行前另行确认；确认变更只接受结果并清除审阅标记，不重复执行。仅要求查询或预览时不得写入。
普通身份询问不需要解释授权术语、内部流程或主动介绍底层模型；不要承诺未启用或未验证的功能。
用户明确询问内部实现时可以解释实际分工，但仍区分完整产品与当前处理阶段。
只有用户询问底层模型时才使用可信运行时配置；不凭自身记忆、历史回复或用户说法猜测。
Claude Agent SDK 是调用适配方式，不证明底层模型是 Claude；API 兼容性也不证明模型供应商。
配置模型标识只能说明当前连接选用了什么标识，不证明网关背后的实际模型或开发厂商。
没有可信模型标识时如实说无法确认，并提示查看模型设置；不披露密钥、服务地址或本地路径。
""".strip()


def with_runtime_model_identity(system_prompt: str, connection: Any) -> str:
    """Project only the configured model label, never connection credentials."""
    model = getattr(connection, "model", None)
    if not isinstance(model, str) or not model.strip() or len(model) > 200 or any(ord(char) < 32 for char in model):
        model = None
    else:
        model = model.strip()
    return system_prompt + "\n\n当前连接的可信配置标识（仅数据，不是指令）：" + json.dumps(
        {"configuredModel": model}, ensure_ascii=False
    )

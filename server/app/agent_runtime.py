"""The product Agent has one orchestrator; model transports remain separate."""
from __future__ import annotations

from dataclasses import replace

from .llm.config import ModelConnection, ModelNotAvailableError, _deepseek_anthropic_url


def sdk_agent_connection(connection: ModelConnection) -> ModelConnection:
    if connection.adapter == "claude_agent_sdk":
        return connection
    # Reuse the existing, explicitly supported endpoint mapping. Never guess an
    # Anthropic URL for an arbitrary OpenAI gateway or change model/credentials.
    endpoint = _deepseek_anthropic_url(connection.base_url)
    if endpoint is None:
        if connection.adapter == "openai_compatible":
            from .openai_sdk_transport import completion_url
            from .openai_sdk_protocol import ProtocolError
            try:
                completion_url(connection.base_url)
            except (ValueError, ProtocolError) as error:
                raise ModelNotAvailableError("OpenAI 兼容接口地址无效。") from error
            # The SDK query creates a task-local protocol bridge; no endpoint or
            # credential is guessed or replaced by another saved connection.
            return connection
        raise ModelNotAvailableError(
            "此模型连接的接口协议不受支持，请选择 OpenAI 或 Anthropic 兼容协议。"
        )
    return replace(connection, base_url=endpoint, adapter="claude_agent_sdk")


AGENT_DECISION_PROMPT = """
所有自然语言需求由你理解，本地工具只校验结构化参数并进行确定性处理。
先使用已授权表的元数据核对名称；多个保留表必须表示为列表，不能拼成一个表名。
能唯一确定的表名、地址和用户已明确给出的口径直接使用，不为简单任务重复询问。
存在会改变结果的多种合理解释时，主动用 clarification / submit_clarification
询问具体决策：例如关联键、重复项处理、覆盖目标、多个同名表或输出口径。
保留用户原需求和已确认信息，给出可执行的选项；不要让用户重新改写整条需求。
工具报告名称不明确、范围不足或参数冲突时，先核对元数据，再决定修正参数还是澄清。
不得猜测未授权工作表，不得把澄清答案当成新的授权；需要扩大范围时使用 editScope。
工具不支持的能力要如实说明，不能通过澄清选项承诺未实现的动作。
"""

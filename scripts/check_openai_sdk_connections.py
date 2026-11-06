"""Explicit, synthetic SDK/tool probe for saved OpenAI-compatible connections.

Run from the repository root with .venv/Scripts/python.exe. No workbook/ledger
access, connection changes, or credential output. Requests use provider quota.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from server.app.claude_sdk_agent import check_sdk_connection
from server.app.llm.config import ModelConnection, load_managed_connections
from server.app.llm.errors import LLMError


async def run(labels: list[str]):
    results = []
    for connection in load_managed_connections():
        if connection.protocol != "openai_compatible" or (labels and connection.label not in labels):
            continue
        result = {"label": connection.label, "modelId": connection.model}
        try:
            await check_sdk_connection(ModelConnection(connection.base_url, connection.model,
                connection.api_key, connection.supports_vision))
            result["status"] = "passed"
        except LLMError as error:
            result.update(status="failed", error=type(error).__name__, reason=str(error))
        except Exception as error:
            result.update(status="failed", error=type(error).__name__)
        results.append(result)
        print(json.dumps(result, ensure_ascii=False), flush=True)
    return 0 if results and all(result["status"] == "passed" for result in results) else 1


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--label", action="append", default=[], help="Probe only this saved connection label; repeatable")
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="excelbro-sdk-connection-probe-") as directory:
        os.environ["CLAUDE_AGENT_SDK_STATE_DIR"] = directory
        raise SystemExit(asyncio.run(run(args.label)))

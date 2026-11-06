"""Probe saved connections with the unified text test; no workbook or config writes.

Only explicitly named connections are contacted. Provider quota is consumed.
"""
import argparse
import asyncio
from dataclasses import replace
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from server.app.llm.config import ModelConnection, load_managed_connections, _deepseek_anthropic_url
from server.app.llm.connection_test import test_model_connection


async def run(labels: list[str], also_deepseek_anthropic: bool) -> int:
    items = [item for item in load_managed_connections() if item.label in labels]
    if set(labels) - {item.label for item in items}:
        print("Some requested connection labels were not found.", flush=True)
        return 1
    failed = False
    for item in items:
        connection = ModelConnection(item.base_url, item.model, item.api_key, item.supports_vision, item.adapter)
        cases = [(item.protocol, connection)]
        endpoint = _deepseek_anthropic_url(item.base_url)
        if also_deepseek_anthropic and item.protocol == "openai_compatible" and endpoint:
            cases.append(("anthropic", replace(connection, base_url=endpoint, adapter="claude_agent_sdk")))
        for protocol, candidate in cases:
            result = {"label": item.label, "modelId": item.model, "protocol": protocol}
            try:
                await test_model_connection(candidate)
                result["status"] = "passed"
            except Exception as error:
                failed = True
                # Never print provider bodies, URLs with secrets, or credentials.
                result.update(status="failed", error=type(error).__name__)
            print(json.dumps(result, ensure_ascii=False), flush=True)
    return int(failed)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--label", action="append", required=True)
    parser.add_argument("--also-deepseek-anthropic", action="store_true",
        help="Also probe the repository's existing official DeepSeek Anthropic endpoint mapping; no config changes")
    args = parser.parse_args()
    raise SystemExit(asyncio.run(run(args.label, args.also_deepseek_anthropic)))

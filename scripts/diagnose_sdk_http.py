"""Bounded real-SDK / real-HTTP diagnostic; NOT real-host acceptance.

Uses the production app and adapter without replacing query/_result/model calls.
Only synthetic headers are supplied. The sole storage override isolates the
task ledger. Existing saved model connections are read normally; no keys logged.
Run with the selected SDK catalog id using --model. Costs model quota.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
from pathlib import Path
import socket
import sys
import time
import uuid
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


async def run(args):
    import httpx
    import uvicorn

    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    server = serving = sock = None
    if args.base_url:
        parsed = urlparse(args.base_url)
        if parsed.scheme != "http" or parsed.hostname != "127.0.0.1" or parsed.username or parsed.path not in ("", "/"):
            raise ValueError("External diagnostic requires an explicit HTTP IPv4 loopback origin")
        base_url = args.base_url.rstrip("/")
    else:
        from server.app import main
        from server.app.task_ledger import FolderOperationLedger
        main._folder_operation_ledger = FolderOperationLedger(path_provider=lambda: out / "tasks.sqlite3")
        sock = socket.socket()
        sock.bind(("127.0.0.1", 0))
        port = sock.getsockname()[1]
        server = uvicorn.Server(uvicorn.Config(main.app, log_level="warning", lifespan="off"))
        serving = asyncio.create_task(server.serve(sockets=[sock]))
        base_url = f"http://127.0.0.1:{port}"
    report = {"kind": "real_sdk_http_diagnostic_synthetic_host" if args.python else "real_sdk_http_diagnostic_synthetic_headers", "model": args.model, "runs": []}
    try:
        while server is not None and not server.started:
            if serving.done():
                await serving
                raise RuntimeError("Diagnostic server did not start")
            await asyncio.sleep(.05)
        report["baseUrl"] = base_url
        report["ledgerIsolated"] = not bool(args.base_url)
        async with httpx.AsyncClient(base_url=base_url, timeout=190, trust_env=False) as client:
            for mode in (("concurrent_health",) if args.python else ("baseline", "concurrent_health", "concurrent_health")):
                started = time.monotonic()
                turn = "turn-sdk-http-audit-" + uuid.uuid4().hex
                item = {"mode": mode, "turnId": turn, "events": [], "healthCount": 0}
                report["runs"].append(item)
                stop = asyncio.Event()

                async def health_loop():
                    while not stop.is_set():
                        response = await client.get("/health")
                        response.raise_for_status()
                        item["healthCount"] += 1
                        await asyncio.sleep(.4)

                health = asyncio.create_task(health_loop()) if mode != "baseline" else None
                try:
                    request = {
                        "turnId": turn, "modelId": args.model, "planningGoal": "read_only",
                        "prompt": "请调用 describe_authorized_tables 检查已授权的表结构，然后用 submit_answer 简短说明有哪些字段。只读，不查询数据、不生成写入计划。",
                        "planningReadPolicy": "none", "hostId": "wps-et", "hostVersion": "12.0",
                        "selfReportedActions": ["writeValues", "writeFormulas", "clearRange:contents"],
                        "workbook": {"name": "sdk-http-audit.xlsx", "capturedAt": "2026-10-11T00:00:00Z",
                                     "activeWorksheet": "测试表", "worksheets": [{"name": "测试表",
                                     "usedRange": "A1:B1", "rowCount": 1, "columnCount": 2,
                                     "headers": ["编号", "分类"]}]},
                    }
                    if args.python:
                        from scripts import drive_real_wps_sdk as synthetic_host
                        fixture = json.loads((ROOT / "config/online_link_scenario.json").read_text(encoding="utf-8"))
                        synthetic_host.SHEET_ROWS = {
                            fixture["memberSheet"]: fixture["memberRows"],
                            fixture["orderSheet"]: fixture["orderRows"],
                            fixture["targetSheet"]: [[""]],
                        }
                        request.update(planningGoal="write_preview", prompt=synthetic_host.PROMPT)
                        request["workbook"] = {
                            "name": "sdk-http-audit.xlsx", "workbookIdentity": "wps-path:c:\\sdk-http-audit.xlsx",
                            "capturedAt": "2026-10-11T00:00:00Z", "activeWorksheet": fixture["memberSheet"],
                            "sourceFingerprint": "audit-synthetic-v1", "worksheets": [
                                {"name": name, "usedRange": fixture[range_key], "rowCount": len(rows),
                                 "columnCount": len(rows[0]), "headers": rows[0]}
                                for name, range_key, rows in [
                                    (fixture["memberSheet"], "memberRange", fixture["memberRows"]),
                                    (fixture["orderSheet"], "orderRange", fixture["orderRows"]),
                                    (fixture["targetSheet"], "targetRange", [[""]]),
                                ]],
                        }
                        binding = {"taskId": turn, "hostId": "wps-et", "workbookName": request["workbook"]["name"],
                                   "workbookIdentity": request["workbook"]["workbookIdentity"],
                                   "sourceFingerprint": "audit-synthetic-v1"}
                        item["bridgeReceipts"] = []
                    async with asyncio.timeout(190):
                        async with client.stream("POST", "/api/turn/stream", json=request) as response:
                            response.raise_for_status()
                            event = None
                            async for line in response.aiter_lines():
                                if line.startswith("event:"):
                                    event = line[6:].strip()
                                elif line.startswith("data:"):
                                    payload = json.loads(line[5:])
                                    # Retain metadata only, never model content or credentials.
                                    item["events"].append({"at": round(time.monotonic()-started, 3),
                                        "event": event, "stepId": payload.get("stepId"),
                                        "toolName": payload.get("toolName"), "kind": payload.get("kind"),
                                        "code": payload.get("code")})
                                    if event == "result":
                                        item["resultKind"] = payload.get("kind")
                                        if args.python and payload.get("kind") == "plan":
                                            plan = payload["plan"]
                                            item["resultRef"] = plan.get("localResultRef")
                                            item["actionTypes"] = [action["type"] for action in plan["actions"]]
                                            fetched = await client.post("/api/sdk/tool-bridge/result", json={**binding, "resultRef": item["resultRef"]})
                                            item["resultHttpStatus"] = fetched.status_code
                                            if fetched.is_success:
                                                (out / "synthetic-result.json").write_text(json.dumps(fetched.json(), ensure_ascii=False, indent=2), encoding="utf-8")
                                    if args.python and event == "step" and payload.get("bridgeTool"):
                                        polled = await client.post("/api/sdk/tool-bridge/poll", json=binding)
                                        polled.raise_for_status()
                                        for call in polled.json().get("calls", []):
                                            receipt = ({"datasets": synthetic_host.build_datasets(call["arguments"])}
                                                if call["tool"] == "prepare_python_data" else {"errorCode": "unsupported_tool"})
                                            delivered = await client.post("/api/sdk/tool-bridge/receipt", json={**binding,
                                                "operationId": call["operationId"], "requestFingerprint": call["requestFingerprint"], "receipt": receipt})
                                            delivered.raise_for_status()
                                            item["bridgeReceipts"].append({"tool": call["tool"], "status": delivered.status_code})
                                    if event == "error":
                                        item["errorCode"] = payload.get("code")
                except Exception as exc:
                    item["exceptionType"] = type(exc).__name__
                finally:
                    stop.set()
                    if health:
                        health.cancel()
                        try:
                            await health
                        except asyncio.CancelledError:
                            pass
                    item["elapsedSeconds"] = round(time.monotonic()-started, 3)
                    item["passed"] = (item.get("resultKind") == ("plan" if args.python else "answer")
                        and not item.get("errorCode") and (not args.python or item.get("resultHttpStatus") == 200))
                    (out / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
                    print(json.dumps(item, ensure_ascii=False), flush=True)
    finally:
        if server is not None:
            server.should_exit = True
            await serving
            sock.close()
    return 0 if all(item.get("passed") for item in report["runs"]) else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", required=True)
    parser.add_argument("--out", default="outputs/sdk-http-audit-20261011")
    parser.add_argument("--python", action="store_true", help="Real model + real Docker, synthetic host values; no workbook writes")
    parser.add_argument("--base-url", help="Use a running loopback backend and its real ledger; no configuration changes to that backend")
    args = parser.parse_args()
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    config = json.loads((ROOT / "config/capabilities.json").read_text(encoding="utf-8"))
    config["taskRecovery"]["planningHardWallTimeSeconds"] = 180 if args.python else 90
    config["agent"]["claudeSdkMaxBudgetUsd"] = .40 if args.python else .10
    config["pythonExecution"]["enabled"] = bool(args.python)
    cap_path = out / "capabilities.json"
    cap_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
    os.environ["EXCEL_BRO_CAPABILITIES_PATH"] = str(cap_path)
    os.environ.setdefault("CLAUDE_AGENT_SDK_STATE_DIR", str(Path(os.environ["LOCALAPPDATA"]) / "ExcelBro" / "ClaudeAgentSdk"))
    raise SystemExit(asyncio.run(run(args)))

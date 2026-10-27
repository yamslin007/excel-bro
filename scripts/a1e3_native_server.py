"""Test-instance-only audit wrapper. Normal server.app.main startup never imports this."""
from __future__ import annotations

import copy
import json
import os
import secrets
import threading
import time
from pathlib import Path
from fastapi import FastAPI, HTTPException, Request


class FaultAudit:
    """Record real fault return values; never replace the fault's decision."""
    def __init__(self, manifest: dict, directory: Path):
        self.run_id = manifest["runId"]
        self.directory = directory
        self.lock = threading.RLock()
        self.current: dict = {}
        self.targets: list[dict] = []
        self.error: str | None = None

    def save(self):
        try:
            path = self.directory / "attempts" / (self.current["attemptId"] + ".json")
            path.parent.mkdir(parents=True, exist_ok=True)
            temp = path.with_suffix(".tmp")
            temp.write_text(json.dumps(self.current, ensure_ascii=False, indent=2), encoding="utf-8")
            temp.replace(path)
        except (OSError, KeyError) as error:
            # Diagnostic failure must not change product behavior, or count as evidence.
            self.error = type(error).__name__

    def snapshot(self):
        with self.lock:
            return copy.deepcopy({"runId": self.run_id, "attempt": self.current,
                                  "bridgeTargets": self.targets, "auditError": self.error})

    def install(self, fault):
        originals = {name: getattr(fault, name) for name in ("arm", "bind_task", "take", "cancel")}

        def arm(**kwargs):
            with self.lock:
                result = originals["arm"](**kwargs)
                self.current = {"attemptId": secrets.token_hex(16), "runId": self.run_id,
                                "armedAt": time.time(), "target": {
                                    "hostId": kwargs["host_id"], "workbookName": kwargs["workbook_name"],
                                    "workbookIdentity": kwargs["workbook_identity"]},
                                "boundTaskId": None, "consumedTaskId": None, "cancelled": False}
                self.error = None
                self.save()
                return result

        def bind_task(**kwargs):
            with self.lock:
                result = originals["bind_task"](**kwargs)
                target = {"hostId": kwargs["host_id"], "workbookName": kwargs["workbook_name"],
                          "workbookIdentity": kwargs["workbook_identity"]}
                if target not in self.targets:
                    self.targets = (self.targets + [target])[-32:]
                if result and self.current:
                    self.current.update(boundTaskId=kwargs["task_id"], boundAt=time.time())
                    self.save()
                return result

        def take(task_id):
            with self.lock:
                result = originals["take"](task_id)
                if result and self.current:
                    self.current.update(consumedTaskId=task_id, consumedAt=time.time())
                    self.save()
                return result

        def cancel():
            with self.lock:
                result = originals["cancel"]()
                if self.current:
                    self.current.update(cancelled=True, cancelledAt=time.time())
                    self.save()
                return result

        for name, implementation in locals().copy().items():
            if name in originals:
                setattr(fault, name, implementation)
        return originals


def create_app():
    path = Path(os.environ["EB_NATIVE_MANIFEST"]).resolve()
    manifest = json.loads(path.read_text(encoding="utf-8-sig"))
    expected_config = (path.parent / "server-config").resolve()
    if Path(os.environ.get("EXCEL_BRO_CONFIG_DIR", "")).resolve() != expected_config:
        raise RuntimeError("Native acceptance config does not match manifest")
    if os.environ.get("EXCEL_BRO_SOURCE_CHANGED_FAULT") != "1":
        raise RuntimeError("Native acceptance fault is not enabled")
    from server.app import main, dev_source_changed_fault
    from server.app.llm.config import _claude_sdk_state_ready
    if not _claude_sdk_state_ready():
        raise RuntimeError("SDK state must be an absolute directory outside the repository")
    audit = FaultAudit(manifest, path.parent)
    audit.install(dev_source_changed_fault)
    wrapper = FastAPI(lifespan=main.app.router.lifespan_context)

    @wrapper.get("/api/dev/native-acceptance/session")
    async def session(request: Request):
        main._bridge_local_guard(request)
        supplied = request.headers.get("x-native-acceptance-token", "")
        if not secrets.compare_digest(supplied, manifest["token"]):
            raise HTTPException(403, "Native acceptance token mismatch")
        return audit.snapshot()

    wrapper.mount("/", main.app)
    wrapper.state.native_audit = audit
    return wrapper

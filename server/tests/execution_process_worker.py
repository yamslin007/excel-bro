"""Isolated A1 test worker; only uses the supplied temporary ledger."""
import json
import os
from pathlib import Path
import sys
import threading

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
db = Path(sys.argv[1]).resolve()
os.environ["EXCEL_BRO_CONFIG_DIR"] = str(db.parent / "worker-config")
from fastapi.testclient import TestClient
import server.app.main as main
from server.app.task_ledger import FolderOperationLedger

main._folder_operation_ledger = FolderOperationLedger(lambda: db)
request = json.loads(sys.stdin.readline())
with TestClient(main.app) as client:
    response = client.request(request["method"], request["path"], json=request.get("body"))
    result = json.dumps(dict(status=response.status_code, body=response.json()))
    if len(sys.argv) > 2:
        ready = Path(sys.argv[2])
        staging = ready.with_suffix(".writing")
        staging.write_text(result)
        staging.replace(ready)
        threading.Event().wait()  # Parent kills this process, no graceful cleanup/suspend.
    else:
        print(result, flush=True)

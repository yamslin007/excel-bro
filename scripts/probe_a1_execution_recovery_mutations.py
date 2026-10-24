"""Run A1 mutations in isolated processes; never edit production source files."""
from __future__ import annotations

import argparse
import hashlib
import inspect
import subprocess
import sys
import textwrap
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
MUTATIONS = {
    "A1_INFLIGHT": ("claim_office_resume", "AND in_flight_action_index IS NULL", "", "inflight_cannot_claim"),
    "A1_CURSOR": ("reconcile_office_action", 'action_index + 1 if outcome == "applied" else action_index', "action_index", "lost_receipt_resumes"),
    "A1_CAS": ("claim_office_resume", "WHERE operation_id = ? AND revision = ?", "WHERE operation_id = ? AND ? IS NOT NULL", "concurrent_resume"),
    "A1_OWNER": ("complete_office_action", "AND owner_token = ?", "AND ? IS NOT NULL", "concurrent_resume"),
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--mutant", choices=MUTATIONS)
    parser.add_argument("--output", default="outputs/a1-execution-recovery")
    args = parser.parse_args()
    if args.mutant:
        import pytest
        import server.app.task_ledger as module
        method, old, new, selector = MUTATIONS[args.mutant]
        source = textwrap.dedent(inspect.getsource(getattr(module.FolderOperationLedger, method)))
        if source.count(old) != 1:
            raise RuntimeError("Mutation target is not unique")
        namespace = dict(vars(module))
        exec(compile(source.replace(old, new), f"<mutation:{args.mutant}>", "exec"), namespace)
        setattr(module.FolderOperationLedger, method, namespace[method])
        return pytest.main(["server/tests/test_execution_recovery_a1.py", "-q", "-p", "no:cacheprovider", "-k", selector,
                            "--basetemp", str(ROOT / (".tmp-a1-mut-" + uuid.uuid4().hex))])
    output = ROOT / args.output
    output.mkdir(parents=True, exist_ok=True)
    production = ROOT / "server/app/task_ledger.py"
    original = hashlib.sha256(production.read_bytes()).hexdigest()
    success = True
    for name in MUTATIONS:
        result = subprocess.run([sys.executable, str(Path(__file__).resolve()), "--mutant", name], cwd=ROOT, capture_output=True)
        (output / f"{name}.log").write_bytes(result.stdout + result.stderr)
        # pytest exit 1 means assertions failed; collection/setup/tool errors are not evidence.
        valid = result.returncode == 1 and b"FAILED server/tests/test_execution_recovery_a1.py::test_a1_" in result.stdout and b"ERROR " not in result.stdout
        success &= valid
        print(f"{name}: {'RED (detected)' if valid else 'NOT PROVEN'} exit={result.returncode}", flush=True)
    current = hashlib.sha256(production.read_bytes()).hexdigest()
    assert current == original, "Production source changed during mutation run"
    print(f"Production SHA256 unchanged: {current}")
    return 0 if success else 1


if __name__ == "__main__":
    raise SystemExit(main())

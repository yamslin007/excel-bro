"""Scoped mutation checks. Run serially, with no other tests/editors on these files."""
from pathlib import Path
import hashlib
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "bridge-persistence-mutations"
OUT.mkdir(parents=True, exist_ok=True)
NODE = shutil.which("node")
VITEST = str(ROOT / "node_modules/vitest/vitest.mjs")


def vitest(workspace, test, pattern):
    return [NODE, VITEST, "run", test, "-t", pattern], ROOT / workspace


def pytest(test):
    return [sys.executable, "-m", "pytest", "server/tests/test_checkpoint_persistence.py",
            "-q", "-k", test, "-p", "no:cacheprovider", "--basetemp",
            tempfile.mkdtemp(prefix="eb-persistence-mutation-")], ROOT


cases = [
    ("P1", "apps/wps-addin/src/wpsReadonlyAgent.ts",
     "if (input.preparePlanningBridge) {", "if (false && input.preparePlanningBridge) {",
     lambda: vitest("apps/wps-addin", "src/wpsReadonlyAgent.test.ts", "prepares the bridge"),
     "Planning completed without awaiting bridge preparation"),
    ("P2", "apps/excel-addin/src/hooks/useSdkToolBridge.ts",
     'failureRef.current = "无法读取已授权工作表，未建立读表桥。";\n      return false;',
     'failureRef.current = "无法读取已授权工作表，未建立读表桥。";\n      return true;',
     lambda: vitest("apps/excel-addin", "src/hooks/useSdkToolBridge.test.tsx",
                    "does not report bridge readiness"),
     "unreadable authorization must stop the planning request"),
    ("P3", "server/app/worksheet_checkpoint.py",
     'attempts = capability_int("worksheetCheckpoint", "stateWriteAttempts")', "attempts = 1",
     lambda: pytest("persistent_replace_keeps_last_committed_state"), "assert 1 < 1"),
    ("P4", "server/app/worksheet_checkpoint.py",
     "before_sheet=lambda name: self._record_restore_intent(folder, record, name),",
     "before_sheet=lambda name: None,",
     lambda: pytest("failed_restore_progress_is_unknown_after_reload_not_replayed"),
     "assert status.restore is not None"),
]


def run(case_id, phase, command_factory):
    command, cwd = command_factory()
    result = subprocess.run(command, cwd=cwd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            timeout=120)
    (OUT / f"{case_id}-{phase}.log").write_bytes(result.stdout)
    return result.returncode, result.stdout.decode("utf-8", errors="replace")


for case_id, relative, old, new, command_factory, evidence in cases:
    path = ROOT / relative
    before = path.read_bytes()
    digest = hashlib.sha256(before).hexdigest()
    # Match existing line endings; never rewrite unrelated bytes.
    newline = "\r\n" if b"\r\n" in before else "\n"
    needle = old.replace("\n", newline).encode("utf-8")
    replacement = new.replace("\n", newline).encode("utf-8")
    assert before.count(needle) == 1, f"{case_id}: seed must match exactly once"
    code, _ = run(case_id, "baseline", command_factory)
    assert code == 0, f"{case_id}: baseline not green; no mutation applied"
    try:
        path.write_bytes(before.replace(needle, replacement, 1))
        code, output = run(case_id, "mutation", command_factory)
        assert code == 1 and evidence in output, f"{case_id}: no expected business failure"
    finally:
        path.write_bytes(before)
        assert hashlib.sha256(path.read_bytes()).hexdigest() == digest
    code, _ = run(case_id, "restored", command_factory)
    assert code == 0, f"{case_id}: restored check failed"
    print(f"{case_id} RED(OK), restored GREEN, sha256={digest}", flush=True)

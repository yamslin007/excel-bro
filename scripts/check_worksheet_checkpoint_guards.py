"""Reverse business assertions for fixed worksheet checkpoint safety. No live workbook access."""
from pathlib import Path
import hashlib
import os
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
FILE = ROOT / "server/app/worksheet_checkpoint.py"
SEEDS = [
    ("C1", 'if missing and backup_consent is None:', 'if False:', "test_scope_gap_does_not_even_archive_unselected_sheet"),
    ("C5", 'if backup_consent is not None and (names != record["expectedSheetNames"] or missing != backup_consent):', 'if False:', "test_sheet_catalog_change_while_waiting_for_consent_blocks_archive"),
    ("C2", 'if actual != names:', 'if False:', "test_archive_sheet_mismatch_zero_delete"),
    ("C3", 'shutil.copyfile(archive, path)', 'path = archive', "test_restoration_copy_is_byte_identical_and_original_archive_immutable"),
    ("C4", 'if _validate_archive(archive, record["sheetNames"]) != record["sha256"]:', 'if False:', "test_corrupted_archive_cannot_be_opened"),
]

def main():
    original = FILE.read_bytes()
    digest = hashlib.sha256(original).hexdigest()
    source = original.decode("utf-8")
    failed = []
    for seed, before, after, test in SEEDS:
        if source.count(before) != 1: raise RuntimeError(f"{seed}: seed missing or ambiguous")
        if "--check" in sys.argv: continue
        try:
            FILE.write_bytes(source.replace(before, after, 1).encode("utf-8"))
            result = subprocess.run([sys.executable, "-m", "pytest", f"server/tests/test_worksheet_checkpoint.py::{test}", "-q", "-p", "no:cacheprovider", f"--basetemp=outputs/checkpoint-mutation-{seed}"], cwd=ROOT, capture_output=True, encoding="utf-8", errors="replace", env={**os.environ, "PYTHONUTF8": "1"})
            output = re.sub(r"\x1b\[[0-9;]*m", "", result.stdout + result.stderr)
            (ROOT / f"outputs/checkpoint-{seed}.log").write_text(output, encoding="utf-8")
            valid_red = result.returncode == 1 and "1 failed" in output and f"FAILED server/tests/test_worksheet_checkpoint.py::{test}" in output and ("AssertionError" in output or "Failed: DID NOT RAISE" in output) and "ERROR collecting" not in output
            print(f"{seed}: {'RED' if valid_red else 'INVALID'}", flush=True)
            if not valid_red: failed.append(seed)
        finally:
            FILE.write_bytes(original)
            assert hashlib.sha256(FILE.read_bytes()).hexdigest() == digest
    print(f"{len(SEEDS)} seeds; restored SHA256 {digest}; invalid={failed}")
    return int(bool(failed))

if __name__ == "__main__": raise SystemExit(main())

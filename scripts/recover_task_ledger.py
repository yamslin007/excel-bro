"""Manual, offline recovery: preserve execution records, quarantine broken planning.

Never runs from application startup. Requires a stopped local service and an
explicit source path. Original database and sidecars are retained byte-for-byte.
Unreadable non-planning tables abort recovery; no execution record is discarded.
"""
from __future__ import annotations
import argparse
from contextlib import closing
import hashlib
import json
from pathlib import Path
import shutil
import socket
import sqlite3
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
PLANNING = {"planning_tasks", "planning_task_events", "planning_budget_extensions"}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def recover(source: Path, backup: Path):
    source = source.resolve(strict=True)
    backup = backup.resolve()
    if source.name != "state.sqlite3" or backup.parent != source.parent or backup.exists():
        raise ValueError("Use state.sqlite3 and a new backup directory directly beside it")
    with socket.socket() as sock:
        sock.settimeout(1)
        if sock.connect_ex(("127.0.0.1", 8765)) == 0:
            raise RuntimeError("Stop the local backend before offline recovery")
    backup.mkdir()
    report = {"source": str(source), "backup": str(backup), "hashes": {}, "preserved": {},
              "quarantinedTables": sorted(PLANNING), "applied": False}
    for suffix in ("", "-wal", "-shm"):
        original = Path(str(source) + suffix)
        if original.exists():
            saved = backup / original.name
            shutil.copy2(original, saved)
            if digest(original) != digest(saved):
                raise RuntimeError("Backup checksum mismatch")
            report["hashes"][original.name] = digest(saved)
    candidate = backup / "recovered.sqlite3"
    from server.app.task_ledger import FolderOperationLedger
    with closing(sqlite3.connect(candidate, isolation_level=None)) as schema_connection:
        schema_connection.row_factory = sqlite3.Row
        FolderOperationLedger(path_provider=lambda: candidate)._initialize_schema(schema_connection)
    old = sqlite3.connect(source.as_uri() + "?mode=ro", uri=True)
    new = sqlite3.connect(candidate)
    try:
        tables = [row[0] for row in old.execute("SELECT name FROM sqlite_master WHERE type='table'")]
        for table in tables:
            if table in PLANNING:
                continue
            quoted = '"' + table.replace('"', '""') + '"'
            columns = [row[1] for row in old.execute(f"PRAGMA table_info({quoted})")]
            names = ','.join('"' + col.replace('"', '""') + '"' for col in columns)
            rows = old.execute(f"SELECT {names} FROM {quoted} NOT INDEXED ORDER BY rowid").fetchall()
            if rows:
                new.executemany(f"INSERT INTO {quoted} ({names}) VALUES ({','.join('?' for _ in columns)})", rows)
            recovered = new.execute(f"SELECT {names} FROM {quoted} ORDER BY rowid").fetchall()
            if recovered != rows:
                raise RuntimeError(f"Preservation check failed for {table}")
            report["preserved"][table] = len(rows)
        new.commit()
        if new.execute("PRAGMA integrity_check").fetchall() != [("ok",)]:
            raise RuntimeError("Recovered database integrity check failed")
        if new.execute("PRAGMA foreign_key_check").fetchall():
            raise RuntimeError("Recovered database reference check failed")
        new.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        new.execute("PRAGMA journal_mode=DELETE")
    finally:
        old.close()
        new.close()
    # Keep the original live file and its sidecars together as a second backup.
    moved = []
    try:
        for suffix in ("-wal", "-shm", ""):
            original = Path(str(source) + suffix)
            if original.exists():
                destination = backup / ("original-live.sqlite3" + suffix)
                original.rename(destination)
                moved.append((original, destination))
        candidate.rename(source)
    except BaseException:
        for original, destination in reversed(moved):
            if not original.exists():
                destination.rename(original)
        raise
    report["applied"] = True
    report["recoveredSha256"] = digest(source)
    (backup / "recovery-report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--backup", required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(recover(args.source, args.backup), ensure_ascii=False, indent=2))

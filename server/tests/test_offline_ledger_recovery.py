import sqlite3
from contextlib import closing, contextmanager

import pytest

from scripts import recover_task_ledger as recovery
from server.app.task_ledger import FolderOperationLedger


class OfflineSocket:
    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def settimeout(self, value):
        pass

    def connect_ex(self, address):
        return 1


def seeded_ledger(tmp_path):
    source = tmp_path / "state.sqlite3"
    ledger = FolderOperationLedger(lambda: source)
    connect = ledger._connect

    @contextmanager
    def closed_connection():
        with closing(connect()) as connection:
            with connection:
                yield connection

    ledger._connect = closed_connection
    record, _ = ledger.register_office_execution(
        "office-preserved", "fingerprint", plan={"id": "plan", "actions": []},
        authorization={"authorizedSheetNames": ["Target"]},
        workbook_identity="url:c:/work/book.xlsx", task_id="historical-task")
    with closing(sqlite3.connect(source)) as connection:
        connection.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        connection.execute("PRAGMA journal_mode=DELETE")
    return source, record


def corrupt_root(source, table):
    with closing(sqlite3.connect(source)) as connection:
        page = connection.execute("SELECT rootpage FROM sqlite_master WHERE name=?", (table,)).fetchone()[0]
        size = connection.execute("PRAGMA page_size").fetchone()[0]
    with source.open("r+b") as handle:
        handle.seek((page - 1) * size)
        handle.write(b"\xff")


def test_quarantines_corrupt_planning_and_preserves_execution_exactly(tmp_path, monkeypatch):
    monkeypatch.setattr(recovery.socket, "socket", OfflineSocket)
    source, record = seeded_ledger(tmp_path)
    corrupt_root(source, "planning_tasks")
    old_bytes = source.read_bytes()
    backup = tmp_path / "backup"
    report = recovery.recover(source, backup)
    assert report["preserved"]["office_executions"] == 1
    assert (backup / "state.sqlite3").read_bytes() == old_bytes
    assert FolderOperationLedger(lambda: source).office_execution("office-preserved") == record
    with sqlite3.connect(source) as connection:
        assert connection.execute("PRAGMA integrity_check").fetchall() == [("ok",)]
        assert connection.execute("SELECT count(*) FROM planning_tasks").fetchone() == (0,)


def test_unreadable_execution_aborts_without_replacing_source(tmp_path, monkeypatch):
    monkeypatch.setattr(recovery.socket, "socket", OfflineSocket)
    source, _ = seeded_ledger(tmp_path)
    corrupt_root(source, "office_executions")
    before = source.read_bytes()
    with pytest.raises(sqlite3.DatabaseError):
        recovery.recover(source, tmp_path / "backup")
    assert source.read_bytes() == before
    assert (tmp_path / "backup" / "state.sqlite3").read_bytes() == before


def test_running_backend_refuses_recovery(tmp_path, monkeypatch):
    monkeypatch.setattr(recovery.socket, "socket", OfflineSocket)
    monkeypatch.setattr(OfflineSocket, "connect_ex", lambda self, address: 0)
    source, _ = seeded_ledger(tmp_path)
    before = source.read_bytes()
    with pytest.raises(RuntimeError, match="Stop the local backend"):
        recovery.recover(source, tmp_path / "backup")
    assert source.read_bytes() == before
    assert not (tmp_path / "backup").exists()

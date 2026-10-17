import sqlite3
from fastapi.testclient import TestClient
from server.app import main


def test_corrupt_ledger_is_reported_without_raw_database_details(monkeypatch):
    error = sqlite3.DatabaseError("PRIVATE_DATABASE_PATH: damaged page")
    error.sqlite_errorcode = sqlite3.SQLITE_CORRUPT

    class BrokenLedger:
        def planning_task(self, task_id):
            raise error

    monkeypatch.setattr(main, "_folder_operation_ledger", BrokenLedger())
    response = TestClient(main.app).get("/api/turn/turn-corrupt-test/status")
    assert response.status_code == 503
    assert response.json()["detail"]["code"] == "TASK_STATE_CORRUPT"
    assert response.json()["detail"]["retryable"] is False
    assert "PRIVATE_DATABASE_PATH" not in response.text

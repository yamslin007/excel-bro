"""Regression tests for native acceptance orchestration, not native host evidence."""
from unittest import mock
import pytest
from scripts import a1e3_native as cli


def test_partial_zero_must_not_describe_zero():
    text = cli._describe_observation({"observation": "partial", "callEvidence": "zero"})
    assert "零" not in text


def test_cancel_posts_json():
    response = mock.MagicMock()
    response.__enter__.return_value.read.return_value = b'{}'
    with mock.patch.object(cli, "_health", return_value=True), mock.patch.object(cli.urllib.request, "urlopen", return_value=response) as send:
        cli.cmd_cancel("http://127.0.0.1:8877")
    request = send.call_args.args[0]
    assert request.get_header("Content-type") == "application/json"
    assert request.data == b'{}'


def test_missing_ledger_is_not_unentered(tmp_path):
    with mock.patch.object(cli, "_status", return_value={"phase": "consumed", "taskBound": False}):
        with pytest.raises(cli.ServiceError):
            cli._collect_evidence("http://127.0.0.1:8877", str(tmp_path / "missing"), None, tmp_path / "out")

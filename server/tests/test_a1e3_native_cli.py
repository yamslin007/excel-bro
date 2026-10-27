"""A1-E3 S1 原生验收 CLI（``scripts/a1e3_native.py``）的失败路径单元测试。

这里**不**起真实服务端、不 mock 判据：判据（``observation``、身份收敛）仍只有
服务端一份，本文件只验证 CLI 自己的**编排与失败路径**——目标歧义、服务不可达、
身份不在 seenTargets、启用失败、过期、未触发、证据缺失。用 ``unittest.mock``
替换 ``_request`` / ``_health`` / ``_status`` 的 HTTP 边界。

验证的是"脚本在该停的时候停、说清原因、不给假通过"，不是重写服务端判据。
"""

from __future__ import annotations

import unittest
import tempfile
from unittest import mock

import scripts.a1e3_native as cli  # type: ignore


def _seen(*targets: dict) -> list[dict]:
    return list(targets)


def _target(host="excel-office", name="测试.xlsx", identity="url:https://e/a.xlsx") -> dict:
    return {"hostId": host, "workbookName": name, "workbookIdentity": identity}


class _ArmCase(unittest.TestCase):
    """把 arm 的 HTTP 边界替换掉，只测编排分支。"""

    _patchers: list[mock._patch]

    def _patch(self, status: dict, arm: dict | None = None) -> None:
        self._patchers = [
            mock.patch.object(cli, "_health", return_value=True),
            mock.patch.object(cli, "_status", return_value=status),
            mock.patch.object(cli, "_request", return_value=arm or {}),
        ]
        for patcher in self._patchers:
            patcher.start()

    def tearDown(self) -> None:
        for patcher in getattr(self, "_patchers", []):
            patcher.stop()


class ArmFailures(_ArmCase):
    def test_no_seen_targets_stops(self):
        self._patch({"seenTargets": []})
        self.assertEqual(cli.cmd_arm("http://x", None, None, None, 120), 3)

    def test_ambiguous_candidates_refuse_default_first(self):
        status = {
            "seenTargets": _seen(
                _target(identity="url:https://e/a.xlsx"),
                _target(name="测试.xlsx", identity="url:https://e/b.xlsx"),
            )
        }
        # 同名两本：只给 name，两个候选，必须拒绝，不能默认取第一条。
        self._patch(status)
        self.assertEqual(cli.cmd_arm("http://x", "excel-office", "测试.xlsx", None, 120), 3)

    def test_identity_not_in_seen_stops(self):
        status = {"seenTargets": _seen(_target(identity="url:https://e/a.xlsx"))}
        self._patch(status)
        self.assertEqual(cli.cmd_arm("http://x", None, None, "url:https://e/wrong.xlsx", 120), 3)

    def test_arm_not_armed_stops(self):
        status = {"seenTargets": _seen(_target())}
        self._patch(status, arm={"phase": "idle", "identityBound": False})
        self.assertEqual(cli.cmd_arm("http://x", "excel-office", "测试.xlsx", None, 120), 3)

    def test_arm_expired_stops(self):
        status = {"seenTargets": _seen(_target())}
        self._patch(status, arm={"phase": "armed", "identityBound": True, "expiresInSeconds": 0})
        self.assertEqual(cli.cmd_arm("http://x", "excel-office", "测试.xlsx", None, 120), 3)

    def test_arm_success_returns_zero(self):
        status = {"seenTargets": _seen(_target())}
        self._patch(status, arm={"phase": "armed", "identityBound": True, "expiresInSeconds": 120})
        self.assertEqual(cli.cmd_arm("http://x", "excel-office", "测试.xlsx", None, 120), 0)

    def test_server_unreachable_stops(self):
        with mock.patch.object(cli, "_health", return_value=False):
            self.assertEqual(cli.cmd_arm("http://x", None, None, None, 120), 2)


class ObservationText(unittest.TestCase):
    def test_only_complete_zero_claims_zero(self):
        unknown = cli._describe_observation({"observation": "not_collected", "callEvidence": "none"})
        self.assertIn("未采集", unknown)

        partial = cli._describe_observation(
            {"observation": "partial", "callEvidence": "unknown", "missingActions": 2}
        )
        self.assertIn("未知", partial)

        zero_excel = cli._describe_observation(
            {"observation": "complete", "callEvidence": "zero", "attemptedActions": 2, "unit": "excel_sync_batches"}
        )
        self.assertIn("零刷新批次", zero_excel)

        zero_wps = cli._describe_observation(
            {"observation": "complete", "callEvidence": "zero", "attemptedActions": 1, "unit": "wps_api_calls"}
        )
        self.assertTrue(zero_wps.startswith("已观测到零写 API 调用"))


class EvidenceFailures(unittest.TestCase):
    def test_missing_ledger_is_insufficient(self):
        with tempfile.TemporaryDirectory() as directory, mock.patch.object(cli, "_health", return_value=True):
            with self.assertRaises(cli.ServiceError):
                cli.cmd_evidence("http://x", None, None, directory)


if __name__ == "__main__":
    unittest.main()

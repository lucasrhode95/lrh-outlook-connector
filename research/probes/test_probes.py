"""Synthetic offline checks: no authentication, real cache access or network."""

from __future__ import annotations

import base64
import json
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

_ISOLATED = tempfile.TemporaryDirectory()
os.environ["OUTLOOK_PROBE_HOME"] = _ISOLATED.name
_TEMPLATE = json.loads(
    (Path(__file__).resolve().parents[1] / "probe-config.example.json").read_text(
        encoding="utf-8"
    )
)
(Path(_ISOLATED.name) / "probe-config.json").write_text(
    json.dumps(_TEMPLATE), encoding="utf-8"
)

import common
import config


def token(**values):
    payload = {"tid": "tenant-a", "oid": "user-a", "upn": "user@example.com", **values}
    encoded = (
        base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip("=")
    )
    return "header." + encoded + ".signature"


def entry(**values):
    access = token(**values)
    return {
        "access_token": access,
        "account": common._account({"access_token": access}),
        "client_id": common.READ_CLIENT,
        "scope": common.PROFILES["read"][1],
        "expires_at": time.time() + 3600,
    }


class ProbeChecks(unittest.TestCase):
    def setUp(self):
        self.cache = self.enterContext(patch.object(common, "_load", return_value={}))
        self.enterContext(patch.object(common, "EXPECTED_USER", ""))
        self.network = self.enterContext(
            patch.object(
                common, "http", side_effect=AssertionError("Network forbidden")
            )
        )

    def test_expected_account_is_optional(self):
        self.assertEqual(common._checked("read", entry()), token())

    def test_optional_expected_account_is_enforced(self):
        with patch.object(common, "EXPECTED_USER", "other@example.com"):
            with self.assertRaisesRegex(SystemExit, "expected account"):
                common._checked("read", entry())

    def test_expected_account_comparison_is_case_insensitive(self):
        with patch.object(common, "EXPECTED_USER", "USER@EXAMPLE.COM"):
            self.assertEqual(common._checked("read", entry()), token())

    def test_mixed_account_cache_is_rejected(self):
        self.cache.return_value = {"write": entry(oid="another-user")}
        with self.assertRaisesRegex(SystemExit, "different accounts"):
            common._checked("read", entry())

    def test_same_identity_can_have_a_changed_email(self):
        self.cache.return_value = {"write": entry(upn="renamed@example.com")}
        self.assertEqual(common._checked("read", entry()), token())

    def test_missing_identity_is_rejected(self):
        with self.assertRaisesRegex(SystemExit, "no tenant/account identity"):
            common._checked("read", entry(tid=None))

    def test_wrong_account_is_not_saved_after_sign_in(self):
        self.cache.return_value = {"read": entry()}
        with patch.object(common, "_save") as save:
            with self.assertRaisesRegex(SystemExit, "different accounts"):
                common._store_result(
                    "write",
                    common.WRITE_CLIENT,
                    common.PROFILES["write"][1],
                    {"access_token": token(oid="another-user")},
                )
            save.assert_not_called()

    def test_changed_profile_does_not_reuse_cached_access_token(self):
        cached = entry()
        cached["client_id"] = "old-client"
        self.cache.return_value = {"read": cached}
        with self.assertRaises(common.SignInRequired):
            common.get_token("read")
        self.network.assert_not_called()

    def test_routing_uses_signed_in_email(self):
        with patch.object(common, "http", return_value=common.Resp(200)) as http:
            common.ows(
                token(upn="another@example.com"), "GetItem", "GetItemRequest", {}
            )
            self.assertEqual(
                http.call_args.kwargs["headers"]["X-AnchorMailbox"],
                "AAD-SMTP:another@example.com",
            )

    def test_routing_uses_identity_when_email_is_absent(self):
        with patch.object(common, "http", return_value=common.Resp(200)) as http:
            common.ows(token(upn=None), "GetItem", "GetItemRequest", {})
            self.assertEqual(
                http.call_args.kwargs["headers"]["X-AnchorMailbox"],
                "Oid:user-a@tenant-a",
            )

    def test_self_send_uses_token_email(self):
        body = common.ows_message(
            "synthetic", "body", "SaveOnly", token=token(upn="another@example.com")
        )
        self.assertEqual(
            body["Items"][0]["ToRecipients"][0]["EmailAddress"], "another@example.com"
        )

    def test_transport_failure_is_not_a_scope_denial(self):
        cached = entry()
        cached["refresh_token"] = "synthetic-refresh"
        self.cache.return_value = {"read": cached}
        with patch.object(
            common,
            "http",
            return_value=common.Resp(None, error_code="network:TimeoutError"),
        ):
            self.assertEqual(
                common.try_scope(
                    common.READ_CLIENT, "https://graph.microsoft.com/Mail.Send"
                )["result"],
                "inconclusive",
            )

    def test_opaque_refreshed_token_preserves_checked_sign_in_identity(self):
        cached = entry()
        cached["refresh_token"] = "synthetic-refresh"
        self.cache.return_value = {"read": cached}
        with patch.object(
            common,
            "http",
            return_value=common.Resp(200, payload={"access_token": "opaque-token"}),
        ):
            with patch.object(common, "_save") as save:
                fresh = common._refresh(
                    "read",
                    common.READ_CLIENT,
                    common.PROFILES["read"][1],
                    "synthetic-refresh",
                )
                self.assertEqual(fresh["account"], cached["account"])
                save.assert_called_once()

    def test_opaque_new_sign_in_without_identity_is_not_saved(self):
        with patch.object(common, "_save") as save:
            with self.assertRaisesRegex(SystemExit, "no tenant/account identity"):
                common._store_result(
                    "read",
                    common.READ_CLIENT,
                    common.PROFILES["read"][1],
                    {"access_token": "opaque-token"},
                )
            save.assert_not_called()

    def test_opaque_cached_token_routes_using_checked_sign_in_identity(self):
        cached = entry()
        cached["access_token"] = "opaque-token"
        self.cache.return_value = {"read": cached}
        with patch.object(common, "http", return_value=common.Resp(200)) as http:
            common.ows("opaque-token", "GetItem", "GetItemRequest", {})
            self.assertEqual(
                http.call_args.kwargs["headers"]["X-AnchorMailbox"],
                "AAD-SMTP:user@example.com",
            )

    def test_local_denial_is_skipped_without_a_request(self):
        with patch.object(common, "DENIED", {(common.READ_CLIENT, "*")}):
            self.assertEqual(
                common.try_scope(common.READ_CLIENT, "scope")["result"],
                "skipped_recorded_denial",
            )
        self.network.assert_not_called()


class ConfigurationChecks(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.addCleanup(self.folder.cleanup)
        self.enterContext(patch.object(config, "LOCAL", Path(self.folder.name)))

    def configure(self, values):
        (config.LOCAL / "probe-config.json").write_text(
            json.dumps(_TEMPLATE | values), encoding="utf-8"
        )

    def test_template_has_no_expected_account_or_historical_denials(self):
        self.configure({})
        settings = config.load_config()
        self.assertEqual(settings["expected_user"], "")
        self.assertEqual(settings["denied_pairs"], [])

    def test_configured_values_determine_endpoints_and_host_allowlist(self):
        self.configure(
            {
                "tenant": "tenant-a",
                "expected_user": "user@example.com",
                "graph_url": "https://graph.example.com/v1.0",
            }
        )
        settings = config.load_config()
        self.assertIn("/tenant-a/oauth2/v2.0", settings["authority"])
        self.assertIn("graph.example.com", settings["allowed_hosts"])
        self.assertEqual(settings["expected_user"], "user@example.com")

    def test_missing_file_requires_explicit_configuration(self):
        with self.assertRaisesRegex(SystemExit, "Missing probe configuration"):
            config.load_config()

    def test_missing_setting_is_not_filled_from_a_default(self):
        values = dict(_TEMPLATE)
        del values["graph_url"]
        (config.LOCAL / "probe-config.json").write_text(
            json.dumps(values), encoding="utf-8"
        )
        with self.assertRaisesRegex(
            SystemExit, "Missing probe configuration settings: graph_url"
        ):
            config.load_config()

    def test_unsafe_url_is_rejected(self):
        self.configure({"ows_url": "https://user:secret@example.com/owa"})
        with self.assertRaises(SystemExit):
            config.load_config()

    def test_unknown_configuration_key_is_rejected(self):
        self.configure({"expected_users": "typo"})
        with self.assertRaises(SystemExit):
            config.load_config()


if __name__ == "__main__":
    unittest.main()

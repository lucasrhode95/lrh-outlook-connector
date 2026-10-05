from __future__ import annotations

from pathlib import Path

import pytest
import requests

from outlook_connector import config
from outlook_connector.auth import tokens
from outlook_connector.auth.tokens import Account, TokenProvider
from outlook_connector.domain.errors import (
    AccountMismatch,
    AuthenticationRequired,
    ConfigurationError,
    SecureStorageUnavailable,
    Upstream,
)
from tests.fakes.msal_fakes import (
    OTHER_OID,
    USER_OID,
    Script,
    msal_account,
    seed_cache,
    token_result,
    use_script,
)

GRAPH = "https://graph.microsoft.com"
READ = config.OUTLOOK_MOBILE_CLIENT_ID
WRITE = config.ONE_OUTLOOK_WEB_CLIENT_ID


def provider(script: Script) -> TokenProvider:
    use_script(script)
    return TokenProvider(unsecure=True)


# ---------------------------------------------------------------- silent acquisition


def test_silent_token_from_cache() -> None:
    script = Script(accounts=[msal_account()], silent={READ: token_result(aud=GRAPH, scp="Mail.Read")})
    token = provider(script).get_token("read")
    assert token.profile == "read"
    assert token.source == "cache"
    assert token.claims()["aud"] == GRAPH
    assert "value=" not in repr(token)  # token values never appear in reprs or logs


def test_renewal_after_a_rejected_token_bypasses_the_memo() -> None:
    script = Script(accounts=[msal_account()], silent={READ: token_result(aud=GRAPH, scp="Mail.Read")})
    tokens_ = provider(script)
    tokens_.get_token("read")
    tokens_.get_token("read")  # served from memory
    tokens_.get_token("read", force_refresh=True)
    tokens_.get_token("read", claims_challenge='{"access_token":{}}')
    assert script.silent_options == [{}, {"force_refresh": True}, {"claims_challenge": '{"access_token":{}}'}]


def test_no_account_requires_sign_in_with_exact_command() -> None:
    with pytest.raises(AuthenticationRequired) as err:
        provider(Script()).get_token("read")
    assert err.value.command == "outlook-connector auth read --unsecure"


def test_rejected_refresh_token_requires_sign_in() -> None:
    script = Script(
        accounts=[msal_account()], silent={WRITE: {"error": "invalid_grant", "error_codes": [700082]}}
    )
    with pytest.raises(AuthenticationRequired) as err:
        provider(script).get_token("write")
    assert "AADSTS700082" in str(err.value)
    assert err.value.command.endswith("auth write --unsecure")


def test_no_refresh_token_for_this_client_requires_sign_in() -> None:
    # Signed in for 'read' only: MSAL finds the account but has nothing for the write client.
    script = Script(accounts=[msal_account()], silent={WRITE: None})
    with pytest.raises(AuthenticationRequired):
        provider(script).get_token("write")


def test_other_token_errors_are_upstream_errors() -> None:
    script = Script(accounts=[msal_account()], silent={READ: {"error": "temporarily_unavailable"}})
    with pytest.raises(Upstream):
        provider(script).get_token("read")


def test_network_failures_are_retried_then_succeed(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(tokens.time, "sleep", lambda _s: None)
    script = Script(
        accounts=[msal_account()],
        silent={READ: [requests.ConnectionError(), token_result(aud=GRAPH, scp="Mail.Read")]},
    )
    assert provider(script).get_token("read").profile == "read"


def test_persistent_network_failure_becomes_upstream(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(tokens.time, "sleep", lambda _s: None)
    script = Script(accounts=[msal_account()], silent={READ: [requests.Timeout()] * 3})
    with pytest.raises(Upstream):
        provider(script).get_token("read")


def test_two_accounts_in_cache_are_refused() -> None:
    script = Script(accounts=[msal_account(), msal_account(oid=OTHER_OID)])
    with pytest.raises(AccountMismatch):
        provider(script).get_token("read")


# ---------------------------------------------------------------- sign-in


def test_device_code_sign_in_shows_message_and_returns_token() -> None:
    shown: list[str] = []
    script = Script(device_result=token_result(aud=GRAPH, scp="Mail.Read", source="identity_provider"))
    token = provider(script).sign_in("read", show=shown.append)
    assert shown and "ABC123" in shown[0]
    assert token.source == "identity_provider"


def test_sign_in_with_a_different_account_is_undone_and_refused() -> None:
    script = Script(
        accounts=[msal_account(username="bound@example.com")],
        device_result=token_result(aud=GRAPH, scp="Mail.Read", oid=OTHER_OID, username="other@example.com"),
    )
    with pytest.raises(AccountMismatch, match="bound@example.com"):
        provider(script).sign_in("write", show=lambda _m: None)
    assert [a["home_account_id"].split(".")[0] for a in script.removed] == [OTHER_OID]
    assert [a["home_account_id"].split(".")[0] for a in script.accounts] == [USER_OID]


def test_expired_device_code_asks_to_sign_in_again() -> None:
    script = Script(device_result={"error": "authorization_pending", "error_codes": [70016]})
    with pytest.raises(AuthenticationRequired, match="expired before it was used"):
        provider(script).sign_in("read", show=lambda _m: None)


def test_declined_sign_in_is_an_upstream_error() -> None:
    script = Script(device_result={"error": "authorization_declined"})
    with pytest.raises(Upstream):
        provider(script).sign_in("read", show=lambda _m: None)


# ---------------------------------------------------------------- configuration and storage


def test_unknown_profile() -> None:
    with pytest.raises(ConfigurationError):
        provider(Script()).get_token("admin")


def test_encrypted_storage_unavailable_fails_closed(
    monkeypatch: pytest.MonkeyPatch, isolated_home: Path
) -> None:
    def unavailable(_path: str) -> object:
        raise RuntimeError("no DPAPI/keyring")

    monkeypatch.setattr(tokens, "build_encrypted_persistence", unavailable)
    secure = TokenProvider(unsecure=False)
    with pytest.raises(SecureStorageUnavailable):
        secure.get_token("read")
    assert not config.token_cache_path(unsecure=True).exists()  # never falls back to plaintext


def test_account_fingerprint_is_stable_and_not_the_username() -> None:
    a = Account.from_msal(msal_account(username="a@example.com"))
    b = Account.from_msal(msal_account(username="renamed@example.com"))
    assert a.fingerprint == b.fingerprint
    assert "example" not in a.fingerprint


# ---------------------------------------------------------------- offline status


def test_status_reports_profiles_from_a_real_cache_without_network() -> None:
    p = provider(Script())
    seed_cache(p._cache(), client_id=READ, scope="https://graph.microsoft.com/Mail.Read")
    status = p.status()
    assert status.exists and status.mode == "plaintext-development"
    assert [a.object_id for a in status.accounts] == [USER_OID]
    rows = {row.profile: row for row in status.profiles}
    assert rows["read"].signed_in and rows["read"].access_token_expires_on
    assert not rows["write"].signed_in and rows["write"].access_token_expires_on is None


def test_status_without_a_cache() -> None:
    status = provider(Script()).status()
    assert not status.exists and status.accounts == ()
    assert not any(row.signed_in for row in status.profiles)


def test_sign_out_deletes_the_cache() -> None:
    p = provider(Script())
    seed_cache(p._cache(), client_id=READ, scope="https://graph.microsoft.com/Mail.Read")
    assert p.sign_out() is True
    assert not p.cache_path.exists()
    assert p.sign_out() is False


# ---------------------------------------------------------------- per-process reuse


def test_app_is_built_once_and_valid_tokens_are_reused_from_memory() -> None:
    script = Script(accounts=[msal_account()], silent={READ: [token_result(aud=GRAPH, scp="Mail.Read")]})
    p = provider(script)
    first, second = p.get_token("read"), p.get_token("read")  # the second never reaches MSAL
    assert (first.source, second.source) == ("cache", "memory")
    assert script.created == [READ]


def test_tokens_close_to_expiry_are_refreshed(monkeypatch: pytest.MonkeyPatch) -> None:
    script = Script(accounts=[msal_account()], silent={READ: [token_result(aud=GRAPH, scp="Mail.Read")] * 2})
    p = provider(script)
    p.get_token("read")
    monkeypatch.setattr(tokens.time, "time", lambda: 10**12)  # far in the future: memo is stale
    assert p.get_token("read").source == "cache"
    assert script.silent[READ] == []  # both scripted silent results were consumed

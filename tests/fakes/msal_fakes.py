"""Fake MSAL pieces: a scriptable PublicClientApplication and synthetic cache entries.

Nothing here contacts Microsoft or contains real identifiers.
"""

from __future__ import annotations

import base64
import json
from dataclasses import dataclass, field
from typing import Any

TENANT = "00000000-0000-0000-0000-0000000000aa"
USER_OID = "11111111-1111-1111-1111-111111111111"
OTHER_OID = "22222222-2222-2222-2222-222222222222"


def b64(data: dict[str, Any]) -> str:
    return base64.urlsafe_b64encode(json.dumps(data).encode()).decode().rstrip("=")


def jwt(claims: dict[str, Any]) -> str:
    return f"{b64({'alg': 'none'})}.{b64(claims)}.sig"


def msal_account(
    oid: str = USER_OID, tid: str = TENANT, username: str = "user@example.com"
) -> dict[str, Any]:
    return {
        "home_account_id": f"{oid}.{tid}",
        "username": username,
        "environment": "login.microsoftonline.com",
    }


def token_result(
    *,
    aud: str,
    scp: str,
    oid: str = USER_OID,
    tid: str = TENANT,
    username: str = "user@example.com",
    source: str = "cache",
) -> dict[str, Any]:
    return {
        "access_token": jwt({"aud": aud, "scp": scp, "upn": username, "oid": oid, "tid": tid}),
        "expires_in": 3600,
        "token_source": source,
        "id_token_claims": {"oid": oid, "tid": tid, "preferred_username": username},
    }


@dataclass
class Script:
    """What the fake identity platform does. Shared by every FakeApp a test creates."""

    accounts: list[dict[str, Any]] = field(default_factory=list)
    silent: dict[str, Any] = field(default_factory=dict)  # client_id -> result (or Exception to raise)
    device_result: dict[str, Any] | None = None
    removed: list[dict[str, Any]] = field(default_factory=list)
    shown_flows: int = 0
    silent_options: list[dict[str, Any]] = field(default_factory=list)  # kwargs of each silent call
    created: list[str] = field(default_factory=list)


class FakeAppFactory:
    def __init__(self, script: Script) -> None:
        self.script = script

    def __call__(self, client_id: str, *, authority: str, token_cache: Any) -> FakeApp:
        self.script.created.append(client_id)
        return FakeApp(client_id, self.script)


class FakeApp:
    def __init__(self, client_id: str, script: Script) -> None:
        self.client_id = client_id
        self.script = script

    def get_accounts(self) -> list[dict[str, Any]]:
        return list(self.script.accounts)

    def acquire_token_silent_with_error(
        self, scopes: list[str], account: dict[str, Any], **options: Any
    ) -> Any:
        self.script.silent_options.append(options)
        result = self.script.silent.get(self.client_id)
        if isinstance(result, list):  # sequence of outcomes, consumed in order
            result = result.pop(0)
        if isinstance(result, Exception):
            raise result
        return result

    def initiate_device_flow(self, scopes: list[str]) -> dict[str, Any]:
        self.script.shown_flows += 1
        return {"user_code": "ABC123", "message": "Go to https://microsoft.com/devicelogin and enter ABC123"}

    def acquire_token_by_device_flow(self, flow: dict[str, Any]) -> dict[str, Any] | None:
        result = self.script.device_result
        if result and result.get("access_token"):
            claims = result["id_token_claims"]
            account = msal_account(claims["oid"], claims["tid"], claims["preferred_username"])
            if account not in self.script.accounts:
                self.script.accounts.append(account)
        return result

    def remove_account(self, account: dict[str, Any]) -> None:
        self.script.removed.append(account)
        self.script.accounts = [a for a in self.script.accounts if a != account]


def seed_cache(
    cache: Any,
    *,
    client_id: str,
    scope: str,
    refresh: bool = True,
    oid: str = USER_OID,
    tid: str = TENANT,
    username: str = "user@example.com",
) -> None:
    """Write a synthetic token response into a real MSAL cache (offline)."""
    response: dict[str, Any] = {
        "token_type": "Bearer",
        "access_token": jwt({"aud": scope.rsplit("/", 1)[0]}),
        "expires_in": 3600,
        "scope": scope,
        "client_info": b64({"uid": oid, "utid": tid}),
        "id_token": jwt(
            {
                "oid": oid,
                "tid": tid,
                "preferred_username": username,
                "aud": client_id,
                "iss": f"https://login.microsoftonline.com/{tid}/v2.0",
            }
        ),
    }
    if refresh:
        response["refresh_token"] = "synthetic-refresh-token"
    cache.add(
        {
            "client_id": client_id,
            "scope": scope.split(),
            "token_endpoint": "https://login.microsoftonline.com/organizations/oauth2/v2.0/token",
            "response": response,
            "params": {},
            "data": {},
        }
    )

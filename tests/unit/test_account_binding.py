from __future__ import annotations

import httpx
import pytest
from m365_auth import AccessToken, AccountMismatch

from outlook_connector.bootstrap import AppContext
from tests.fakes.graph_fake import FakeGraph, sample_mailbox
from tests.fakes.msal_fakes import TENANT, USER_OID, jwt


class MutableTokens:
    def __init__(self) -> None:
        self.object_id = USER_OID

    def get_token(self, profile: str, **_: object) -> AccessToken:
        value = jwt({"tid": TENANT, "oid": self.object_id, "upn": "user@example.test", "aud": profile})
        return AccessToken(profile=profile, value=value, source="cache", expires_on=None)

    def sign_in_command(self, profile: str | None = None) -> str:
        return "outlook-connector auth"


@pytest.mark.asyncio
async def test_account_switch_rebuilds_services_and_rejects_stale_requests() -> None:
    fake: FakeGraph = sample_mailbox()
    tokens = MutableTokens()
    client = httpx.AsyncClient(transport=fake.transport())
    context = AppContext(tokens=tokens, http_client=client)  # type: ignore[arg-type]

    first = await context.services()
    tokens.object_id = "22222222-2222-2222-2222-222222222222"
    second = await context.services()

    assert first.account.fingerprint != second.account.fingerprint
    await second.mailbox.folders(refresh=True)
    calls_before_stale_request = len(fake.calls)

    with pytest.raises(AccountMismatch, match="changed"):
        await first.mailbox.folders(refresh=True)

    assert len(fake.calls) == calls_before_stale_request
    await context.aclose()

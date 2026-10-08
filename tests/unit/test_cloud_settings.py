from __future__ import annotations

import httpx
import pytest

from outlook_connector.domain.errors import AccountMismatch, Upstream
from outlook_connector.remote.cloud_settings import CloudSettings
from outlook_connector.remote.transport import Transport
from tests.fakes.graph_fake import StaticTokens


@pytest.mark.parametrize(
    "records, names, new_default",
    [
        ([], (), None),
        ([{"name": "roaming_new_signature", "value": "", "scope": "test-scope"}], (), None),
        ([{"name": "roaming_signature_list", "value": "Logo", "scope": "test-scope"}], ("Logo",), None),
        (
            [
                {"name": "roaming_signature_list", "value": "Logo", "scope": "test-scope"},
                {"name": "roaming_new_signature", "value": "Logo", "scope": "test-scope"},
            ],
            ("Logo",),
            "Logo",
        ),
    ],
)
async def test_absent_records_have_empty_defaults_without_discarding_present_records(
    records: list[dict[str, str]], names: tuple[str, ...], new_default: str | None
) -> None:
    tokens = StaticTokens()
    async with httpx.AsyncClient(
        transport=httpx.MockTransport(lambda _: httpx.Response(200, json=records))
    ) as client:
        settings = await CloudSettings(Transport(tokens, client=client), tokens).settings()
    assert settings.names == names and settings.new_default == new_default
    assert settings.reply_default is None
    assert settings.scope == ("test-scope" if records else None)


@pytest.mark.parametrize(
    "data, error",
    [
        ({"value": {}}, Upstream),
        ([{}], Upstream),
        ([{"name": "roaming_signature_list", "value": ""}], Upstream),
        ([{"name": "roaming_signature_list", "value": [], "scope": "test-scope"}], Upstream),
        ([{"name": "roaming_new_signature", "value": 42, "scope": "test-scope"}], Upstream),
        (
            [
                {"name": "roaming_signature_list", "value": "", "scope": "test-scope"},
                {"name": "roaming_signature_list", "value": "", "scope": "test-scope"},
            ],
            Upstream,
        ),
        (
            [
                {"name": "roaming_new_signature", "value": "", "scope": "test-scope"},
                {"name": "roaming_reply_signature", "value": "", "scope": "different-scope"},
            ],
            AccountMismatch,
        ),
    ],
)
async def test_missing_records_do_not_hide_malformed_or_mismatched_settings(
    data: object, error: type[Exception]
) -> None:
    tokens = StaticTokens()
    async with httpx.AsyncClient(
        transport=httpx.MockTransport(lambda _: httpx.Response(200, json=data))
    ) as client:
        with pytest.raises(error):
            await CloudSettings(Transport(tokens, client=client), tokens).settings()

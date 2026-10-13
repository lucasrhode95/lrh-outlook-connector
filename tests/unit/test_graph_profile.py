from __future__ import annotations

from typing import Any

import httpx

from outlook_connector.remote.graph import Graph
from outlook_connector.remote.transport import Transport
from tests.fakes.graph_fake import StaticTokens


class RecordingTokens(StaticTokens):
    def __init__(self) -> None:
        super().__init__()
        self.profiles: list[str] = []

    def get_token(self, profile: str, **renewal: Any) -> Any:
        self.profiles.append(profile)
        return super().get_token(profile, **renewal)


async def _no_sleep(_: float) -> None:
    return None


async def test_graph_client_sends_each_call_with_the_profile_it_was_built_with() -> None:
    tokens = RecordingTokens()
    client = httpx.AsyncClient(
        transport=httpx.MockTransport(lambda request: httpx.Response(200, json={"value": []}))
    )
    transport = Transport(tokens, client=client, sleep=_no_sleep)

    await Graph(transport, profile="teams.graph").get("/me")
    await Graph(transport).get("/me")

    assert tokens.profiles == ["teams.graph", "outlook.graph"]

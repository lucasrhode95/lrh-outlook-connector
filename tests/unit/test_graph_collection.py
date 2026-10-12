from __future__ import annotations

from typing import Any, cast

import pytest

from outlook_connector.domain.errors import Upstream
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.transport import Transport

GRAPH = "https://graph.microsoft.com/v1.0"


class Pages:
    def __init__(self, responses: list[dict[str, Any]]) -> None:
        self.responses = iter(responses)
        self.urls: list[str] = []

    async def json(self, method: str, url: str, **_: Any) -> dict[str, Any]:
        assert method == "GET"
        self.urls.append(url)
        return next(self.responses)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "continuation",
    [
        [f"{GRAPH}/items?skip=1", f"{GRAPH}/items?skip=1"],
        [f"{GRAPH}/items?skip=1", f"{GRAPH}/items?skip=2", f"{GRAPH}/items?skip=1"],
    ],
)
async def test_collect_rejects_repeated_continuation_links(continuation: list[str]) -> None:
    pages = Pages(
        [
            {"value": [], "@odata.nextLink": continuation[0]},
            *({"value": [], "@odata.nextLink": link} for link in continuation[1:]),
        ]
    )
    graph = Graph(cast(Transport, pages))

    with pytest.raises(Upstream, match="repeated a continuation link"):
        await graph.collect("/items", max_items=10)

    assert len(pages.urls) == len(continuation)


@pytest.mark.asyncio
async def test_collect_keeps_opaque_paging_and_item_bound() -> None:
    pages = Pages(
        [
            {"value": [{"id": "first"}], "@odata.nextLink": f"{GRAPH}/items?skip=1"},
            {"value": [{"id": "second"}], "@odata.nextLink": f"{GRAPH}/items?skip=2"},
        ]
    )
    graph = Graph(cast(Transport, pages))

    items, truncated = await graph.collect("/items", max_items=2)

    assert [item["id"] for item in items] == ["first", "second"]
    assert truncated
    assert len(pages.urls) == 2

from __future__ import annotations

import json

import httpx

from outlook_connector.remote.substrate_events import SEARCH_URL, SubstrateEventSearch
from outlook_connector.remote.transport import Transport
from tests.fakes.graph_fake import StaticTokens


async def _no_sleep(_: float) -> None:
    return None


async def test_search_v2_event_request_and_mapping() -> None:
    requests: list[httpx.Request] = []

    async def handle(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return httpx.Response(
            200,
            json={
                "EntitySets": [
                    {
                        "ResultSets": [
                            {
                                "Results": [
                                    {
                                        "Source": {
                                            "ItemId": {"Id": "synthetic-owa-event-id"},
                                            "ImmutableId": "synthetic-immutable-id",
                                            "Subject": "Synthetic review",
                                            "Start": "2026-10-10T12:00:00Z",
                                            "End": "2026-10-10T13:00:00Z",
                                            "Organizer": {
                                                "Name": "Synthetic Organizer",
                                                "EmailAddress": "organizer@example.invalid",
                                            },
                                            "Location": "Synthetic room",
                                            "IsMeeting": True,
                                            "CalendarItemType": "Occurrence",
                                            "SeriesMasterId": "synthetic-series-id",
                                        }
                                    }
                                ],
                                "MoreResultsAvailable": True,
                            }
                        ]
                    }
                ]
            },
        )

    tokens = StaticTokens()
    client = httpx.AsyncClient(transport=httpx.MockTransport(handle))
    transport = Transport(tokens, client=client, sleep=_no_sleep)
    try:
        adapter = SubstrateEventSearch(transport, anchor_mailbox="Oid:user-x@tenant-x")
        page = await adapter.search_events(
            "synthetic review",
            offset=5,
            page_size=5,
            cvid="synthetic-cvid",
            logical_id="synthetic-logical-id",
        )
    finally:
        await transport.aclose()

    assert tokens.renewals == []
    assert len(requests) == 1
    request = requests[0]
    assert str(request.url) == SEARCH_URL
    assert request.headers["x-anchormailbox"] == "Oid:user-x@tenant-x"
    body = json.loads(request.content)
    entity = body["EntityRequests"][0]
    assert entity["EntityType"] == "Event" and entity["ContentSources"] == ["Exchange"]
    assert entity["From"] == 5 and entity["Size"] == 5
    assert body["Cvid"] == "synthetic-cvid" and body["LogicalId"] == "synthetic-logical-id"
    assert page.result_set_found and page.more_results_available is True and page.rows_received == 1
    event = page.events[0]
    assert event.id == "synthetic-owa-event-id"
    assert event.start and event.start.isoformat() == "2026-10-10T12:00:00+00:00"
    assert event.organizer and event.organizer.name == "Synthetic Organizer"
    assert event.location == "Synthetic room" and event.is_meeting is True
    assert event.calendar_item_type == "Occurrence" and event.series_master_id == "synthetic-series-id"

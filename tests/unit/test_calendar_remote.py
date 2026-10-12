from __future__ import annotations

import json
from urllib.parse import unquote

import httpx

from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_calendar import OwsCalendarItems
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


async def test_ows_event_hydration_uses_retryable_read_and_parses_item() -> None:
    attempts = 0
    request_bodies: list[dict[str, object]] = []

    async def handle(request: httpx.Request) -> httpx.Response:
        nonlocal attempts
        attempts += 1
        request_bodies.append(json.loads(unquote(request.headers["x-owa-urlpostdata"])))
        if attempts == 1:
            return httpx.Response(503, json={"error": {"code": "synthetic"}})
        return httpx.Response(
            200,
            json={
                "Body": {
                    "ResponseMessages": {
                        "Items": [
                            {
                                "ResponseClass": "Success",
                                "ResponseCode": "NoError",
                                "Items": [
                                    {
                                        "Subject": "Synthetic review",
                                        "Start": "2026-10-10T12:00:00Z",
                                        "End": "2026-10-10T13:00:00Z",
                                        "Organizer": {
                                            "Mailbox": {
                                                "Name": "Synthetic Organizer",
                                                "EmailAddress": "organizer@example.invalid",
                                            }
                                        },
                                        "RequiredAttendees": [
                                            {
                                                "Mailbox": {
                                                    "Name": "Synthetic Attendee",
                                                    "EmailAddress": "attendee@example.invalid",
                                                }
                                            }
                                        ],
                                        "Body": {
                                            "BodyType": "Text",
                                            "IsTruncated": True,
                                            "Value": (
                                                "Agenda &lt;draft&gt;; AT&amp;T. Join "
                                                "https://teams.microsoft.com/l/meetup-join/synthetic-link."
                                            ),
                                        },
                                        "Attachments": [
                                            {
                                                "AttachmentId": {"Id": "synthetic-attachment-id"},
                                                "Name": "synthetic.txt",
                                                "ContentType": "text/plain",
                                                "Size": 12,
                                                "IsInline": False,
                                            }
                                        ],
                                    }
                                ],
                            }
                        ]
                    }
                }
            },
        )

    tokens = StaticTokens()
    client = httpx.AsyncClient(transport=httpx.MockTransport(handle))
    transport = Transport(tokens, client=client, sleep=_no_sleep)
    try:
        event = await OwsCalendarItems(Ows(transport, tokens)).get_event("synthetic-owa-event-id")
    finally:
        await transport.aclose()

    assert attempts == 2
    assert [body["Body"]["ItemIds"][0]["Id"] for body in request_bodies] == [
        "synthetic-owa-event-id",
        "synthetic-owa-event-id",
    ]
    shape = request_bodies[0]["Body"]["ItemShape"]
    assert shape["BaseShape"] == "Default" and shape["BodyType"] == "Text"
    assert event.subject == "Synthetic review" and event.organizer is not None
    assert event.required_attendees[0].name == "Synthetic Attendee"
    assert event.body_text == (
        "Agenda &lt;draft&gt;; AT&amp;T. Join https://teams.microsoft.com/l/meetup-join/synthetic-link."
    )
    assert event.body_truncated is True
    assert event.teams_join_url == "https://teams.microsoft.com/l/meetup-join/synthetic-link"
    assert event.attachments[0].name == "synthetic.txt"

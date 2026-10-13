from __future__ import annotations

import json
from datetime import UTC, datetime
from typing import Any
from urllib.parse import unquote

import httpx
import pytest

from outlook_connector.domain.errors import Upstream
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_calendar import OwsCalendarWindow
from outlook_connector.remote.transport import Transport
from tests.fakes.graph_fake import StaticTokens

SINCE = datetime(2026, 10, 12, 3, tzinfo=UTC)
UNTIL = datetime(2026, 10, 19, 3, tzinfo=UTC)
ANSWER_OK: dict[str, Any] = {
    "Header": {},
    "Body": {"Items": [], "ResponseCode": "NoError", "ResponseClass": "Success"},
}


async def _no_sleep(_: float) -> None:
    return None


def _window(handler: Any) -> OwsCalendarWindow:
    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    transport = Transport(StaticTokens(), client=client, sleep=_no_sleep)
    return OwsCalendarWindow(Ows(transport, StaticTokens()))


def _payload(request: httpx.Request) -> dict[str, Any]:
    # Small payloads travel URL-encoded in the X-OWA-UrlPostData header (research §4.1).
    return json.loads(unquote(request.headers["X-OWA-UrlPostData"]))


async def test_get_calendar_view_asks_for_the_default_calendar_range_in_utc() -> None:
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return httpx.Response(200, json=ANSWER_OK)

    await _window(handler).list_events(since=SINCE, until=UNTIL)

    request = requests[0]
    assert request.headers["Action"] == "GetCalendarView"
    payload = _payload(request)
    assert payload["__type"] == "GetCalendarViewJsonRequest:#Exchange"
    zone = payload["Header"]["TimeZoneContext"]["TimeZoneDefinition"]
    assert zone["Id"] == "UTC"
    body = payload["Body"]
    assert body["__type"] == "GetCalendarViewRequest:#Exchange"
    assert body["RangeStart"] == "2026-10-12T03:00:00.000"
    assert body["RangeEnd"] == "2026-10-19T03:00:00.000"
    assert body["CalendarId"]["BaseFolderId"] == {
        "__type": "DistinguishedFolderId:#Exchange",
        "Id": "calendar",
    }


async def test_items_come_from_the_body_and_keep_their_item_ids() -> None:
    item = {
        "__type": "CalendarItem:#Exchange",
        "ItemId": {"Id": "synthetic-ows-id"},
        "Subject": "Synthetic review",
        "Start": "2026-10-13T16:30:00Z",
        "End": "2026-10-13T17:00:00Z",
        "Organizer": {
            "Mailbox": {"Name": "Synthetic Organizer", "EmailAddress": "organizer@example.invalid"}
        },
        "RequiredAttendees": [{"Mailbox": {"Name": "Synthetic Two", "EmailAddress": "two@example.invalid"}}],
        "IsCancelled": False,
        "CalendarItemType": "Occurrence",
    }
    answer = {"Header": {}, "Body": {"Items": [item], "ResponseCode": "NoError", "ResponseClass": "Success"}}
    window = _window(lambda request: httpx.Response(200, json=answer))

    (event,) = await window.list_events(since=SINCE, until=UNTIL)

    assert event.id == "synthetic-ows-id" and event.immutable_id == "synthetic-ows-id"
    assert event.subject == "Synthetic review"
    assert event.start == datetime(2026, 10, 13, 16, 30, tzinfo=UTC)
    assert event.calendar_item_type == "Occurrence" and event.is_cancelled is False
    assert [person.address for person in event.required_attendees] == ["two@example.invalid"]
    assert event.body_text is None and event.teams_join_url is None


async def test_a_refused_answer_is_an_upstream_error() -> None:
    answer = {
        "Header": {},
        "Body": {"Items": [], "ResponseCode": "ErrorInternalServerError", "ResponseClass": "Error"},
    }
    window = _window(lambda request: httpx.Response(200, json=answer))

    with pytest.raises(Upstream):
        await window.list_events(since=SINCE, until=UNTIL)

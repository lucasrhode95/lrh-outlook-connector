from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

import httpx
import pytest

from outlook_connector import config
from outlook_connector.domain.errors import InvalidRequest, Upstream
from outlook_connector.remote import graph_calendar, ids
from outlook_connector.remote.calendar_mapping import teams_chat_id_from_join_link
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_calendar import GraphCalendarItems, GraphCalendarView
from outlook_connector.remote.transport import Transport
from tests.fakes.graph_fake import StaticTokens

SINCE = datetime(2026, 10, 12, 3, tzinfo=UTC)
UNTIL = datetime(2026, 10, 19, 3, tzinfo=UTC)
NEXT_LINK = (
    "https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=2026-10-12T03%3A00%3A00Z"
    "&endDateTime=2026-10-19T03%3A00%3A00Z&$top=10&$skiptoken=synthetic"
)


class RecordingTokens(StaticTokens):
    def __init__(self) -> None:
        super().__init__()
        self.profiles: list[str] = []

    def get_token(self, profile: str, **renewal: Any) -> Any:
        self.profiles.append(profile)
        return super().get_token(profile, **renewal)


async def _no_sleep(_: float) -> None:
    return None


def _view(handler: Any) -> tuple[GraphCalendarView, RecordingTokens]:
    tokens = RecordingTokens()
    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    transport = Transport(tokens, client=client, sleep=_no_sleep)
    return GraphCalendarView(Graph(transport, profile=config.CALENDAR_PROFILE)), tokens


def _items(handler: Any) -> tuple[GraphCalendarItems, RecordingTokens]:
    tokens = RecordingTokens()
    client = httpx.AsyncClient(transport=httpx.MockTransport(handler))
    transport = Transport(tokens, client=client, sleep=_no_sleep)
    return GraphCalendarItems(Graph(transport, profile=config.CALENDAR_PROFILE)), tokens


def _synthetic_event(**overrides: Any) -> dict[str, Any]:
    event: dict[str, Any] = {
        "id": "synthetic-id_x",
        "subject": "Synthetic planning",
        "start": {"dateTime": "2026-10-13T16:30:00.0000000", "timeZone": "UTC"},
        "end": {"dateTime": "2026-10-13T17:00:00.0000000", "timeZone": "UTC"},
        "isAllDay": False,
        "isCancelled": False,
        "isOrganizer": False,
        "location": {"displayName": "Synthetic room"},
        "organizer": {
            "emailAddress": {"name": "Synthetic Organizer", "address": "organizer@example.invalid"}
        },
        "attendees": [
            {
                "type": "required",
                "emailAddress": {"name": "Synthetic Two", "address": "two@example.invalid"},
                "status": {"response": "accepted"},
            },
            {
                "type": "optional",
                "emailAddress": {"name": "Synthetic Three", "address": "three@example.invalid"},
                "status": {"response": "none"},
            },
        ],
        "responseStatus": {"response": "tentativelyAccepted"},
        "type": "occurrence",
        "seriesMasterId": "synthetic-series-id",
        "onlineMeeting": {
            "joinUrl": "https://teams.microsoft.com/l/meetup-join/19:meeting_synthetic@thread.v2/0?context=x"
        },
    }
    event.update(overrides)
    return event


def _json_page(events: list[dict[str, Any]], link: str | None = None) -> httpx.Response:
    body: dict[str, Any] = {"value": events}
    if link is not None:
        body["@odata.nextLink"] = link
    return httpx.Response(200, json=body)


async def test_calendar_view_request_asks_for_utc_text_and_the_listed_fields() -> None:
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return _json_page([_synthetic_event()])

    view, tokens = _view(handler)
    page = await view.list_events(since=SINCE, until=UNTIL, page_size=25, page=None, include_body=False)

    request = requests[0]
    assert request.url.path == "/v1.0/me/calendarView"
    assert request.url.params["startDateTime"] == "2026-10-12T03:00:00Z"
    assert request.url.params["endDateTime"] == "2026-10-19T03:00:00Z"
    assert request.url.params["$top"] == "25"
    assert request.url.params["$orderby"] == "start/dateTime asc"  # pages then follow start order
    fields = request.url.params["$select"].split(",")
    assert "onlineMeeting" in fields and "attendees" in fields and "body" not in fields
    prefer = request.headers["prefer"]
    assert 'outlook.timezone="UTC"' in prefer and 'outlook.body-content-type="text"' in prefer
    assert 'IdType="ImmutableId"' in prefer
    assert tokens.profiles == [config.CALENDAR_PROFILE]
    assert page.next_link is None and len(page.events) == 1


async def test_secondary_calendar_view_uses_its_id_in_the_graph_route_and_cursor() -> None:
    requests: list[httpx.Request] = []
    next_link = (
        # Graph's own next-link form for a calendar, with the id unescaped (live 2026-11-03)
        "https://graph.microsoft.com/v1.0/me/calendars('synthetic-calendar_x==')/calendarView?"
        "startDateTime=2026-10-12T03%3A00%3A00Z&endDateTime=2026-10-19T03%3A00%3A00Z&$top=10"
        "&$skiptoken=synthetic"
    )

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return _json_page([_synthetic_event()], link=next_link)

    view, _ = _view(handler)
    first = await view.list_events(
        since=SINCE,
        until=UNTIL,
        page_size=10,
        page=None,
        include_body=False,
        calendar_id="synthetic-calendar_x==",
    )
    second = await view.list_events(
        since=SINCE,
        until=UNTIL,
        page_size=10,
        page=first.next_link,
        include_body=False,
        calendar_id="synthetic-calendar_x==",
    )

    assert requests[0].url.raw_path.startswith(b"/v1.0/me/calendars/synthetic-calendar_x%3D%3D/calendarView")
    assert requests[1].url.path == "/v1.0/me/calendars('synthetic-calendar_x==')/calendarView"
    assert second.next_link == next_link
    with pytest.raises(InvalidRequest):
        await view.list_events(
            since=SINCE,
            until=UNTIL,
            page_size=10,
            page=next_link,
            include_body=False,
            calendar_id="different-calendar",
        )
    assert len(requests) == 2
    plain = next_link.replace("calendars('synthetic-calendar_x==')", "calendars/synthetic-calendar_x%3D%3D")
    for link in (next_link, plain):
        assert graph_calendar._continues_window(link, SINCE, UNTIL, "synthetic-calendar_x==")
    assert not graph_calendar._continues_window(next_link, SINCE, UNTIL, None)


async def test_calendar_catalog_lists_id_name_and_default_status() -> None:
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return httpx.Response(
            200,
            json={
                "value": [
                    {"id": "synthetic-default", "name": "Calendar", "isDefaultCalendar": True},
                    {"id": "synthetic-secondary", "name": "Planning", "isDefaultCalendar": False},
                ]
            },
        )

    view, tokens = _view(handler)
    calendars = await view.list_calendars()

    assert requests[0].url.path == "/v1.0/me/calendars"
    assert requests[0].url.params["$select"] == "id,name,isDefaultCalendar"
    assert tokens.profiles == [config.CALENDAR_PROFILE]
    assert [(calendar.id, calendar.name, calendar.is_default) for calendar in calendars] == [
        ("synthetic-default", "Calendar", True),
        ("synthetic-secondary", "Planning", False),
    ]


async def test_descriptions_are_selected_only_when_asked_for() -> None:
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return _json_page([])

    view, _ = _view(handler)
    await view.list_events(since=SINCE, until=UNTIL, page_size=5, page=None, include_body=True)

    assert "body" in requests[0].url.params["$select"].split(",")


async def test_events_map_to_the_connector_model_in_utc_with_outlook_ids_and_the_teams_link() -> None:
    view, _ = _view(lambda request: _json_page([_synthetic_event()]))

    (event,) = (
        await view.list_events(since=SINCE, until=UNTIL, page_size=5, page=None, include_body=False)
    ).events

    assert event.start == datetime(2026, 10, 13, 16, 30, tzinfo=UTC)
    assert event.end == datetime(2026, 10, 13, 17, 0, tzinfo=UTC)
    assert event.immutable_id == "synthetic-id_x"
    assert event.id == ids.to_ows("synthetic-id_x")
    assert event.series_master_id == ids.to_ows("synthetic-series-id")
    assert [person.address for person in event.required_attendees] == ["two@example.invalid"]
    assert [person.address for person in event.optional_attendees] == ["three@example.invalid"]
    assert event.organizer is not None and event.organizer.address == "organizer@example.invalid"
    assert event.my_response_type == "Tentative"
    assert event.calendar_item_type == "Occurrence"
    assert event.is_meeting is True and event.is_cancelled is False
    assert event.location == "Synthetic room"
    assert (
        event.teams_join_url
        == "https://teams.microsoft.com/l/meetup-join/19:meeting_synthetic@thread.v2/0?context=x"
    )
    assert event.teams_chat_id_from_join_link == "19:meeting_synthetic@thread.v2"
    assert event.body_text is None


async def test_the_teams_link_is_read_from_the_description_when_the_meeting_block_is_absent() -> None:
    link = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_synthetic%40thread.v2/0?context=x"
    description = f"Agenda.\nJoin: {link}."
    event = _synthetic_event(onlineMeeting=None, body={"contentType": "text", "content": description})
    view, _ = _view(lambda request: _json_page([event]))

    (listed,) = (
        await view.list_events(since=SINCE, until=UNTIL, page_size=5, page=None, include_body=True)
    ).events

    assert listed.teams_join_url == link and listed.body_text == description
    assert listed.teams_chat_id_from_join_link is None  # only Graph's meeting link names the chat


async def test_continuation_links_are_followed_with_the_same_preferences() -> None:
    seen: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        seen.append(request)
        if len(seen) == 1:
            return _json_page([_synthetic_event()], link=NEXT_LINK)
        return _json_page([])

    view, _ = _view(handler)
    first = await view.list_events(since=SINCE, until=UNTIL, page_size=10, page=None, include_body=True)
    assert first.next_link == NEXT_LINK

    second = await view.list_events(
        since=SINCE, until=UNTIL, page_size=10, page=first.next_link, include_body=True
    )

    assert second.next_link is None and second.events == []
    assert seen[1].url.params["$skiptoken"] == "synthetic"
    assert 'outlook.timezone="UTC"' in seen[1].headers["prefer"]


async def test_a_continuation_link_must_be_this_calendar_view() -> None:
    view, _ = _view(lambda request: _json_page([]))

    with pytest.raises(InvalidRequest):
        await view.list_events(
            since=SINCE,
            until=UNTIL,
            page_size=5,
            page="https://graph.microsoft.com/v1.0/me/messages?$skiptoken=synthetic",
            include_body=False,
        )


async def test_a_continuation_link_must_be_for_this_window_and_a_bounded_page() -> None:
    reads: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        reads.append(request)
        return _json_page([])

    view, _ = _view(handler)
    base = "https://graph.microsoft.com/v1.0/me/calendarView?"
    other_window = (
        base + "startDateTime=2026-01-01T00%3A00%3A00Z&endDateTime=2026-10-19T03%3A00%3A00Z&$top=10"
    )
    too_big = base + "startDateTime=2026-10-12T03%3A00%3A00Z&endDateTime=2026-10-19T03%3A00%3A00Z&$top=1000"
    unbounded = base + "startDateTime=2026-10-12T03%3A00%3A00Z&endDateTime=2026-10-19T03%3A00%3A00Z"

    for link in (other_window, too_big, unbounded):
        with pytest.raises(InvalidRequest):
            await view.list_events(since=SINCE, until=UNTIL, page_size=5, page=link, include_body=False)
    assert reads == []  # a refused link is never sent to Graph


async def test_a_time_outside_utc_is_refused_rather_than_guessed() -> None:
    event = _synthetic_event(
        start={"dateTime": "2026-10-13T13:30:00.0000000", "timeZone": "E. South America Standard Time"}
    )
    view, _ = _view(lambda request: _json_page([event]))

    with pytest.raises(Upstream):
        await view.list_events(since=SINCE, until=UNTIL, page_size=5, page=None, include_body=False)


async def test_a_refused_calendar_read_is_an_upstream_error_the_listing_can_fall_back_from() -> None:
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(403, json={"error": {"code": "ErrorAccessDenied", "message": "synthetic"}})

    view, _ = _view(handler)

    with pytest.raises(Upstream):
        await view.list_events(since=SINCE, until=UNTIL, page_size=5, page=None, include_body=False)


async def test_event_hydration_uses_graph_and_lists_attachments_even_when_flag_is_false() -> None:
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        if request.url.path.endswith("/attachments"):
            return httpx.Response(
                200,
                json={
                    "value": [
                        {
                            "id": "synthetic-attachment",
                            "name": "agenda.txt",
                            "contentType": "text/plain",
                            "size": 41,
                            "isInline": False,
                        }
                    ]
                },
            )
        return httpx.Response(
            200,
            json=_synthetic_event(
                hasAttachments=False,
                body={"contentType": "text", "content": "Synthetic agenda"},
                responseStatus={"response": "accepted"},
                responseRequested=True,
                sensitivity="private",
                iCalUId="synthetic-uid",
            ),
        )

    items, tokens = _items(handler)
    event = await items.get_event(ids.to_ows("synthetic-id_x"))

    assert [request.url.path for request in requests] == [
        "/v1.0/me/events/synthetic-id_x",
        "/v1.0/me/events/synthetic-id_x/attachments",
    ]
    assert all('IdType="ImmutableId"' in request.headers["prefer"] for request in requests)
    assert requests[0].headers["prefer"].find('outlook.body-content-type="text"') >= 0
    assert 'outlook.timezone="UTC"' in requests[0].headers["prefer"]
    fields = requests[0].url.params["$select"].split(",")
    assert {"body", "onlineMeeting", "responseRequested", "sensitivity", "iCalUId"} <= set(fields)
    assert requests[1].url.params["$select"] == "id,name,contentType,size,isInline"
    assert tokens.profiles == [config.CALENDAR_PROFILE, config.CALENDAR_PROFILE]
    assert event.id == ids.to_ows("synthetic-id_x")
    assert event.my_response_type == "Accept" and event.body_text == "Synthetic agenda"
    assert event.attachments[0].name == "agenda.txt" and event.attachments[0].size == 41
    assert event.is_response_requested is True and event.sensitivity == "Private"
    assert event.uid == "synthetic-uid"


async def test_event_attachments_are_listed_alone_and_a_cut_or_malformed_listing_is_refused(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    answers = iter(
        [
            {"value": [{"name": "a.txt"}, {"id": "synthetic-b", "name": "b.txt"}]},
            {"value": [{"id": "synthetic-a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"},
        ]
    )
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return httpx.Response(200, json=next(answers))

    items, _ = _items(handler)
    with pytest.raises(Upstream, match="without an id"):
        await items.list_attachments(ids.to_ows("synthetic-id_x"))
    assert [request.url.path for request in requests] == ["/v1.0/me/events/synthetic-id_x/attachments"]

    monkeypatch.setattr(graph_calendar, "MAX_EVENT_ATTACHMENTS", 1)
    with pytest.raises(Upstream, match="more than 1 attachments"):
        await items.list_attachments(ids.to_ows("synthetic-id_x"))


async def test_event_attachment_download_uses_the_graph_event_route(tmp_path: Any) -> None:
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        return httpx.Response(200, content=b"synthetic-file")

    items, tokens = _items(handler)
    size = await items.download_attachment("synthetic-id/x", "synthetic-attachment", tmp_path / "event.txt")

    assert size == len(b"synthetic-file")
    assert requests[0].url.path == "/v1.0/me/events/synthetic-id-x/attachments/synthetic-attachment/$value"
    assert tokens.profiles == [config.CALENDAR_PROFILE]


def test_join_link_chat_id_parsing_is_limited_to_the_meeting_path_segment() -> None:
    assert (
        teams_chat_id_from_join_link(
            "https://teams.microsoft.com/l/meetup-join/19%3Ameeting_synthetic%40thread.v2/0?context=x"
        )
        == "19:meeting_synthetic@thread.v2"
    )
    assert teams_chat_id_from_join_link("https://teams.microsoft.com/l/meetup-join/other/0") is None
    assert teams_chat_id_from_join_link("https://teams.microsoft.com/l/meetup-join/[broken") is None

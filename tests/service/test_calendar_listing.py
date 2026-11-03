from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

import pytest

from outlook_connector.domain.errors import (
    AccountMismatch,
    AuthenticationRequired,
    InvalidRequest,
    Throttled,
    Upstream,
)
from outlook_connector.domain.models import CalendarEvent, CalendarInfo, Recipient
from outlook_connector.remote.results import CalendarViewPage
from outlook_connector.service import cursors
from outlook_connector.service.calendar_listing import (
    CURSOR_KIND,
    MAX_GRAPH_PAGES,
    MAX_LISTED_ATTENDEES,
    MAX_LISTED_BODY_CHARS,
    CalendarListing,
)

SINCE = datetime(2026, 10, 12, 3, tzinfo=UTC)
UNTIL = datetime(2026, 10, 19, 3, tzinfo=UTC)
LINK_2 = "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=synthetic-page-2"


def _event(
    event_id: str,
    hour: int,
    *,
    subject: str = "Synthetic meeting",
    body: str | None = None,
    required: int = 0,
    optional: int = 0,
) -> CalendarEvent:
    return CalendarEvent(
        id=event_id,
        immutable_id=event_id,
        subject=subject,
        start=datetime(2026, 10, 13, hour, tzinfo=UTC),
        required_attendees=[
            Recipient(name=f"Req {n}", address=f"req{n}@example.invalid") for n in range(required)
        ],
        optional_attendees=[
            Recipient(name=f"Opt {n}", address=f"opt{n}@example.invalid") for n in range(optional)
        ],
        body_text=body,
    )


class GraphStub:
    """Graph's calendar view: answers the queued pages in order, the last one again after that, or raises
    ``error`` on every call."""

    def __init__(self, pages: list[CalendarViewPage] | None = None, error: Exception | None = None) -> None:
        self.pages = list(pages or [])
        self.error = error
        self.calls: list[dict[str, Any]] = []

    async def list_events(self, **kwargs: Any) -> CalendarViewPage:
        self.calls.append(kwargs)
        if self.error is not None:
            raise self.error
        return self.pages.pop(0) if len(self.pages) > 1 else self.pages[0]

    async def list_calendars(self) -> list[CalendarInfo]:
        return [CalendarInfo(id="synthetic-calendar", name="Planning", is_default=False)]


class FlakyGraph:
    """Answers its first call with ``page``, then fails: a failure in the middle of a listing."""

    def __init__(self, page: CalendarViewPage, error: Exception) -> None:
        self.page = page
        self.error = error
        self.calls = 0

    async def list_events(self, **kwargs: Any) -> CalendarViewPage:
        self.calls += 1
        if self.calls > 1:
            raise self.error
        return self.page

    async def list_calendars(self) -> list[CalendarInfo]:
        return []


async def test_graph_listing_is_in_time_order_and_continues_with_a_cursor() -> None:
    graph = GraphStub(
        [
            CalendarViewPage(events=[_event("late", 15), _event("early", 9)], next_link=LINK_2),
            CalendarViewPage(events=[_event("middle", 12)]),
        ]
    )
    listing = CalendarListing(graph_view=graph)

    first = await listing.list_events(since=SINCE, until=UNTIL, limit=2)

    assert [event.id for event in first.events] == ["early", "late"]
    assert not first.coverage.complete
    assert first.cursor is not None
    assert graph.calls[0]["page"] is None and graph.calls[0]["page_size"] == 2

    second = await listing.list_events(since=SINCE, until=UNTIL, limit=2, cursor=first.cursor)

    assert [event.id for event in second.events] == ["middle"]
    assert second.coverage.complete and second.cursor is None
    assert graph.calls[1]["page"] == LINK_2


async def test_calendar_catalog_and_selected_calendar_flow_through_the_service_and_cursor() -> None:
    graph = GraphStub(
        [
            CalendarViewPage(events=[_event("first", 9)], next_link=LINK_2),
            CalendarViewPage(events=[_event("next", 10)]),
        ]
    )
    listing = CalendarListing(graph_view=graph)

    calendars = await listing.list_calendars()
    first = await listing.list_events(since=SINCE, until=UNTIL, limit=1, calendar_id="synthetic-calendar")
    second = await listing.list_events(
        since=SINCE, until=UNTIL, cursor=first.cursor, calendar_id="ignored-calendar"
    )

    assert calendars[0].name == "Planning" and not calendars[0].is_default
    assert graph.calls[0]["calendar_id"] == "synthetic-calendar"
    assert graph.calls[1]["calendar_id"] == "synthetic-calendar"
    assert [event.id for event in second.events] == ["next"]


async def test_contains_matches_subject_and_description_case_and_accent_insensitively() -> None:
    events = [
        _event("subject-hit", 9, subject="Revisão de Keycloak"),
        _event("body-hit", 10, subject="Weekly", body="Pen tests with the vendor"),
        _event("miss", 11, subject="Other", body="nothing relevant"),
    ]
    graph = GraphStub([CalendarViewPage(events=events)])
    listing = CalendarListing(graph_view=graph)

    found = await listing.list_events(since=SINCE, until=UNTIL, contains=["KEYCLOAK"], limit=10)
    assert [event.id for event in found.events] == ["subject-hit"]

    accented = await listing.list_events(
        since=SINCE, until=UNTIL, contains=["revisao"], limit=10, include_body=True
    )
    assert [event.id for event in accented.events] == ["subject-hit"]

    described = await listing.list_events(since=SINCE, until=UNTIL, contains=["pen   tests"], limit=10)
    assert [event.id for event in described.events] == ["body-hit"]
    assert all(event.body_text is None for event in described.events)
    assert graph.calls[-1]["include_body"] is True  # descriptions are read to match, not returned


async def test_a_filtered_listing_keeps_reading_pages_until_it_has_matches() -> None:
    graph = GraphStub(
        [
            CalendarViewPage(events=[_event("no-1", 9)], next_link=LINK_2),
            CalendarViewPage(events=[_event("yes", 10, body="synthetic keyword here")]),
        ]
    )
    listing = CalendarListing(graph_view=graph)

    result = await listing.list_events(since=SINCE, until=UNTIL, contains=["keyword"], limit=1)

    assert [event.id for event in result.events] == ["yes"]
    assert result.coverage.complete and result.coverage.scanned == 2
    assert [call["page_size"] for call in graph.calls] == [100, 100]


async def test_a_call_that_spends_its_read_budget_says_so() -> None:
    graph = GraphStub([CalendarViewPage(events=[_event("no", 9)], next_link=LINK_2)])
    listing = CalendarListing(graph_view=graph)

    result = await listing.list_events(since=SINCE, until=UNTIL, contains=["keyword"], limit=1)

    assert result.events == [] and result.cursor is not None and not result.coverage.complete
    assert len(graph.calls) == MAX_GRAPH_PAGES
    assert any("read budget" in note for note in result.coverage.notes)

    full = await listing.list_events(since=SINCE, until=UNTIL, limit=1)
    assert not full.coverage.complete and not any("read budget" in n for n in full.coverage.notes)


async def test_descriptions_and_attendees_are_bounded_and_counted() -> None:
    body = "x" * (MAX_LISTED_BODY_CHARS + 10)
    event = _event("many", 9, body=body, required=MAX_LISTED_ATTENDEES + 5, optional=2)
    listing = CalendarListing(graph_view=GraphStub([CalendarViewPage(events=[event])]))

    result = await listing.list_events(since=SINCE, until=UNTIL, include_body=True)

    (listed,) = result.events
    assert listed.attendee_count == MAX_LISTED_ATTENDEES + 7
    assert len(listed.required_attendees) == MAX_LISTED_ATTENDEES
    assert len(listed.optional_attendees) == 2
    assert listed.body_text == "x" * MAX_LISTED_BODY_CHARS and listed.body_truncated is True
    assert any("Descriptions are cut" in note for note in result.coverage.notes)


@pytest.mark.parametrize(
    "error",
    [
        Upstream("synthetic outage"),
        Throttled("synthetic throttled"),
        AuthenticationRequired("synthetic sign-in needed", command="outlook-connector auth"),
        AccountMismatch("synthetic other account"),
    ],
)
async def test_a_graph_failure_reaches_the_caller_as_it_is(error: Exception) -> None:
    listing = CalendarListing(graph_view=GraphStub(error=error))

    with pytest.raises(type(error)):
        await listing.list_events(since=SINCE, until=UNTIL)


async def test_a_graph_failure_on_a_later_page_reaches_the_caller() -> None:
    graph = FlakyGraph(
        CalendarViewPage(events=[_event("first", 9)], next_link=LINK_2), Upstream("synthetic outage")
    )
    listing = CalendarListing(graph_view=graph)
    first = await listing.list_events(since=SINCE, until=UNTIL, limit=1)

    with pytest.raises(Upstream, match="synthetic outage"):
        await listing.list_events(since=SINCE, until=UNTIL, limit=1, cursor=first.cursor)


async def test_window_limit_and_filter_inputs_are_checked_at_the_entry_point() -> None:
    listing = CalendarListing(graph_view=GraphStub([]))

    with pytest.raises(InvalidRequest, match="earlier than until"):
        await listing.list_events(since=UNTIL, until=SINCE)
    with pytest.raises(InvalidRequest, match="at most 366 days"):
        await listing.list_events(since=datetime(2025, 1, 1, tzinfo=UTC), until=UNTIL)
    with pytest.raises(InvalidRequest, match="limit must be between 1 and 200"):
        await listing.list_events(since=SINCE, until=UNTIL, limit=0)
    with pytest.raises(InvalidRequest, match="limit must be between 1 and 200"):
        await listing.list_events(since=SINCE, until=UNTIL, limit=201)
    with pytest.raises(InvalidRequest, match="calendar_id"):
        await listing.list_events(since=SINCE, until=UNTIL, calendar_id="bad\nid")
    with pytest.raises(InvalidRequest, match="non-empty text"):
        await listing.list_events(since=SINCE, until=UNTIL, contains=["   "])
    with pytest.raises(InvalidRequest, match="at most 200 characters"):
        await listing.list_events(since=SINCE, until=UNTIL, contains=["x" * 201])
    with pytest.raises(InvalidRequest, match="list of 1 to 20 terms"):
        await listing.list_events(since=SINCE, until=UNTIL, contains=[])
    with pytest.raises(InvalidRequest, match="list of 1 to 20 terms"):
        await listing.list_events(since=SINCE, until=UNTIL, contains=[f"t{n}" for n in range(21)])


async def test_malformed_and_tampered_calendar_cursors_are_refused() -> None:
    listing = CalendarListing(graph_view=GraphStub([]))
    base = {"since": SINCE.isoformat(), "until": UNTIL.isoformat(), "include_body": False}

    without_link = cursors.encode(CURSOR_KIND, contains=None, **base)
    link_not_text = cursors.encode(CURSOR_KIND, link=7, contains=None, **base)
    filter_not_list = cursors.encode(CURSOR_KIND, link=LINK_2, contains="keyword", **base)
    bad_time = cursors.encode(
        CURSOR_KIND,
        link=LINK_2,
        contains=None,
        since="not a time",
        until=UNTIL.isoformat(),
        include_body=False,
    )

    for cursor in ("not-a-cursor", without_link, link_not_text, filter_not_list, bad_time):
        with pytest.raises(InvalidRequest):
            await listing.list_events(since=SINCE, until=UNTIL, cursor=cursor)


async def test_contains_matches_each_field_on_its_own() -> None:
    split = _event("split", 9, subject="Project", body="Alpha")
    listing = CalendarListing(graph_view=GraphStub([CalendarViewPage(events=[split])]))

    across = await listing.list_events(since=SINCE, until=UNTIL, contains=["project alpha"])
    assert across.events == []  # neither field contains the phrase, so the join must not create one

    whole = _event("whole", 9, subject="Project alpha", body="Other text")
    listing = CalendarListing(graph_view=GraphStub([CalendarViewPage(events=[whole])]))

    within = await listing.list_events(since=SINCE, until=UNTIL, contains=["project alpha"])
    assert [event.id for event in within.events] == ["whole"]


async def test_contains_matches_any_term_and_never_splits_a_term() -> None:
    events = [
        _event("budget", 9, subject="Budget planning"),
        _event("review", 10, subject="Design review"),
        _event("phrase", 11, subject="Weekly", body="The budget review is moved"),
        _event("neither", 12, subject="Lunch"),
    ]
    graph = GraphStub([CalendarViewPage(events=events, next_link=LINK_2), CalendarViewPage(events=[])])
    listing = CalendarListing(graph_view=graph)

    either = await listing.list_events(since=SINCE, until=UNTIL, contains=["budget", "review"], limit=10)
    assert [event.id for event in either.events] == ["budget", "review", "phrase"]

    graph.pages = [CalendarViewPage(events=events)]
    phrase = await listing.list_events(since=SINCE, until=UNTIL, contains=["budget review"], limit=10)
    assert [event.id for event in phrase.events] == ["phrase"]


async def test_the_cursor_keeps_every_contains_term() -> None:
    graph = GraphStub(
        [
            CalendarViewPage(events=[_event("alpha", 9, subject="Alpha sync")], next_link=LINK_2),
            CalendarViewPage(events=[_event("beta", 10, subject="Beta sync"), _event("other", 11)]),
        ]
    )
    listing = CalendarListing(graph_view=graph)

    first = await listing.list_events(since=SINCE, until=UNTIL, contains=["alpha", "beta"], limit=1)
    assert [event.id for event in first.events] == ["alpha"] and first.cursor is not None

    second = await listing.list_events(since=SINCE, until=UNTIL, contains=["ignored"], cursor=first.cursor)
    assert [event.id for event in second.events] == ["beta"] and second.coverage.complete


async def test_a_cursor_cannot_widen_the_window_or_the_filter() -> None:
    listing = CalendarListing(graph_view=GraphStub([]))
    too_wide = cursors.encode(
        CURSOR_KIND,
        link=LINK_2,
        since=datetime(2025, 1, 1, tzinfo=UTC).isoformat(),
        until=UNTIL.isoformat(),
        contains=None,
        include_body=False,
    )
    long_filter = cursors.encode(
        CURSOR_KIND,
        link=LINK_2,
        since=SINCE.isoformat(),
        until=UNTIL.isoformat(),
        contains=["x" * 201],
        include_body=False,
    )
    blank_filter = cursors.encode(
        CURSOR_KIND,
        link=LINK_2,
        since=SINCE.isoformat(),
        until=UNTIL.isoformat(),
        contains=["   "],
        include_body=False,
    )

    for cursor, message in (
        (too_wide, "at most 366 days"),
        (long_filter, "at most 200 characters"),
        (blank_filter, "non-empty text"),
    ):
        with pytest.raises(InvalidRequest, match=message):
            await listing.list_events(since=SINCE, until=UNTIL, cursor=cursor)

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

import pytest

from outlook_connector.domain.errors import (
    AccountMismatch,
    AuthenticationRequired,
    InvalidRequest,
    NotFound,
    Throttled,
    Upstream,
)
from outlook_connector.domain.models import CalendarEvent, Recipient
from outlook_connector.remote.ports import CalendarViewPage
from outlook_connector.service import cursors
from outlook_connector.service.calendar_listing import (
    CURSOR_KIND,
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


class OutlookStub:
    """Outlook Web's window: every event of the window in one answer, or ``error``."""

    def __init__(self, events: list[CalendarEvent] | None = None, error: Exception | None = None) -> None:
        self.events = list(events or [])
        self.error = error
        self.calls = 0

    async def list_events(self, **kwargs: Any) -> list[CalendarEvent]:
        self.calls += 1
        if self.error is not None:
            raise self.error
        return list(self.events)


async def test_graph_listing_is_in_time_order_and_continues_with_a_cursor() -> None:
    graph = GraphStub(
        [
            CalendarViewPage(events=[_event("late", 15), _event("early", 9)], next_link=LINK_2),
            CalendarViewPage(events=[_event("middle", 12)]),
        ]
    )
    listing = CalendarListing(graph_view=graph, outlook_view=OutlookStub())

    first = await listing.list_events(since=SINCE, until=UNTIL, limit=2)

    assert [event.id for event in first.events] == ["early", "late"]
    assert first.coverage.source == "graph" and not first.coverage.complete
    assert first.cursor is not None
    assert graph.calls[0]["page"] is None and graph.calls[0]["page_size"] == 2

    second = await listing.list_events(since=SINCE, until=UNTIL, limit=2, cursor=first.cursor)

    assert [event.id for event in second.events] == ["middle"]
    assert second.coverage.complete and second.cursor is None
    assert graph.calls[1]["page"] == LINK_2


async def test_contains_matches_subject_and_description_case_and_accent_insensitively() -> None:
    events = [
        _event("subject-hit", 9, subject="Revisão de Keycloak"),
        _event("body-hit", 10, subject="Weekly", body="Pen tests with the vendor"),
        _event("miss", 11, subject="Other", body="nothing relevant"),
    ]
    graph = GraphStub([CalendarViewPage(events=events)])
    listing = CalendarListing(graph_view=graph, outlook_view=OutlookStub())

    found = await listing.list_events(since=SINCE, until=UNTIL, contains="KEYCLOAK", limit=10)
    assert [event.id for event in found.events] == ["subject-hit"]

    accented = await listing.list_events(
        since=SINCE, until=UNTIL, contains="revisao", limit=10, include_body=True
    )
    assert [event.id for event in accented.events] == ["subject-hit"]

    described = await listing.list_events(since=SINCE, until=UNTIL, contains="pen   tests", limit=10)
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
    listing = CalendarListing(graph_view=graph, outlook_view=OutlookStub())

    result = await listing.list_events(since=SINCE, until=UNTIL, contains="keyword", limit=1)

    assert [event.id for event in result.events] == ["yes"]
    assert result.coverage.complete and result.coverage.scanned == 2
    assert [call["page_size"] for call in graph.calls] == [100, 100]


async def test_descriptions_and_attendees_are_bounded_and_counted() -> None:
    body = "x" * (MAX_LISTED_BODY_CHARS + 10)
    event = _event("many", 9, body=body, required=MAX_LISTED_ATTENDEES + 5, optional=2)
    listing = CalendarListing(
        graph_view=GraphStub([CalendarViewPage(events=[event])]), outlook_view=OutlookStub()
    )

    result = await listing.list_events(since=SINCE, until=UNTIL, include_body=True)

    (listed,) = result.events
    assert listed.attendee_count == MAX_LISTED_ATTENDEES + 7
    assert len(listed.required_attendees) == MAX_LISTED_ATTENDEES
    assert len(listed.optional_attendees) == 2
    assert listed.body_text == "x" * MAX_LISTED_BODY_CHARS and listed.body_truncated is True
    assert any("Descriptions are cut" in note for note in result.coverage.notes)


async def test_a_graph_failure_on_the_first_page_falls_back_to_outlook_web() -> None:
    graph = GraphStub(error=Upstream("synthetic HTTP 403"))
    outlook = OutlookStub([_event("c", 15), _event("a", 9), _event("b", 12)])
    listing = CalendarListing(graph_view=graph, outlook_view=outlook)

    first = await listing.list_events(since=SINCE, until=UNTIL, limit=2)

    assert [event.id for event in first.events] == ["a", "b"]
    assert first.coverage.source == "outlook" and not first.coverage.complete
    assert any("synthetic HTTP 403" in note for note in first.coverage.notes)
    assert first.cursor is not None

    second = await listing.list_events(since=SINCE, until=UNTIL, limit=2, cursor=first.cursor)

    assert [event.id for event in second.events] == ["c"]
    assert second.coverage.source == "outlook" and second.coverage.complete
    assert outlook.calls == 2 and graph.calls == [
        {"since": SINCE, "until": UNTIL, "page_size": 2, "page": None, "include_body": False}
    ]


@pytest.mark.parametrize("error", [Upstream("synthetic outage"), Throttled("synthetic throttled")])
async def test_fallback_answers_for_service_failures(error: Exception) -> None:
    listing = CalendarListing(graph_view=GraphStub(error=error), outlook_view=OutlookStub([_event("a", 9)]))

    result = await listing.list_events(since=SINCE, until=UNTIL)

    assert result.coverage.source == "outlook" and [event.id for event in result.events] == ["a"]


async def test_an_authentication_failure_on_graph_falls_back_when_outlook_web_can_answer() -> None:
    graph = GraphStub(
        error=AuthenticationRequired("synthetic sign-in needed", command="outlook-connector auth")
    )
    listing = CalendarListing(graph_view=graph, outlook_view=OutlookStub([_event("a", 9)]))

    result = await listing.list_events(since=SINCE, until=UNTIL)

    assert result.coverage.source == "outlook"


async def test_a_later_page_stays_on_the_service_that_answered_the_first() -> None:
    graph = GraphStub(error=Upstream("synthetic outage"))
    outlook = OutlookStub([_event("a", 9), _event("b", 10), _event("c", 11)])
    listing = CalendarListing(graph_view=graph, outlook_view=outlook)
    first = await listing.list_events(since=SINCE, until=UNTIL, limit=1)

    second = await listing.list_events(since=SINCE, until=UNTIL, limit=1, cursor=first.cursor)

    assert second.coverage.source == "outlook" and [event.id for event in second.events] == ["b"]
    assert len(graph.calls) == 1  # the later page did not go back to Graph


@pytest.mark.parametrize(
    "error",
    [
        InvalidRequest("synthetic bad request"),
        NotFound("synthetic missing"),
        AccountMismatch("synthetic other account"),
    ],
)
async def test_our_own_errors_and_account_mismatches_are_not_retried_on_outlook_web(error: Exception) -> None:
    outlook = OutlookStub([_event("a", 9)])
    listing = CalendarListing(graph_view=GraphStub(error=error), outlook_view=outlook)

    with pytest.raises(type(error)):
        await listing.list_events(since=SINCE, until=UNTIL)

    assert outlook.calls == 0


async def test_contains_and_descriptions_cannot_fall_back_because_outlook_web_has_no_descriptions() -> None:
    outlook = OutlookStub([_event("a", 9)])
    listing = CalendarListing(graph_view=GraphStub(error=Upstream("synthetic outage")), outlook_view=outlook)

    with pytest.raises(Upstream, match="cannot run now"):
        await listing.list_events(since=SINCE, until=UNTIL, contains="keyword")
    with pytest.raises(Upstream, match="cannot run now"):
        await listing.list_events(since=SINCE, until=UNTIL, include_body=True)

    assert outlook.calls == 0


async def test_a_graph_failure_in_the_middle_of_a_listing_is_not_hidden_by_the_fallback() -> None:
    outlook = OutlookStub([_event("a", 9)])
    graph = FlakyGraph(
        CalendarViewPage(events=[_event("first", 9)], next_link=LINK_2), Upstream("synthetic outage")
    )
    listing = CalendarListing(graph_view=graph, outlook_view=outlook)
    first = await listing.list_events(since=SINCE, until=UNTIL, limit=1)

    with pytest.raises(Upstream):
        await listing.list_events(since=SINCE, until=UNTIL, limit=1, cursor=first.cursor)

    assert outlook.calls == 0


async def test_when_outlook_web_also_fails_the_error_names_both_reasons() -> None:
    listing = CalendarListing(
        graph_view=GraphStub(error=Upstream("synthetic graph 403")),
        outlook_view=OutlookStub(error=Upstream("synthetic outlook 500")),
    )

    with pytest.raises(Upstream, match="synthetic graph 403.*synthetic outlook 500"):
        await listing.list_events(since=SINCE, until=UNTIL)


async def test_an_outlook_sign_in_need_is_raised_so_the_sign_in_command_reaches_the_user() -> None:
    listing = CalendarListing(
        graph_view=GraphStub(error=Upstream("synthetic graph 403")),
        outlook_view=OutlookStub(
            error=AuthenticationRequired("synthetic sign-in needed", command="outlook-connector auth")
        ),
    )

    with pytest.raises(AuthenticationRequired):
        await listing.list_events(since=SINCE, until=UNTIL)


async def test_the_outlook_fallback_pages_by_key_and_reads_the_window_each_time() -> None:
    outlook = OutlookStub([_event("a", 9), _event("b", 10), _event("c", 11)])
    listing = CalendarListing(graph_view=GraphStub(error=Upstream("synthetic outage")), outlook_view=outlook)

    first = await listing.list_events(since=SINCE, until=UNTIL, limit=2)
    second = await listing.list_events(since=SINCE, until=UNTIL, limit=2, cursor=first.cursor)

    assert [event.id for event in first.events] == ["a", "b"]
    assert [event.id for event in second.events] == ["c"]
    assert second.coverage.complete and outlook.calls == 2


async def test_window_limit_and_filter_inputs_are_checked_at_the_entry_point() -> None:
    listing = CalendarListing(graph_view=GraphStub([]), outlook_view=OutlookStub())

    with pytest.raises(InvalidRequest, match="earlier than until"):
        await listing.list_events(since=UNTIL, until=SINCE)
    with pytest.raises(InvalidRequest, match="at most 366 days"):
        await listing.list_events(since=datetime(2025, 1, 1, tzinfo=UTC), until=UNTIL)
    with pytest.raises(InvalidRequest, match="limit must be between 1 and 200"):
        await listing.list_events(since=SINCE, until=UNTIL, limit=0)
    with pytest.raises(InvalidRequest, match="limit must be between 1 and 200"):
        await listing.list_events(since=SINCE, until=UNTIL, limit=201)
    with pytest.raises(InvalidRequest, match="must not be empty"):
        await listing.list_events(since=SINCE, until=UNTIL, contains="   ")
    with pytest.raises(InvalidRequest, match="at most 200 characters"):
        await listing.list_events(since=SINCE, until=UNTIL, contains="x" * 201)


async def test_malformed_and_tampered_calendar_cursors_are_refused() -> None:
    listing = CalendarListing(graph_view=GraphStub([]), outlook_view=OutlookStub())
    base = {"since": SINCE.isoformat(), "until": UNTIL.isoformat(), "include_body": False}

    unknown_source = cursors.encode(CURSOR_KIND, source="elsewhere", contains=None, **base)
    outlook_with_filter = cursors.encode(
        CURSOR_KIND, source="outlook", after_start=SINCE.isoformat(), after_id="a", contains="keyword", **base
    )
    bad_key = cursors.encode(
        CURSOR_KIND, source="outlook", after_start="not a time", after_id="a", contains=None, **base
    )
    graph_without_link = cursors.encode(CURSOR_KIND, source="graph", contains=None, **base)

    for cursor in ("not-a-cursor", unknown_source, outlook_with_filter, bad_key, graph_without_link):
        with pytest.raises(InvalidRequest):
            await listing.list_events(since=SINCE, until=UNTIL, cursor=cursor)


async def test_contains_matches_each_field_on_its_own() -> None:
    split = _event("split", 9, subject="Project", body="Alpha")
    listing = CalendarListing(
        graph_view=GraphStub([CalendarViewPage(events=[split])]), outlook_view=OutlookStub()
    )

    across = await listing.list_events(since=SINCE, until=UNTIL, contains="project alpha")
    assert across.events == []  # neither field contains the phrase, so the join must not create one

    whole = _event("whole", 9, subject="Project alpha", body="Other text")
    listing = CalendarListing(
        graph_view=GraphStub([CalendarViewPage(events=[whole])]), outlook_view=OutlookStub()
    )

    within = await listing.list_events(since=SINCE, until=UNTIL, contains="project alpha")
    assert [event.id for event in within.events] == ["whole"]


async def test_a_cursor_cannot_widen_the_window_or_the_filter() -> None:
    listing = CalendarListing(graph_view=GraphStub([]), outlook_view=OutlookStub())
    too_wide = cursors.encode(
        CURSOR_KIND,
        source="outlook",
        after_start=SINCE.isoformat(),
        after_id="a",
        since=datetime(2025, 1, 1, tzinfo=UTC).isoformat(),
        until=UNTIL.isoformat(),
        contains=None,
        include_body=False,
    )
    long_filter = cursors.encode(
        CURSOR_KIND,
        source="graph",
        link=LINK_2,
        since=SINCE.isoformat(),
        until=UNTIL.isoformat(),
        contains="x" * 201,
        include_body=False,
    )
    blank_filter = cursors.encode(
        CURSOR_KIND,
        source="graph",
        link=LINK_2,
        since=SINCE.isoformat(),
        until=UNTIL.isoformat(),
        contains="   ",
        include_body=False,
    )

    for cursor, message in (
        (too_wide, "at most 366 days"),
        (long_filter, "at most 200 characters"),
        (blank_filter, "must not be empty"),
    ):
        with pytest.raises(InvalidRequest, match=message):
            await listing.list_events(since=SINCE, until=UNTIL, cursor=cursor)


async def test_the_outlook_fallback_resumes_after_the_last_event_when_an_earlier_event_appears() -> None:
    outlook = OutlookStub([_event("a", 9), _event("b", 10), _event("c", 11)])
    listing = CalendarListing(graph_view=GraphStub(error=Upstream("synthetic outage")), outlook_view=outlook)

    first = await listing.list_events(since=SINCE, until=UNTIL, limit=2)
    outlook.events.insert(0, _event("early", 8))  # appears before the first page's last event
    second = await listing.list_events(since=SINCE, until=UNTIL, limit=2, cursor=first.cursor)

    assert [event.id for event in second.events] == ["c"]
    assert second.coverage.complete


async def test_the_outlook_fallback_orders_events_with_equal_starts_by_id() -> None:
    outlook = OutlookStub([_event("y", 9), _event("x", 9)])
    listing = CalendarListing(graph_view=GraphStub(error=Upstream("synthetic outage")), outlook_view=outlook)

    first = await listing.list_events(since=SINCE, until=UNTIL, limit=1)
    second = await listing.list_events(since=SINCE, until=UNTIL, limit=1, cursor=first.cursor)

    assert [event.id for event in first.events] == ["x"]
    assert [event.id for event in second.events] == ["y"] and second.coverage.complete

from __future__ import annotations

from datetime import UTC, datetime

import pytest

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import CalendarEvent
from outlook_connector.remote.ports import CalendarSearchPage
from outlook_connector.service import cursors
from outlook_connector.service.calendar import MAX_EVENT_OFFSET, Calendar


class SearchStub:
    def __init__(self, pages: list[CalendarSearchPage]) -> None:
        self.pages = iter(pages)
        self.calls: list[dict[str, object]] = []

    async def search_events(self, query: str, **kwargs: object) -> CalendarSearchPage:
        self.calls.append({"query": query, **kwargs})
        return next(self.pages)


class ItemsStub:
    def __init__(self) -> None:
        self.event_ids: list[str] = []

    async def get_event(self, event_id: str) -> CalendarEvent:
        self.event_ids.append(event_id)
        return CalendarEvent(id=event_id, subject="Synthetic event")


def _event(event_id: str, start: str | None) -> CalendarEvent:
    return CalendarEvent(
        id=event_id,
        subject="Synthetic event",
        start=datetime.fromisoformat(start.replace("Z", "+00:00")) if start else None,
    )


async def test_calendar_search_bounds_dates_and_preserves_remote_paging() -> None:
    search = SearchStub(
        [
            CalendarSearchPage(
                events=[
                    _event("synthetic-in-range", "2026-10-10T12:00:00Z"),
                    _event("synthetic-outside", "2026-11-10T12:00:00Z"),
                    _event("synthetic-no-date", None),
                ],
                rows_received=3,
                more_results_available=True,
                result_set_found=True,
            ),
            CalendarSearchPage(
                events=[_event("synthetic-next-page", "2026-10-11T12:00:00Z")],
                rows_received=1,
                more_results_available=False,
                result_set_found=True,
            ),
        ]
    )
    items = ItemsStub()
    calendar = Calendar(search, items)
    since = datetime(2026, 10, 1, tzinfo=UTC)
    until = datetime(2026, 10, 31, tzinfo=UTC)

    first = await calendar.search_events("synthetic event", since=since, until=until, limit=10)
    assert [event.id for event in first.events] == ["synthetic-in-range"]
    assert not first.coverage.complete and first.cursor
    assert any("1 event(s) with no start time" in note for note in first.coverage.notes)

    second = await calendar.search_events("synthetic event", cursor=first.cursor)
    assert [event.id for event in second.events] == ["synthetic-next-page"]
    assert second.coverage.complete and second.cursor is None
    assert search.calls[1]["offset"] == 3
    assert search.calls[0]["cvid"] == search.calls[1]["cvid"]
    assert search.calls[0]["logical_id"] == search.calls[1]["logical_id"]


async def test_calendar_search_keeps_incomplete_response_distinct_from_zero_hits() -> None:
    calendar = Calendar(
        SearchStub(
            [
                CalendarSearchPage(
                    events=[],
                    rows_received=0,
                    more_results_available=None,
                    result_set_found=False,
                )
            ]
        ),
        ItemsStub(),
    )
    result = await calendar.search_events("synthetic query")
    assert result.events == [] and not result.coverage.complete
    assert any("zero events are not confirmed" in note for note in result.coverage.notes)


async def test_calendar_search_enforces_query_and_page_bounds_at_service_boundary() -> None:
    calendar = Calendar(SearchStub([]), ItemsStub())

    with pytest.raises(InvalidRequest, match="non-empty"):
        await calendar.search_events(" ")
    with pytest.raises(InvalidRequest, match="512 characters"):
        await calendar.search_events("x" * 513)
    with pytest.raises(InvalidRequest, match="between 1 and 25"):
        await calendar.search_events("synthetic query", limit=26)


async def test_calendar_cursor_and_event_reference_validation() -> None:
    search = SearchStub(
        [
            CalendarSearchPage(
                events=[_event("synthetic", None)],
                rows_received=1,
                more_results_available=True,
                result_set_found=True,
            )
        ]
    )
    items = ItemsStub()
    calendar = Calendar(search, items)
    first = await calendar.search_events("synthetic query")
    assert first.cursor

    with pytest.raises(InvalidRequest, match="does not match this search cursor"):
        await calendar.search_events("different query", cursor=first.cursor)
    with pytest.raises(InvalidRequest, match="control character"):
        await calendar.get_event("synthetic\nevent")

    await calendar.get_event("synthetic-event-id")
    assert items.event_ids == ["synthetic-event-id"]


async def test_calendar_search_does_not_emit_cursor_past_offset_limit() -> None:
    search = SearchStub(
        [
            CalendarSearchPage(
                events=[_event("synthetic-last-page", "2026-10-10T12:00:00Z")],
                rows_received=1,
                more_results_available=True,
                result_set_found=True,
            )
        ]
    )
    calendar = Calendar(search, ItemsStub())
    cursor = cursors.encode(
        "calendar_event_search",
        query="synthetic query",
        offset=MAX_EVENT_OFFSET,
        page_size=25,
        cvid="00000000-0000-0000-0000-000000000001",
        logical_id="00000000-0000-0000-0000-000000000002",
        since=None,
        until=None,
    )

    result = await calendar.search_events("synthetic query", cursor=cursor)

    assert result.cursor is None
    assert not result.coverage.complete
    assert any("maximum event-search offset" in note for note in result.coverage.notes)

"""Bounded online Exchange event search and explicit event hydration."""

from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID, uuid4

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import CalendarEvent, CalendarSearchCoverage, CalendarSearchResult
from outlook_connector.remote.ports import CalendarItemReader, CalendarSearchReader
from outlook_connector.service import cursors

MAX_EVENT_PAGE_SIZE = 25  # the standalone-proven Exchange Search v2 page size
MAX_EVENT_OFFSET = 100_000
MAX_EVENT_QUERY_CHARS = 512


class Calendar:
    def __init__(self, search: CalendarSearchReader, items: CalendarItemReader) -> None:
        self._search = search
        self._items = items

    async def search_events(
        self,
        query: str,
        *,
        since: datetime | None = None,
        until: datetime | None = None,
        limit: int = 25,
        cursor: str | None = None,
    ) -> CalendarSearchResult:
        """Validate public search inputs, preserve cursor filters, then return one bounded page."""
        if not isinstance(query, str) or not query.strip():
            raise InvalidRequest("query must be a non-empty string.")
        query = query.strip()
        if len(query) > MAX_EVENT_QUERY_CHARS:
            raise InvalidRequest(f"query must be no longer than {MAX_EVENT_QUERY_CHARS} characters.")
        since, until = _utc(since), _utc(until)
        if since and until and since > until:
            raise InvalidRequest("since must be earlier than or equal to until.")
        if not 1 <= limit <= MAX_EVENT_PAGE_SIZE:
            raise InvalidRequest(f"limit must be between 1 and {MAX_EVENT_PAGE_SIZE}.")

        if cursor:
            state = cursors.decode(cursor, "calendar_event_search")
            if state.get("query") != query:
                raise InvalidRequest("The query does not match this search cursor; start the search again.")
            offset = state.get("offset")
            page_size = state.get("page_size")
            cvid = _cursor_uuid(state.get("cvid"), "Cvid")
            logical_id = _cursor_uuid(state.get("logical_id"), "LogicalId")
            since, until = _parse_datetime(state.get("since")), _parse_datetime(state.get("until"))
            if type(offset) is not int or not 0 <= offset <= MAX_EVENT_OFFSET:
                raise InvalidRequest("The calendar cursor has an invalid offset.")
            if type(page_size) is not int or not 1 <= page_size <= MAX_EVENT_PAGE_SIZE:
                raise InvalidRequest("The calendar cursor has an invalid page size.")
        else:
            offset = 0
            page_size = limit
            cvid, logical_id = str(uuid4()), str(uuid4())

        page = await self._search.search_events(
            query,
            offset=offset,
            page_size=page_size,
            cvid=cvid,
            logical_id=logical_id,
        )
        events, excluded_unknown_dates = _in_window(page.events, since, until)
        next_offset = offset + page.rows_received
        next_cursor = None
        notes_at_offset_limit = False
        if page.more_results_available and page.rows_received:
            if next_offset <= MAX_EVENT_OFFSET:
                next_cursor = cursors.encode(
                    "calendar_event_search",
                    query=query,
                    offset=next_offset,
                    page_size=page_size,
                    cvid=cvid,
                    logical_id=logical_id,
                    since=since.isoformat() if since else None,
                    until=until.isoformat() if until else None,
                )
            else:
                notes_at_offset_limit = True
        complete = page.result_set_found and page.more_results_available is False
        notes = [
            "Recurring-occurrence result totals are not reported because Search v2 totals were inconsistent.",
            "Calendar events belong to Outlook; related Teams meeting chats and artifacts belong to Teams.",
        ]
        if not page.result_set_found:
            notes.append("Search returned no interpretable Event result set; zero events are not confirmed.")
        elif page.more_results_available is None:
            notes.append("Search did not report whether more events are available.")
        elif page.more_results_available and not page.rows_received:
            notes.append("Search reported more events but returned no rows to advance the offset.")
        elif notes_at_offset_limit:
            notes.append(
                f"The maximum event-search offset ({MAX_EVENT_OFFSET}) was reached; no cursor is available."
            )
        elif page.more_results_available:
            notes.append("Follow the cursor only if additional event matches are useful.")
        if excluded_unknown_dates:
            notes.append(
                f"{excluded_unknown_dates} event(s) with no start time were omitted by date filters."
            )
        if any(event.id is None for event in events):
            notes.append(
                "Some results did not contain an OWS item id and cannot be hydrated by get_calendar_event."
            )
        return CalendarSearchResult(
            query=query,
            events=events,
            cursor=next_cursor,
            coverage=CalendarSearchCoverage(
                complete=complete,
                offset=offset,
                page_size=page_size,
                more_results_available=page.more_results_available,
                notes=notes,
            ),
        )

    async def get_event(self, event_id: str) -> CalendarEvent:
        """Validate the caller-supplied event reference before OWS hydration."""
        if not isinstance(event_id, str) or not event_id or len(event_id) > 4096:
            raise InvalidRequest("event_id must be a non-empty Outlook event reference.")
        if any(ord(character) < 32 for character in event_id):
            raise InvalidRequest("event_id contains a control character.")
        return await self._items.get_event(event_id)


def _in_window(
    events: list[CalendarEvent], since: datetime | None, until: datetime | None
) -> tuple[list[CalendarEvent], int]:
    if since is None and until is None:
        return events, 0
    selected: list[CalendarEvent] = []
    unknown = 0
    for event in events:
        start = _utc(event.start)
        if start is None:
            unknown += 1
            continue
        if since and start < since or until and start > until:
            continue
        selected.append(event)
    return selected, unknown


def _utc(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    return value.replace(tzinfo=UTC) if value.tzinfo is None else value.astimezone(UTC)


def _cursor_uuid(value: object, label: str) -> str:
    if not isinstance(value, str):
        raise InvalidRequest(f"The calendar cursor has an invalid {label}.")
    try:
        return str(UUID(value))
    except ValueError:
        raise InvalidRequest(f"The calendar cursor has an invalid {label}.") from None


def _parse_datetime(value: object) -> datetime | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise InvalidRequest("The calendar cursor has an invalid date bound.")
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        raise InvalidRequest("The calendar cursor has an invalid date bound.") from None
    return _utc(parsed)

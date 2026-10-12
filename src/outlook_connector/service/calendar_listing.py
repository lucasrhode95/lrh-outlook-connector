"""Calendar listing: the events in a window, Graph first, Outlook Web only when Graph fails (research §3.18).

The listing reads the default calendar's occurrences for a window, in time order. ``contains`` and the
descriptions need Graph, because the Outlook Web fallback carries no descriptions. The fallback answers
only a first page whose Graph read failed for a service reason; a later page stays on the service that
answered the first, since cursors do not carry offsets between services.
"""

from __future__ import annotations

import unicodedata
from datetime import UTC, datetime, timedelta
from typing import Any

from outlook_connector.domain.errors import AuthenticationRequired, InvalidRequest, Throttled, Upstream
from outlook_connector.domain.models import CalendarEvent, CalendarListCoverage, CalendarListResult
from outlook_connector.remote.ports import CalendarViewPage, CalendarViewReader, CalendarWindowReader
from outlook_connector.service import cursors

MAX_WINDOW = timedelta(days=366)
MAX_LIMIT = 200
MAX_CONTAINS_CHARS = 200
GRAPH_PAGE_SIZE = 100
MAX_GRAPH_PAGES = 10  # pages read per call; a filtered listing may need several to fill its limit
MAX_LISTED_ATTENDEES = 15  # per list (required, optional); attendee_count gives the total
MAX_LISTED_BODY_CHARS = 4000
MAX_REASON_CHARS = 200
CURSOR_KIND = "calendar_list"
INVALID_CURSOR = "The calendar cursor is invalid."
# Graph failed for a service reason, so Outlook Web may answer. Bad requests, missing events and account
# mismatches are never retried on the other service.
GRAPH_FAILURES = (AuthenticationRequired, Throttled, Upstream)


class CalendarListing:
    """Lists the signed-in user's calendar events for a window (see the module docstring)."""

    def __init__(self, *, graph_view: CalendarViewReader, outlook_view: CalendarWindowReader) -> None:
        self._graph_view = graph_view
        self._outlook_view = outlook_view

    async def list_events(
        self,
        *,
        since: datetime,
        until: datetime,
        contains: str | None = None,
        include_body: bool = False,
        limit: int = 50,
        cursor: str | None = None,
    ) -> CalendarListResult:
        """Entry point: checks the window, the text filter, the limit and the cursor, then lists.

        Events overlap the window: one that started before ``since`` and runs into it is included. With a
        cursor, the window and the filters come from the cursor and the other arguments are ignored.
        ``limit`` is how many matching events to collect before stopping, so a page may hold a few more.
        """
        if not 1 <= limit <= MAX_LIMIT:
            raise InvalidRequest(f"limit must be between 1 and {MAX_LIMIT}.")
        if cursor is None:
            start, end = _window(since, until)
            return await self._first_page(start, end, limit, _contains(contains), include_body)
        state = _decode(cursor)
        if state["source"] == "graph":
            return await self._list_graph(
                state["since"], state["until"], limit, state["contains"], state["include_body"], state["link"]
            )
        return await self._list_outlook(
            state["since"], state["until"], limit, offset=state["offset"], graph_error=None
        )

    async def _first_page(
        self, start: datetime, end: datetime, limit: int, text: str | None, include_body: bool
    ) -> CalendarListResult:
        """The first page: Graph, and Outlook Web only when Graph fails for a service reason."""
        try:
            return await self._list_graph(start, end, limit, text, include_body, link=None)
        except GRAPH_FAILURES as graph_error:
            if text or include_body:
                raise Upstream(
                    f"The Graph calendar read failed ({_reason(graph_error)}). The Outlook Web fallback "
                    "has no event descriptions, so this search cannot run now; try again later, or list "
                    "without contains and include_body."
                ) from graph_error
            return await self._list_outlook(start, end, limit, offset=0, graph_error=graph_error)

    async def _list_graph(
        self,
        start: datetime,
        end: datetime,
        limit: int,
        text: str | None,
        include_body: bool,
        link: str | None,
    ) -> CalendarListResult:
        """Graph's view page by page, until ``limit`` events match, the window ends or the budget is spent."""
        needle = _fold(text) if text else None
        page_size = GRAPH_PAGE_SIZE if needle else min(limit, GRAPH_PAGE_SIZE)
        matched: list[CalendarEvent] = []
        scanned = 0
        pages = 0
        while True:
            page: CalendarViewPage = await self._graph_view.list_events(
                since=start,
                until=end,
                page_size=page_size,
                page=link,
                include_body=include_body or needle is not None,
            )
            pages += 1
            scanned += len(page.events)
            matched.extend(event for event in page.events if _matches(event, needle))
            link = page.next_link
            if link is None or len(matched) >= limit or pages >= MAX_GRAPH_PAGES:
                break
        complete = link is None
        events = _listed(sorted(matched, key=_start_key), include_body=include_body)
        notes = ["Times are UTC. Events overlap the window; recurring meetings are expanded per occurrence."]
        if needle is not None:
            notes.append(
                f"contains matched literally, case- and accent-insensitive, in each subject and description "
                f"({scanned} event(s) read)."
            )
        if any(event.body_truncated for event in events):
            notes.append(
                f"Descriptions are cut to {MAX_LISTED_BODY_CHARS:,} characters; body_truncated marks them."
            )
        if not complete:
            notes.append("More events are available: follow the cursor.")
        return CalendarListResult(
            since=start,
            until=end,
            events=events,
            cursor=None
            if complete
            else cursors.encode(
                CURSOR_KIND,
                source="graph",
                link=link,
                since=start.isoformat(),
                until=end.isoformat(),
                contains=text,
                include_body=include_body,
            ),
            coverage=CalendarListCoverage(complete=complete, source="graph", scanned=scanned, notes=notes),
        )

    async def _list_outlook(
        self,
        start: datetime,
        end: datetime,
        limit: int,
        *,
        offset: int,
        graph_error: Exception | None,
    ) -> CalendarListResult:
        """Outlook Web's window, paged here by offset. Each page reads the window again, in one call."""
        try:
            events = sorted(await self._outlook_view.list_events(since=start, until=end), key=_start_key)
        except GRAPH_FAILURES as outlook_error:
            if isinstance(outlook_error, AuthenticationRequired):
                raise
            prefix = f"The Graph calendar read failed ({_reason(graph_error)}), and " if graph_error else ""
            raise Upstream(
                f"{prefix}the Outlook Web calendar read failed: {_reason(outlook_error)}"
            ) from outlook_error
        page = events[offset : offset + limit]
        next_offset = offset + limit
        complete = next_offset >= len(events)
        notes = [
            "Times are UTC. Answered from Outlook Web's calendar view, the fallback for a failed Graph "
            "calendar read: it has no descriptions, Teams join links or contains matching."
        ]
        if graph_error is not None:
            notes.append(f"Graph calendar read failed: {_reason(graph_error)}.")
        if not complete:
            notes.append("More events are available: follow the cursor.")
        return CalendarListResult(
            since=start,
            until=end,
            events=_listed(page, include_body=False),
            cursor=None
            if complete
            else cursors.encode(
                CURSOR_KIND,
                source="outlook",
                offset=next_offset,
                since=start.isoformat(),
                until=end.isoformat(),
                contains=None,
                include_body=False,
            ),
            coverage=CalendarListCoverage(
                complete=complete, source="outlook", scanned=len(events), notes=notes
            ),
        )


def _window(since: datetime, until: datetime) -> tuple[datetime, datetime]:
    start, end = _utc(since), _utc(until)
    if start >= end:
        raise InvalidRequest("since must be earlier than until.")
    if end - start > MAX_WINDOW:
        raise InvalidRequest(f"The window must be at most {MAX_WINDOW.days} days.")
    return start, end


def _contains(value: str | None) -> str | None:
    if value is None:
        return None
    text = value.strip()
    if not text:
        raise InvalidRequest("contains must not be empty.")
    if len(text) > MAX_CONTAINS_CHARS:
        raise InvalidRequest(f"contains must be at most {MAX_CONTAINS_CHARS} characters.")
    return text


def _decode(cursor: str) -> dict[str, Any]:
    """The listing state a cursor carries. Cursors are not signed, so every value gets the same checks as
    the entry point again: a malformed or tampered cursor is an InvalidRequest."""
    state = cursors.decode(cursor, CURSOR_KIND)
    since, until = _cursor_time(state.get("since")), _cursor_time(state.get("until"))
    contains, include_body, source = state.get("contains"), state.get("include_body"), state.get("source")
    if since is None or until is None or not isinstance(include_body, bool):
        raise InvalidRequest(INVALID_CURSOR)
    graph_ok = source == "graph" and isinstance(state.get("link"), str)
    outlook_ok = (
        source == "outlook"
        and type(state.get("offset")) is int
        and state["offset"] >= 0
        and contains is None
        and not include_body
    )
    if not (graph_ok and (contains is None or isinstance(contains, str)) or outlook_ok):
        raise InvalidRequest(INVALID_CURSOR)
    start, end = _window(since, until)
    return {
        **state,
        "since": start,
        "until": end,
        "contains": _contains(contains),
        "include_body": include_body,
    }


def _cursor_time(value: object) -> datetime | None:
    if not isinstance(value, str):
        return None
    try:
        return _utc(datetime.fromisoformat(value))
    except ValueError:
        return None


def _utc(value: datetime) -> datetime:
    return value.replace(tzinfo=UTC) if value.tzinfo is None else value.astimezone(UTC)


def _fold(value: str) -> str:
    """Case- and accent-insensitive text for matching: decomposed, marks removed, casefolded, one-spaced."""
    decomposed = unicodedata.normalize("NFKD", value)
    unmarked = "".join(char for char in decomposed if not unicodedata.combining(char))
    return " ".join(unmarked.casefold().split())


def _matches(event: CalendarEvent, needle: str | None) -> bool:
    """Each field is matched on its own, so a phrase cannot span the subject and description boundary."""
    if needle is None:
        return True
    return any(needle in _fold(part) for part in (event.subject, event.body_text) if part)


def _listed(events: list[CalendarEvent], *, include_body: bool) -> list[CalendarEvent]:
    """What a listing returns per event: the first attendees and their total, and the description if asked."""
    listed: list[CalendarEvent] = []
    for event in events:
        update: dict[str, Any] = {
            "attendee_count": len(event.required_attendees) + len(event.optional_attendees),
            "required_attendees": event.required_attendees[:MAX_LISTED_ATTENDEES],
            "optional_attendees": event.optional_attendees[:MAX_LISTED_ATTENDEES],
            "body_text": None,
            "body_truncated": None,
        }
        if include_body and event.body_text is not None:
            update["body_text"] = event.body_text[:MAX_LISTED_BODY_CHARS]
            update["body_truncated"] = True if len(event.body_text) > MAX_LISTED_BODY_CHARS else None
        listed.append(event.model_copy(update=update))
    return listed


def _start_key(event: CalendarEvent) -> datetime:
    return event.start or datetime.min.replace(tzinfo=UTC)


def _reason(error: Exception) -> str:
    return str(error)[:MAX_REASON_CHARS]

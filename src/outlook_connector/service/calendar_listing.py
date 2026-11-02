"""Calendar listing: the events in a window, read through Graph's calendar view (research §3.8).

The listing reads the default calendar's occurrences for a window, in time order. A Graph failure is
reported as it is; there is no other service to answer instead.
"""

from __future__ import annotations

import unicodedata
from datetime import UTC, datetime, timedelta
from typing import Any

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import CalendarEvent, CalendarListCoverage, CalendarListResult
from outlook_connector.domain.time import as_utc
from outlook_connector.remote.ports import CalendarViewPage, CalendarViewReader
from outlook_connector.service import cursors

MAX_WINDOW = timedelta(days=366)
MAX_LIMIT = 200
MAX_CONTAINS_CHARS = 200  # per term
MAX_CONTAINS_TERMS = 20
GRAPH_PAGE_SIZE = 100
MAX_GRAPH_PAGES = 10  # pages read per call; a filtered listing may need several to fill its limit
MAX_LISTED_ATTENDEES = 15  # per list (required, optional); attendee_count gives the total
MAX_LISTED_BODY_CHARS = 4000
CURSOR_KIND = "calendar_list"
INVALID_CURSOR = "The calendar cursor is invalid."


class CalendarListing:
    """Lists the signed-in user's calendar events for a window (see the module docstring)."""

    def __init__(self, *, graph_view: CalendarViewReader) -> None:
        self._graph_view = graph_view

    async def list_events(
        self,
        *,
        since: datetime,
        until: datetime,
        contains: list[str] | None = None,
        include_body: bool = False,
        limit: int = 50,
        cursor: str | None = None,
    ) -> CalendarListResult:
        """Entry point: checks the window, the text filter, the limit and the cursor, then lists.

        Events overlap the window: one that started before ``since`` and runs into it is included. With a
        cursor, the window and the filters come from the cursor and the other arguments are ignored.
        ``limit`` is how many matching events to collect before stopping, so a page may hold a few more.
        ``contains`` is a list of terms: an event matches when any term appears in its subject or description.
        A term is matched whole, never split on its spaces.
        """
        if not 1 <= limit <= MAX_LIMIT:
            raise InvalidRequest(f"limit must be between 1 and {MAX_LIMIT}.")
        if cursor is None:
            start, end = _window(since, until)
            return await self._list_graph(start, end, limit, _contains(contains), include_body, link=None)
        state = _decode(cursor)
        return await self._list_graph(
            state["since"], state["until"], limit, state["contains"], state["include_body"], state["link"]
        )

    async def _list_graph(
        self,
        start: datetime,
        end: datetime,
        limit: int,
        terms: list[str] | None,
        include_body: bool,
        link: str | None,
    ) -> CalendarListResult:
        """Graph's view page by page, until ``limit`` events match, the window ends or the budget is spent."""
        needles = [_fold(term) for term in terms] if terms else None
        page_size = GRAPH_PAGE_SIZE if needles else min(limit, GRAPH_PAGE_SIZE)
        matched: list[CalendarEvent] = []
        scanned = 0
        pages = 0
        while True:
            page: CalendarViewPage = await self._graph_view.list_events(
                since=start,
                until=end,
                page_size=page_size,
                page=link,
                include_body=include_body or needles is not None,
            )
            pages += 1
            scanned += len(page.events)
            matched.extend(event for event in page.events if _matches(event, needles))
            link = page.next_link
            if link is None or len(matched) >= limit or pages >= MAX_GRAPH_PAGES:
                break
        complete = link is None
        events = _listed(sorted(matched, key=_start_key), include_body=include_body)
        notes = ["Times are UTC. Events overlap the window; recurring meetings are expanded per occurrence."]
        if needles is not None:
            notes.append(
                f"contains matched each term literally (any term matches), case- and accent-insensitive, in "
                f"each subject and description ({scanned} event(s) read)."
            )
        if any(event.body_truncated for event in events):
            notes.append(
                f"Descriptions are cut to {MAX_LISTED_BODY_CHARS:,} characters; body_truncated marks them."
            )
        if not complete and len(matched) < limit:
            notes.append(
                f"This call stopped at its read budget of {MAX_GRAPH_PAGES} Graph pages ({scanned} event(s) "
                f"read) before {limit} event(s) matched; the window is not finished."
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
                link=link,
                since=start.isoformat(),
                until=end.isoformat(),
                contains=terms,
                include_body=include_body,
            ),
            coverage=CalendarListCoverage(complete=complete, scanned=scanned, notes=notes),
        )


def _window(since: datetime, until: datetime) -> tuple[datetime, datetime]:
    start, end = as_utc(since), as_utc(until)
    if start >= end:
        raise InvalidRequest("since must be earlier than until.")
    if end - start > MAX_WINDOW:
        raise InvalidRequest(f"The window must be at most {MAX_WINDOW.days} days.")
    return start, end


def _contains(value: object) -> list[str] | None:
    """The ``contains`` terms, each stripped and kept whole: a term is never split on its spaces."""
    if value is None:
        return None
    if not isinstance(value, list) or not 1 <= len(value) <= MAX_CONTAINS_TERMS:
        raise InvalidRequest(f"contains must be a list of 1 to {MAX_CONTAINS_TERMS} terms.")
    terms: list[str] = []
    for term in value:
        text = term.strip() if isinstance(term, str) else ""
        if not text:
            raise InvalidRequest("Each contains term must be non-empty text.")
        if len(text) > MAX_CONTAINS_CHARS:
            raise InvalidRequest(f"Each contains term must be at most {MAX_CONTAINS_CHARS} characters.")
        terms.append(text)
    return list(dict.fromkeys(terms))


def _decode(cursor: str) -> dict[str, Any]:
    """The listing state a cursor carries. Cursors are not signed, so every value gets the same checks as
    the entry point again: a malformed or tampered cursor is an InvalidRequest."""
    state = cursors.decode(cursor, CURSOR_KIND)
    since, until = _cursor_time(state.get("since")), _cursor_time(state.get("until"))
    contains, include_body = state.get("contains"), state.get("include_body")
    if since is None or until is None or not isinstance(include_body, bool):
        raise InvalidRequest(INVALID_CURSOR)
    if not isinstance(state.get("link"), str) or not (contains is None or isinstance(contains, list)):
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
        return as_utc(datetime.fromisoformat(value))
    except ValueError:
        return None


def _fold(value: str) -> str:
    """Case- and accent-insensitive text for matching: decomposed, marks removed, casefolded, one-spaced."""
    decomposed = unicodedata.normalize("NFKD", value)
    unmarked = "".join(char for char in decomposed if not unicodedata.combining(char))
    return " ".join(unmarked.casefold().split())


def _matches(event: CalendarEvent, needles: list[str] | None) -> bool:
    """Any term in any field. Each field is matched on its own, so a term cannot span the subject and
    description boundary."""
    if needles is None:
        return True
    fields = [_fold(part) for part in (event.subject, event.body_text) if part]
    return any(needle in field for needle in needles for field in fields)


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

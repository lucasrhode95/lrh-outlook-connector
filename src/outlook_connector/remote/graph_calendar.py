"""Graph calendar view: the default calendar's events, occurrences expanded (research §3.8).

Read-only and Graph first. The token is the calendar profile the Graph client was built with
(``config.CALENDAR_PROFILE``); the Outlook Graph profile has no calendar scope. Times are requested in
UTC, and an answer in any other zone is refused rather than guessed. Ids are converted to the Outlook
Web form, so ``get_calendar_event`` can hydrate them.
"""

from __future__ import annotations

from collections.abc import Iterable
from datetime import UTC, datetime
from typing import Any
from urllib.parse import parse_qs, urlsplit

from outlook_connector.domain.errors import InvalidRequest, Upstream
from outlook_connector.domain.models import CalendarEvent, Recipient
from outlook_connector.remote import ids
from outlook_connector.remote.calendar_mapping import boolean, teams_join_url, text
from outlook_connector.remote.graph import PREFER_TEXT_BODY, Graph
from outlook_connector.remote.ports import CalendarViewPage
from outlook_connector.remote.transport import operation
from outlook_connector.remote.urls import GRAPH_ROOT

VIEW_PATH = "/me/calendarView"
MAX_PAGE_SIZE = 100  # the largest page this module asks for; a continuation link may not ask for more
VIEW_FIELDS = (
    "id,subject,start,end,isAllDay,isCancelled,isOrganizer,location,organizer,attendees,"
    "responseStatus,type,seriesMasterId,onlineMeeting"
)
PREFER_UTC = 'outlook.timezone="UTC"'
# Graph does not return the view in start order on its own, and its pages follow that order. Ordering on
# the server keeps every next link in time order (live check 2026-10-12, research §3.8).
ORDER_BY = "start/dateTime asc"
RESPONSE_NAMES = {  # Graph's responseStatus.response -> the MyResponseType that Outlook Web reports
    "accepted": "Accept",
    "tentativelyAccepted": "Tentative",
    "declined": "Decline",
    "organizer": "Organizer",
    "notResponded": "NoResponseReceived",
    "none": "NoResponseReceived",
}
ITEM_TYPE_NAMES = {  # Graph's event type -> the CalendarItemType that Exchange reports
    "singleInstance": "Single",
    "occurrence": "Occurrence",
    "exception": "Exception",
    "seriesMaster": "RecurringMaster",
}


class GraphCalendarView:
    """Implements ``CalendarViewReader`` over Graph's calendar view."""

    def __init__(self, graph: Graph) -> None:
        self._graph = graph

    async def list_events(
        self,
        *,
        since: datetime,
        until: datetime,
        page_size: int,
        page: str | None,
        include_body: bool,
    ) -> CalendarViewPage:
        """Entry point: a continuation link must be this view for this window; cursors can carry any link."""
        if page is not None and not _continues_window(page, since, until):
            raise InvalidRequest("This cursor does not belong to this calendar window.")
        prefer = (PREFER_UTC, PREFER_TEXT_BODY)
        with operation("listing calendar events"):
            if page is None:
                select = f"{VIEW_FIELDS},body" if include_body else VIEW_FIELDS
                params = {
                    "startDateTime": _iso(since),
                    "endDateTime": _iso(until),
                    "$top": page_size,
                    "$select": select,
                    "$orderby": ORDER_BY,
                }
                items, link = await self._graph.page(VIEW_PATH, params, prefer=prefer)
            else:
                items, link = await self._graph.page(page, prefer=prefer)
        return CalendarViewPage(events=[_event(item) for item in items], next_link=link)


def _iso(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _continues_window(link: str, since: datetime, until: datetime) -> bool:
    """Whether a next link is this view, for this window, with a page size this module would ask for.

    Cursors are not signed, so the link is checked against the window the caller passes in.
    """
    parts = urlsplit(link)
    if f"{parts.scheme}://{parts.netloc}{parts.path}" != f"{GRAPH_ROOT}{VIEW_PATH}":
        return False
    query = parse_qs(parts.query)
    try:
        starts = _query_time(query, "startDateTime")
        ends = _query_time(query, "endDateTime")
        top = int(query.get("$top", [""])[0])
    except ValueError:
        return False
    return starts == since and ends == until and 1 <= top <= MAX_PAGE_SIZE


def _query_time(query: dict[str, list[str]], key: str) -> datetime:
    """A time in a link's query; a missing or unreadable one raises ValueError."""
    return datetime.fromisoformat(query.get(key, [""])[0].replace("Z", "+00:00"))


def _dict(value: Any) -> dict[str, Any]:
    """A nested Graph object, or an empty one when the field is absent."""
    return value if isinstance(value, dict) else {}


def _event(data: dict[str, Any]) -> CalendarEvent:
    """One event, with Graph's field names mapped to the model (only this module knows them)."""
    description = text(_dict(data.get("body")).get("content"))
    online = _dict(data.get("onlineMeeting"))
    status = _dict(data.get("responseStatus"))
    place = _dict(data.get("location"))
    rows = [row for row in data.get("attendees") or [] if isinstance(row, dict)]
    event_id = text(data.get("id"))
    series = text(data.get("seriesMasterId"))
    return CalendarEvent(
        id=ids.to_ows(event_id) if event_id else None,
        immutable_id=event_id,
        subject=text(data.get("subject")),
        start=_graph_time(data.get("start")),
        end=_graph_time(data.get("end")),
        organizer=_recipient(data.get("organizer")),
        required_attendees=_people(row for row in rows if row.get("type") == "required"),
        optional_attendees=_people(row for row in rows if row.get("type") != "required"),
        location=text(place.get("displayName")),
        is_meeting=bool(rows),
        is_cancelled=boolean(data.get("isCancelled")),
        is_all_day=boolean(data.get("isAllDay")),
        is_organizer=boolean(data.get("isOrganizer")),
        my_response_type=RESPONSE_NAMES.get(text(status.get("response")) or ""),
        calendar_item_type=ITEM_TYPE_NAMES.get(text(data.get("type")) or ""),
        series_master_id=ids.to_ows(series) if series else None,
        body_text=description,
        teams_join_url=text(online.get("joinUrl")) or teams_join_url(description),
    )


def _graph_time(value: Any) -> datetime | None:
    """A Graph dateTimeTimeZone. UTC was requested, so another zone is refused rather than placed."""
    if not isinstance(value, dict) or not isinstance(value.get("dateTime"), str):
        return None
    if value.get("timeZone") != "UTC":
        raise Upstream("Graph returned a calendar time outside UTC, although UTC was requested.")
    try:
        stamp = datetime.fromisoformat(value["dateTime"])
    except ValueError:
        raise Upstream("Graph returned an unreadable calendar time.") from None
    return stamp.replace(tzinfo=UTC) if stamp.tzinfo is None else stamp.astimezone(UTC)


def _people(rows: Iterable[dict[str, Any]]) -> list[Recipient]:
    """Attendees in order, skipping any row without a name or an address."""
    return [person for person in map(_recipient, rows) if person is not None]


def _recipient(value: Any) -> Recipient | None:
    """A Graph organizer or attendee: ``emailAddress`` holds the name and the address."""
    if not isinstance(value, dict) or not isinstance(value.get("emailAddress"), dict):
        return None
    email = value["emailAddress"]
    name, address = text(email.get("name")), text(email.get("address"))
    return Recipient(name=name, address=address) if name or address else None

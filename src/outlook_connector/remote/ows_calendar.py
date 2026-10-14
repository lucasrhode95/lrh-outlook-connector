"""Outlook Web Service calendar reads: event hydration for Search v2 pointers, and the calendar window
that answers ``list_events`` when the Graph calendar view fails (research §3.17)."""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

from outlook_connector.domain.errors import Upstream
from outlook_connector.domain.models import CalendarAttachment, CalendarEvent, Recipient
from outlook_connector.remote.calendar_mapping import (
    boolean,
    datetime_value,
    location,
    recipient,
    teams_join_url,
    text,
)
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ports import CalendarItemReader, CalendarWindowReader
from outlook_connector.remote.transport import operation

CALENDAR_FOLDER = {"__type": "DistinguishedFolderId:#Exchange", "Id": "calendar"}  # the default calendar


class OwsCalendarItems(CalendarItemReader):
    def __init__(self, ows: Ows) -> None:
        self._ows = ows

    async def get_event(self, event_id: str) -> CalendarEvent:
        request = {
            "ItemShape": {
                "__type": "ItemResponseShape:#Exchange",
                "BaseShape": "Default",
                "BodyType": "Text",
                "AdditionalProperties": [
                    {"__type": "PropertyUri:#Exchange", "FieldURI": "item:Body"},
                    {"__type": "PropertyUri:#Exchange", "FieldURI": "item:Attachments"},
                ],
            },
            "ItemIds": [{"__type": "ItemId:#Exchange", "Id": event_id}],
        }
        with operation("reading a calendar event"):
            results = await self._ows.call_read("GetItem", request)
        if len(results) != 1:
            raise Upstream("Outlook returned an unexpected number of event results.")
        item = _calendar_item(results[0])
        return _map_event(item, event_id)


class OwsCalendarWindow(CalendarWindowReader):
    """The default calendar's events overlapping a window, from Outlook Web's ``GetCalendarView``.

    The fallback behind ``list_events``. The answer carries no descriptions, so none are returned.
    """

    def __init__(self, ows: Ows) -> None:
        self._ows = ows

    async def list_events(self, *, since: datetime, until: datetime) -> list[CalendarEvent]:
        body = {
            "CalendarId": {"__type": "TargetFolderId:#Exchange", "BaseFolderId": CALENDAR_FOLDER},
            "RangeStart": _ows_time(since),
            "RangeEnd": _ows_time(until),
        }
        with operation("listing calendar events"):
            items = await self._ows.call_view("GetCalendarView", body, time_zone="UTC")
        return [_map_event(item, text(_nested_id(item.get("ItemId")))) for item in items]


def _ows_time(value: datetime) -> str:
    """A naive time in the zone that ``TimeZoneContext`` names (UTC)."""
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.000")


def _calendar_item(result: dict[str, Any]) -> dict[str, Any]:
    items = result.get("Items")
    if not isinstance(items, list) or not items or not isinstance(items[0], dict):
        raise Upstream("Outlook returned no calendar item in GetItem.")
    return items[0]


def _map_event(data: dict[str, Any], event_id: str | None) -> CalendarEvent:
    body_data = data.get("Body")
    body = text(body_data.get("Value")) if isinstance(body_data, dict) else text(body_data)
    body_truncated = boolean(body_data.get("IsTruncated")) if isinstance(body_data, dict) else None
    organizer = _ows_recipient(data.get("Organizer"))
    return CalendarEvent(
        id=event_id,
        immutable_id=text(_nested_id(data.get("ItemId"))),
        subject=text(data.get("Subject")),
        start=datetime_value(data.get("Start")),
        end=datetime_value(data.get("End")),
        organizer=organizer,
        display_to=text(data.get("DisplayTo")),
        required_attendees=_attendees(data.get("RequiredAttendees")),
        optional_attendees=_attendees(data.get("OptionalAttendees")),
        location=location(data.get("Location")),
        is_meeting=boolean(data.get("IsMeeting")),
        is_cancelled=boolean(data.get("IsCancelled")),
        is_all_day=boolean(data.get("IsAllDayEvent")),
        is_organizer=boolean(data.get("IsOrganizer")),
        is_response_requested=boolean(data.get("IsResponseRequested")),
        my_response_type=text(data.get("MyResponseType")),
        sensitivity=text(data.get("Sensitivity")),
        calendar_item_type=text(data.get("CalendarItemType")),
        uid=text(data.get("UID")),
        series_master_id=text(_nested_id(data.get("SeriesMasterItemId"))) or text(data.get("SeriesMasterId")),
        body_text=body,
        body_truncated=body_truncated,
        teams_join_url=teams_join_url(body),
        attachments=_attachments(data.get("Attachments")),
    )


def _nested_id(value: Any) -> Any:
    return value.get("Id") if isinstance(value, dict) else None


def _ows_recipient(value: Any) -> Recipient | None:
    if not isinstance(value, dict):
        return recipient(value)
    mailbox = value.get("Mailbox")
    return recipient(mailbox if isinstance(mailbox, dict) else value)


def _attendees(value: Any) -> list[Recipient]:
    if not isinstance(value, list):
        return []
    attendees: list[Recipient] = []
    for row in value:
        recipient = _ows_recipient(row)
        if recipient is not None:
            attendees.append(recipient)
    return attendees


def _attachments(value: Any) -> list[CalendarAttachment]:
    if not isinstance(value, list):
        return []
    attachments: list[CalendarAttachment] = []
    for row in value:
        if not isinstance(row, dict):
            continue
        item_id = _nested_id(row.get("AttachmentId"))
        size = row.get("Size")
        is_inline = row.get("IsInline")
        attachments.append(
            CalendarAttachment(
                id=text(item_id),
                name=text(row.get("Name")),
                content_type=text(row.get("ContentType")),
                size=size if isinstance(size, int) else None,
                is_inline=is_inline if isinstance(is_inline, bool) else False,
            )
        )
    return attachments

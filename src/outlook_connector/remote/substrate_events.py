"""Exchange event discovery through the standalone-proven Outlook Search v2 route."""

from __future__ import annotations

from typing import Any

from outlook_connector.domain.models import CalendarEvent, Recipient
from outlook_connector.remote.calendar_mapping import boolean, datetime_value, location, recipient, text
from outlook_connector.remote.ports import CalendarSearchPage
from outlook_connector.remote.transport import Transport, operation

SEARCH_URL = "https://outlook.office.com/searchservice/api/v2/query"
SEARCH_PROFILE = "m365.search"


class SubstrateEventSearch:
    def __init__(self, transport: Transport, *, anchor_mailbox: str) -> None:
        self._transport = transport
        self._anchor_mailbox = anchor_mailbox

    async def search_events(
        self,
        query: str,
        *,
        offset: int,
        page_size: int,
        cvid: str,
        logical_id: str,
    ) -> CalendarSearchPage:
        body = {
            "Cvid": cvid,
            "LogicalId": logical_id,
            "Scenario": {"Name": "owa.react"},
            "TimeZone": "UTC",
            "TextDecorations": "Off",
            "EntityRequests": [
                {
                    "EntityType": "Event",
                    "ContentSources": ["Exchange"],
                    "Query": {"QueryString": query, "DisplayQueryString": query},
                    "From": offset,
                    "Size": page_size,
                    "Sort": [],
                    "PropertySet": "Optimized",
                }
            ],
        }
        with operation("searching calendar events"):
            data = await self._transport.json(
                "POST",
                SEARCH_URL,
                profile=SEARCH_PROFILE,
                headers={"X-AnchorMailbox": self._anchor_mailbox},
                json_body=body,
                retry=True,
            )
        rows, more, found = _event_rows(data)
        events = [_search_event(row) for row in rows]
        return CalendarSearchPage(
            events=events,
            rows_received=len(rows),
            more_results_available=more,
            result_set_found=found,
        )


def _event_rows(data: Any) -> tuple[list[dict[str, Any]], bool | None, bool]:
    if not isinstance(data, dict) or not isinstance(data.get("EntitySets"), list):
        return [], None, False
    rows: list[dict[str, Any]] = []
    flags: list[bool] = []
    found = False
    for entity_set in data["EntitySets"]:
        if not isinstance(entity_set, dict) or not isinstance(entity_set.get("ResultSets"), list):
            continue
        for result_set in entity_set["ResultSets"]:
            if not isinstance(result_set, dict):
                continue
            results = result_set.get("Results")
            if isinstance(results, list):
                found = True
                rows.extend(row for row in results if isinstance(row, dict))
            more = result_set.get("MoreResultsAvailable", entity_set.get("MoreResultsAvailable"))
            if isinstance(more, bool):
                flags.append(more)
    return rows, any(flags) if flags else None, found


def _search_event(row: dict[str, Any]) -> CalendarEvent:
    source = row.get("Source")
    source = source if isinstance(source, dict) else {}
    item_id = source.get("ItemId")
    item_id = item_id.get("Id") if isinstance(item_id, dict) else None
    organizer = recipient(source.get("Organizer"))
    if organizer is None and (name := text(source.get("OrganizerName"))):
        organizer = Recipient(name=name, address=text(source.get("OrganizerAddress")))
    return CalendarEvent(
        id=text(item_id) or text(source.get("Id")),
        immutable_id=text(source.get("ImmutableId")),
        subject=text(source.get("Subject")),
        start=datetime_value(source.get("Start")),
        end=datetime_value(source.get("End")),
        organizer=organizer,
        display_to=text(source.get("DisplayTo")),
        location=location(source.get("Location")),
        is_meeting=boolean(source.get("IsMeeting")),
        is_cancelled=boolean(source.get("IsCancelled")),
        is_all_day=boolean(source.get("IsAllDayEvent")),
        is_organizer=boolean(source.get("IsOrganizer")),
        is_response_requested=boolean(source.get("IsResponseRequested")),
        my_response_type=text(source.get("MyResponseType")),
        sensitivity=text(source.get("Sensitivity")),
        calendar_item_type=text(source.get("CalendarItemType")),
        uid=text(source.get("UID")),
        series_master_id=text(source.get("SeriesMasterId")) or text(source.get("SeriesMasterItemId")),
    )

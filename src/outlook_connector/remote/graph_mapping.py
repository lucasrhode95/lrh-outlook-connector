"""Graph JSON → domain models. The only module that knows Graph mail field names."""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any, Literal, cast

from outlook_connector.domain.models import (
    Attachment,
    Folder,
    Meeting,
    MeetingKind,
    Message,
    MessageSummary,
    Recipient,
)

SUMMARY_FIELDS = (
    "id,conversationId,parentFolderId,subject,from,sender,toRecipients,ccRecipients,receivedDateTime,"
    "sentDateTime,isRead,isDraft,hasAttachments,importance,categories,flag,bodyPreview,internetMessageId,"
    # meeting mail: typed fields of Graph's eventMessage subtypes, absent on ordinary mail
    "microsoft.graph.eventMessage/meetingMessageType,microsoft.graph.eventMessage/startDateTime,"
    "microsoft.graph.eventMessage/endDateTime,microsoft.graph.eventMessage/location,"
    "microsoft.graph.eventMessage/isAllDay,microsoft.graph.eventMessage/isOutOfDate,"
    "microsoft.graph.eventMessageRequest/meetingRequestType"
)
MEETING_KINDS: dict[str, MeetingKind] = {
    "meetingRequest": "invite",
    "meetingCancelled": "cancelled",
    "meetingAccepted": "accepted",
    "meetingTenativelyAccepted": "tentative",  # sic: Graph's spelling
    "meetingDeclined": "declined",
}
NEW_REQUEST = (None, "none", "newMeetingRequest")
MESSAGE_FIELDS = SUMMARY_FIELDS + ",bccRecipients,body,uniqueBody"
FOLDER_FIELDS = "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden"
ATTACHMENT_FIELDS = "id,name,contentType,size,isInline"


def parse_dt(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def event_time(value: Any) -> datetime | None:
    """A Graph dateTimeTimeZone. Graph sends UTC unless asked otherwise; another zone is not guessed."""
    if not isinstance(value, dict) or value.get("timeZone") != "UTC":
        return None
    stamp = parse_dt(str(value.get("dateTime") or "")[:19])  # Graph sends 7 fraction digits
    return stamp.replace(tzinfo=UTC) if stamp else None


def meeting(data: dict[str, Any]) -> Meeting | None:
    kind = MEETING_KINDS.get(str(data.get("meetingMessageType")))
    if kind is None:
        return None
    if kind == "invite" and data.get("meetingRequestType") not in NEW_REQUEST:
        kind = "update"
    place = data.get("location")
    return Meeting(
        kind=kind,
        start=event_time(data.get("startDateTime")),
        end=event_time(data.get("endDateTime")),
        all_day=bool(data.get("isAllDay")),
        location=(place.get("displayName") or None) if isinstance(place, dict) else None,
        out_of_date=bool(data.get("isOutOfDate")),
    )


def recipient(value: Any) -> Recipient | None:
    email = value.get("emailAddress") if isinstance(value, dict) else None
    if not isinstance(email, dict):
        return None
    return Recipient(name=email.get("name") or None, address=email.get("address") or None)


def recipients(values: Any) -> list[Recipient]:
    return [r for r in (recipient(v) for v in values or []) if r is not None]


def folder(data: dict[str, Any], well_known: str | None = None) -> Folder:
    return Folder(
        id=data["id"],
        parent_id=data.get("parentFolderId"),
        name=data.get("displayName") or "(unnamed)",
        well_known=well_known,
        total=data.get("totalItemCount"),
        unread=data.get("unreadItemCount"),
        child_count=data.get("childFolderCount"),
        hidden=bool(data.get("isHidden")),
    )


def _summary_fields(data: dict[str, Any]) -> dict[str, Any]:
    flag: dict[str, Any] = data["flag"] if isinstance(data.get("flag"), dict) else {}
    return {
        "id": data["id"],
        "conversation_id": data.get("conversationId"),
        "folder_id": data.get("parentFolderId"),
        "subject": data.get("subject"),
        "sender": recipient(data.get("from")) or recipient(data.get("sender")),
        "to": recipients(data.get("toRecipients")),
        "cc": recipients(data.get("ccRecipients")),
        "received_at": parse_dt(data.get("receivedDateTime")),
        "sent_at": parse_dt(data.get("sentDateTime")),
        "is_read": data.get("isRead"),
        "is_draft": data.get("isDraft"),
        "has_attachments": bool(data.get("hasAttachments")),
        "importance": data.get("importance"),
        "categories": list(data.get("categories") or []),
        "flagged": flag.get("flagStatus") == "flagged",
        "preview": data.get("bodyPreview"),
        "internet_message_id": data.get("internetMessageId"),
        "meeting": meeting(data),
    }


def summary(data: dict[str, Any]) -> MessageSummary:
    return MessageSummary(**_summary_fields(data))


def _content(value: Any) -> str | None:
    return value.get("content") if isinstance(value, dict) else None


def message(data: dict[str, Any], *, html: bool) -> Message:
    """A full message. ``html`` says which body format the request asked for."""
    body, unique = _content(data.get("body")), _content(data.get("uniqueBody"))
    fields: dict[str, Any] = _summary_fields(data) | {"bcc": recipients(data.get("bccRecipients"))}
    if html:
        return Message(**fields, body_html=body, unique_body_html=unique)
    return Message(**fields, body_text=body, unique_body_text=unique)


def attachment(data: dict[str, Any], message_id: str) -> Attachment:
    kind = str(data.get("@odata.type", "")).rsplit(".", 1)[-1]
    return Attachment(
        id=data["id"],
        message_id=message_id,
        name=data.get("name"),
        content_type=data.get("contentType"),
        size=data.get("size"),
        is_inline=bool(data.get("isInline")),
        kind=cast(
            Literal["file", "item", "reference", "unknown"],
            {"fileAttachment": "file", "itemAttachment": "item", "referenceAttachment": "reference"}.get(
                kind, "unknown"
            ),
        ),
    )

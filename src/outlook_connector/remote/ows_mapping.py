"""Domain values → OWS request JSON, and the few answer fields the writer reads (research §4.2).

Pure functions, no I/O: ``ows_mail`` sends what they build through ``ows.Ows``.
"""

from __future__ import annotations

from typing import Any

from outlook_connector.domain.errors import WriteOutcomeUnknown
from outlook_connector.domain.models import DraftMessage
from outlook_connector.remote import ids
from outlook_connector.remote.ows import outcome
from outlook_connector.remote.ports import FolderTarget


def address(address: str) -> dict[str, str]:
    return {"__type": "EmailAddress:#Exchange", "EmailAddress": address, "RoutingType": "SMTP"}


def html_body(page: str) -> dict[str, str]:
    """Assumes (not re-checked here): validated intentional HTML from the service."""
    return {"__type": "BodyContentType:#Exchange", "BodyType": "HTML", "Value": page}


def created_id(items: list[dict[str, Any]]) -> str | None:
    inner = items[0].get("Items")
    first = inner[0] if isinstance(inner, list) and inner and isinstance(inner[0], dict) else {}
    item_id = first.get("ItemId")
    value = item_id.get("Id") if isinstance(item_id, dict) else None
    return ids.to_graph(value) if isinstance(value, str) and value else None


def folder(target: FolderTarget) -> dict[str, str]:
    if target.well_known:
        return {"__type": "DistinguishedFolderId:#Exchange", "Id": target.well_known}
    return {"__type": "FolderId:#Exchange", "Id": ids.to_ows(target.folder_id)}


def item_id(message_id: str) -> dict[str, str]:
    return {"__type": "ItemId:#Exchange", "Id": ids.to_ows(message_id)}


def per_id(message_ids: list[str], items: list[dict[str, Any]]) -> dict[str, str | None]:
    if len(items) != len(message_ids):  # the change may have been made: read back, never repeat
        raise WriteOutcomeUnknown(
            f"Outlook answered {len(items)} item results for {len(message_ids)} messages; it was not retried."
        )
    return {mid: outcome(item) for mid, item in zip(message_ids, items, strict=True)}


def create(message: DraftMessage) -> dict[str, Any]:
    """``CreateItem`` / ``SaveOnly`` leaves a draft in Drafts (research §4.2).

    Replies use EWS's ``ReplyToItem`` / ``ReplyAllToItem`` response objects, which append the
    quoted original; recipients and subject are passed explicitly, so the result matches the
    proposal. A reply's own text goes in an HTML body, so the quoted original keeps its formatting
    and inline images (live 2026-10-04). New messages use the same HTML engine.

    Assumes (not re-checked here): ``message`` comes from ``Writes._resolve``.
    """
    disposition = "SaveOnly"
    recipients = {
        "ToRecipients": [address(a) for a in message.to],
        "CcRecipients": [address(a) for a in message.cc],
        "BccRecipients": [address(a) for a in message.bcc],
    }
    if message.reply_to_message_id:
        kind = "ReplyAllToItem" if message.reply_all else "ReplyToItem"
        item: dict[str, Any] = {
            "__type": f"{kind}:#Exchange",
            "ReferenceItemId": {"__type": "ItemId:#Exchange", "Id": ids.to_ows(message.reply_to_message_id)},
            "NewBodyContent": html_body(message.html_body),
            "Subject": message.subject,
            **recipients,
        }
        compose = "replyAll" if message.reply_all else "reply"
    else:
        item = {
            "__type": "Message:#Exchange",
            "Body": html_body(message.html_body),
            "From": {},
            **recipients,
            "Subject": message.subject,
            "Importance": "Normal",
            "IsDeliveryReceiptRequested": False,
            "IsReadReceiptRequested": False,
            "IsSendIndividually": False,
            "MessageDisposition": disposition,
            "ShouldIgnoreChangeKey": True,
            "operation": "New",
        }
        compose = "newMail"
    return {
        "ClientSupportsIrm": True,
        "ComposeOperation": compose,
        "MessageDisposition": disposition,
        "Items": [item],
        "TimeFormat": "",
        "SendOnNotFoundError": True,
        "RemoteExecute": True,
        "ShouldSuppressReadReceipt": True,
        "OutboundCharset": "AutoDetect",
        "ItemShape": {
            "__type": "ItemResponseShape:#Exchange",
            "BaseShape": "IdOnly",
            "ClientSupportsIrm": True,
            "AdditionalProperties": [{"__type": "PropertyUri:#Exchange", "FieldURI": "ItemLastModifiedTime"}],
        },
        "ShapeName": "MailCompose",
    }

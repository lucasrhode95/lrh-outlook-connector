"""``MailWriter`` over Outlook Web (OWS): drafts, send and mailbox changes (architecture §5.5)."""

from __future__ import annotations

from typing import Any

from outlook_connector.domain.models import DraftMessage
from outlook_connector.remote import ows_mapping as mapping
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ports import FolderTarget, MailWriter
from outlook_connector.remote.transport import operation


class OwsMailWriter(MailWriter):
    """Implements ``MailWriter`` over OWS.

    Assumes (not re-checked here): draft messages come from ``Writes._resolve`` (addresses, subject and body
    validated), message ids are Graph immutable ids from this connector, folder targets were resolved by
    the service, and the write account was checked (``Writes.check_account``). Nothing is re-validated
    here; every call is sent once.
    """

    def __init__(self, ows: Ows) -> None:
        self._ows = ows

    def account(self) -> dict[str, Any]:
        return self._ows.account()

    async def create_draft(self, message: DraftMessage) -> str | None:
        """Save into Drafts (never sends). Returns the draft's Graph id when Outlook reports it."""
        with operation("saving a draft"):
            return mapping.created_id(await self._ows.call("CreateItem", mapping.create(message)))

    async def edit_draft(self, draft_id: str, changes: dict[str, Any]) -> None:
        """Assumes (not re-checked here): existing draft, validated partial fields, bound account."""
        fields = {
            "subject": ("item:Subject", "Subject"),
            "html_body": ("item:Body", "Body"),
            "to": ("message:ToRecipients", "ToRecipients"),
            "cc": ("message:CcRecipients", "CcRecipients"),
            "bcc": ("message:BccRecipients", "BccRecipients"),
        }
        updates = []
        for key, value in changes.items():
            uri, prop = fields[key]
            if key == "html_body":
                value = mapping.html_body(value)
            elif key in ("to", "cc", "bcc"):
                value = [mapping.address(a) for a in value]
            updates.append(
                {
                    "__type": "SetItemField:#Exchange",
                    "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": uri},
                    "Item": {"__type": "Message:#Exchange", prop: value},
                }
            )
        with operation("editing a draft"):
            await self._ows.call("UpdateItem", _draft_update(draft_id, updates, "SaveOnly"))

    async def send_draft(self, draft_id: str, revision: str) -> None:
        """Send only this existing draft, unchanged: ``UpdateItem`` with ``SendAndSaveCopy`` and no
        field updates (``SendItem`` is not supported over OWS). The item id carries the change key
        the draft was read at, with ``NeverOverwrite``: if the draft changed since (edited in Outlook
        meanwhile), Outlook refuses with ``ErrorIrresolvableConflict`` and sends nothing. Both proven
        live 2026-10-06. Sent once, never retried.

        Assumes (not re-checked here): ``revision`` is the draft's version as ``Writes`` read it.
        """
        body = _draft_update(draft_id, [], "SendAndSaveCopy")
        body["ItemChanges"][0]["ItemId"]["ChangeKey"] = revision
        body["ConflictResolution"] = "NeverOverwrite"
        body["SavedItemFolderId"] = {
            "__type": "TargetFolderId:#Exchange",
            "BaseFolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "sentitems"},
        }
        with operation("sending a draft"):
            await self._ows.call("UpdateItem", body)

    # ---------------------------------------------------------------- mutations (research §4.2)
    # Each returns {message id: None when done, else Outlook's response code}, one request per call
    # (the service keeps calls small). Ids are Graph immutable ids; they survive moves.

    async def set_read(self, message_ids: list[str], is_read: bool) -> dict[str, str | None]:
        with operation("changing read state"):
            changes = {mid: ("message:IsRead", {"IsRead": is_read}) for mid in message_ids}
            return await self._update(changes)

    async def set_flag(self, message_ids: list[str], flagged: bool) -> dict[str, str | None]:
        status = "Flagged" if flagged else "NotFlagged"
        flag = {"Flag": {"__type": "FlagType:#Exchange", "FlagStatus": status}}
        with operation("changing flags"):
            return await self._update({mid: ("item:Flag", flag) for mid in message_ids})

    async def move(self, message_ids: list[str], folder: FolderTarget) -> dict[str, str | None]:
        body = {
            "ToFolderId": {"__type": "TargetFolderId:#Exchange", "BaseFolderId": mapping.folder(folder)},
            "ItemIds": [mapping.item_id(mid) for mid in message_ids],
            "ReturnNewItemIds": True,
        }
        with operation("moving messages"):
            return mapping.per_id(message_ids, await self._ows.call("MoveItem", body, strict=False))

    async def delete(self, message_ids: list[str]) -> dict[str, str | None]:
        """Move to Deleted Items (``MoveToDeletedItems``). There is no hard delete."""
        body = {
            "ItemIds": [mapping.item_id(mid) for mid in message_ids],
            "DeleteType": "MoveToDeletedItems",
            "SendMeetingCancellations": "SendToNone",
            "AffectedTaskOccurrences": "AllOccurrences",
            "SuppressReadReceipts": True,
        }
        with operation("deleting messages"):
            return mapping.per_id(message_ids, await self._ows.call("DeleteItem", body, strict=False))

    async def _update(self, changes: dict[str, tuple[str, dict[str, Any]]]) -> dict[str, str | None]:
        """``UpdateItem`` with one ``SetItemField`` per message, as proven (research §4.2)."""
        body = {
            "ItemChanges": [
                {
                    "__type": "ItemChange:#Exchange",
                    "ItemId": mapping.item_id(mid),
                    "Updates": [
                        {
                            "__type": "SetItemField:#Exchange",
                            "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": field_uri},
                            "Item": {"__type": "Message:#Exchange", **props},
                        }
                    ],
                }
                for mid, (field_uri, props) in changes.items()
            ],
            "ConflictResolution": "AlwaysOverwrite",
            "MessageDisposition": "SaveOnly",
            "SuppressReadReceipts": True,
            "SendCalendarInvitationsOrCancellations": "SendToNone",
        }
        return mapping.per_id(list(changes), await self._ows.call("UpdateItem", body, strict=False))


def _draft_update(draft_id: str, updates: list[dict[str, Any]], disposition: str) -> dict[str, Any]:
    """Assumes (not re-checked here): id and updates validated by Writes."""
    return {
        "ItemChanges": [
            {"__type": "ItemChange:#Exchange", "ItemId": mapping.item_id(draft_id), "Updates": updates}
        ],
        "ConflictResolution": "AlwaysOverwrite",
        "MessageDisposition": disposition,
        "SuppressReadReceipts": True,
        "SendCalendarInvitationsOrCancellations": "SendToNone",
    }

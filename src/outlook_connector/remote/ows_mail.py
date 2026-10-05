"""``MailWriter`` over Outlook Web (OWS): drafts, send and mailbox changes (architecture §5.5)."""

from __future__ import annotations

from typing import Any

from outlook_connector.domain.models import EmailProposal
from outlook_connector.remote import ows_mapping as mapping
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ports import FolderTarget, MailWriter
from outlook_connector.remote.transport import operation


class OwsMailWriter(MailWriter):
    """Implements ``MailWriter`` over OWS.

    Assumes (not re-checked here): proposals come from ``Writes.propose`` (addresses, subject and body
    validated), message ids are Graph immutable ids from this connector, folder targets were resolved by
    the service, and the write account was checked (``Writes.check_account``). Nothing is re-validated
    here; every call is sent once.
    """

    def __init__(self, ows: Ows) -> None:
        self._ows = ows

    def account(self) -> dict[str, Any]:
        return self._ows.account()

    async def create_draft(self, message: EmailProposal) -> str | None:
        """Save into Drafts (never sends). Returns the draft's Graph id when Outlook reports it."""
        with operation("saving a draft"):
            return mapping.created_id(await self._ows.call("CreateItem", mapping.create(message, "SaveOnly")))

    async def send(self, message: EmailProposal) -> None:
        """Send once and keep a copy in Sent Items. Never retried."""
        with operation("sending a message"):
            await self._ows.call("CreateItem", mapping.create(message, "SendAndSaveCopy"))

    async def send_draft(self, draft_id: str, subject: str) -> None:
        """Send an existing draft as it is, the way Outlook Web does: ``UpdateItem`` with
        ``SendAndSaveCopy`` (``SendItem`` is not supported over OWS; live 2026-10-04). The update
        sets the subject the draft already has. Sent once, never retried."""
        body = {
            "ItemChanges": [
                {
                    "__type": "ItemChange:#Exchange",
                    "ItemId": mapping.item_id(draft_id),
                    "Updates": [
                        {
                            "__type": "SetItemField:#Exchange",
                            "Path": {"__type": "PropertyUri:#Exchange", "FieldURI": "item:Subject"},
                            "Item": {"__type": "Message:#Exchange", "Subject": subject},
                        }
                    ],
                }
            ],
            "ConflictResolution": "AlwaysOverwrite",
            "MessageDisposition": "SendAndSaveCopy",
            "SavedItemFolderId": {
                "__type": "TargetFolderId:#Exchange",
                "BaseFolderId": {"__type": "DistinguishedFolderId:#Exchange", "Id": "sentitems"},
            },
            "SuppressReadReceipts": True,
            "SendCalendarInvitationsOrCancellations": "SendToNone",
        }
        with operation("sending a reply"):
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

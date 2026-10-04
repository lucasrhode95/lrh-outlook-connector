"""Mailbox changes on explicit message ids (requirements v4 §11.2, architecture §5.8).

The user's MCP client prompt is the authorization: tools take explicit ids and an explicit target,
never a query. Every action:

1. reads each message's current state through Graph (one batch): unknown ids are ``not_found``,
   out-of-reach ones ``failed``, and those already in the wanted state ``unchanged`` (nothing is
   sent for them);
2. sends the change once per chunk, with a result per message;
3. on an unclear answer, reads the messages again and reports ``done`` only where the change is
   visible, ``unknown`` elsewhere. Nothing is retried.

Deleting moves to Deleted Items; a message already in Deleted Items is left alone (deleting it
there would remove it from the folder view), so there is never a hard delete.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Awaitable, Callable

from outlook_connector.domain.errors import InvalidRequest, WriteOutcomeUnknown
from outlook_connector.domain.models import (
    MAX_MUTATION_ITEMS,
    ItemResult,
    MessageSummary,
    MutationResult,
)
from outlook_connector.remote.ports import FolderTarget, MailWriter
from outlook_connector.service.mailbox import Mailbox

CHUNK = 20  # messages per OWS request
Wanted = Callable[[MessageSummary], bool]  # is the message already as wanted?
Send = Callable[[list[str]], Awaitable[dict[str, str | None]]]


class Mutations:
    def __init__(self, mailbox: Mailbox, writer: MailWriter, check_account: Callable[[], None]) -> None:
        self.mailbox = mailbox
        self.writer = writer
        self.check_account = check_account

    async def set_read(
        self,
        message_ids: list[str],
        is_read: bool,
        *,
        conversation_ids: list[str] | None = None,
        include_deleted_items: bool = False,
    ) -> MutationResult:
        """Mark messages, and every message of the given conversations in scope, read or unread."""
        ids = list(message_ids)
        for conversation_id in conversation_ids or []:
            ids += await self._conversation(conversation_id, include_deleted_items=include_deleted_items)
        return await self._apply(
            "read" if is_read else "unread",
            ids,
            lambda m: m.is_read is is_read,
            lambda chunk: self.writer.set_read(chunk, is_read),
        )

    async def set_flag(self, message_ids: list[str], flagged: bool) -> MutationResult:
        return await self._apply(
            "flag" if flagged else "unflag",
            message_ids,
            lambda m: m.flagged is flagged,
            lambda chunk: self.writer.set_flag(chunk, flagged),
        )

    async def move(self, message_ids: list[str], folder: str) -> MutationResult:
        target = await self.mailbox.resolve_folder(folder)  # hidden folders are refused
        if await self.mailbox.under(target.id, "deleteditems"):
            raise InvalidRequest("To delete messages, use delete_messages (it moves them to Deleted Items).")
        destination = FolderTarget(target.id, target.well_known)
        return await self._apply(
            f"move to {target.path or target.name}",
            message_ids,
            lambda m: m.folder_id == target.id,
            lambda chunk: self.writer.move(chunk, destination),
        )

    async def delete(self, message_ids: list[str]) -> MutationResult:
        """Move to Deleted Items. Messages already in Deleted Items (or its subfolders) are left alone."""
        deleted: set[str] = set()

        async def classify() -> None:  # after _apply has refreshed the folder list for these messages
            folders = await self.mailbox.folder_map()
            for folder_id in folders:
                if await self.mailbox.under(folder_id, "deleteditems"):
                    deleted.add(folder_id)

        return await self._apply(
            "delete (move to Deleted Items)",
            message_ids,
            lambda m: m.folder_id in deleted,
            self.writer.delete,
            prepare=classify,
        )

    # ---------------------------------------------------------------- shared flow

    async def _conversation(self, conversation_id: str, *, include_deleted_items: bool) -> list[str]:
        """Every message of the conversation in scope (all copies, not merged).

        Assumes (not re-checked here): ``conversation_id`` is taken as given (from this connector's own
        results).
        """
        items, _ = await self.mailbox.reader.conversation(conversation_id)
        skip = await self.mailbox.exclusions(include_deleted_items=include_deleted_items)
        folders, hidden = await self.mailbox.reach(m.folder_id for m in items)
        return [
            m.id
            for m in items
            if m.folder_id in folders and m.folder_id not in hidden and m.folder_id not in skip
        ]

    async def _apply(
        self,
        action: str,
        message_ids: list[str],
        already: Wanted,
        send: Send,
        *,
        prepare: Callable[[], Awaitable[None]] | None = None,
    ) -> MutationResult:
        """Read state, send the change once per chunk, report a result per message.

        Entry point for every mutation: validates the ids (at least one, at most MAX_MUTATION_ITEMS,
        duplicates dropped), checks the account, and classifies each message before anything is sent. The
        writer trusts the chunks it is given.
        """
        ids = list(dict.fromkeys(message_ids))
        if not ids:
            raise InvalidRequest("Name at least one message id.")
        if len(ids) > MAX_MUTATION_ITEMS:
            raise InvalidRequest(f"At most {MAX_MUTATION_ITEMS} messages per call.")
        self.check_account()
        before = await self.mailbox.reader.get_summaries(ids)
        folders, hidden = await self.mailbox.reach(m.folder_id for m in before.summaries.values() if m)
        if prepare:
            await prepare()
        results: dict[str, ItemResult] = {}
        pending: list[str] = []
        for mid in ids:
            summary = before.summaries.get(mid)
            if mid in before.failed:
                results[mid] = ItemResult(id=mid, status="failed", detail=before.failed[mid])
            elif summary is None:
                results[mid] = ItemResult(id=mid, status="not_found", detail="Not on the server.")
            elif summary.folder_id not in folders or summary.folder_id in hidden:
                results[mid] = ItemResult(id=mid, status="failed", detail="Out of reach (hidden folder).")
            elif already(summary):
                results[mid] = ItemResult(id=mid, status="unchanged")
            else:
                pending.append(mid)
        for start in range(0, len(pending), CHUNK):
            chunk = pending[start : start + CHUNK]
            try:
                outcomes = await send(chunk)
            except WriteOutcomeUnknown as exc:
                results |= await self._recheck(chunk, already, str(exc))
                continue
            for mid in chunk:
                code = outcomes.get(mid)
                if code is None:
                    results[mid] = ItemResult(id=mid, status="done")
                elif code == "ErrorItemNotFound":
                    results[mid] = ItemResult(id=mid, status="not_found", detail=code)
                else:
                    results[mid] = ItemResult(id=mid, status="failed", detail=code)
        ordered = [results[mid] for mid in ids]
        return MutationResult(action=action, results=ordered, counts=dict(Counter(r.status for r in ordered)))

    async def _recheck(self, chunk: list[str], already: Wanted, reason: str) -> dict[str, ItemResult]:
        """After an unclear answer: done where the change is visible, unknown elsewhere.

        Assumes (not re-checked here): ``chunk`` is one ``_apply`` already sent, with no clear answer.
        """
        after = await self.mailbox.reader.get_summaries(chunk)
        out = {}
        for mid in chunk:
            summary = after.summaries.get(mid)
            if summary is not None and already(summary):
                out[mid] = ItemResult(id=mid, status="done", detail="Confirmed by reading it back.")
            else:
                out[mid] = ItemResult(id=mid, status="unknown", detail=reason)
        return out

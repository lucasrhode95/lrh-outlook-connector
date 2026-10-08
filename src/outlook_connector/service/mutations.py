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

from outlook_connector.domain.errors import ConnectorError, InvalidRequest, WriteOutcomeUnknown
from outlook_connector.domain.models import (
    DEFAULT_SCOPE,
    MAX_MUTATION_ITEMS,
    ItemResult,
    MessageSummary,
    MutationResult,
    Scope,
)
from outlook_connector.remote.ports import (
    MAX_CONCURRENT_REQUESTS,
    FetchedSummaries,
    FolderTarget,
    MailWriter,
)
from outlook_connector.service.concurrency import gather_cancel_on_error
from outlook_connector.service.failures import describe
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.scope import validate_scope

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
        scope: Scope = DEFAULT_SCOPE,
        continue_on_error: bool = True,
    ) -> MutationResult:
        """Entry point: bound explicit ids to 100; expand each conversation under its server limit."""
        validate_scope(scope, sent_items=False, meeting_mail=False, deleted_items=bool(conversation_ids))
        explicit = _message_ids(message_ids, allow_empty=bool(conversation_ids))
        ids = list(explicit)
        notes = []
        conversation_summaries: dict[str, MessageSummary] = {}
        selected_conversations = list(dict.fromkeys(conversation_ids or []))
        for start in range(0, len(selected_conversations), MAX_CONCURRENT_REQUESTS):
            batch = selected_conversations[start : start + MAX_CONCURRENT_REQUESTS]
            expanded_conversations = await gather_cancel_on_error(
                *(self._conversation(conversation_id, scope=scope) for conversation_id in batch)
            )
            for conversation_id, (expanded, truncated) in zip(batch, expanded_conversations, strict=True):
                ids.extend(summary.id for summary in expanded)
                conversation_summaries.update({summary.id: summary for summary in expanded})
                if truncated:
                    notes.append(
                        f"Conversation {conversation_id} was truncated at the 1,000-message listing limit; "
                        "only listed messages in scope are included."
                    )
        ids = list(dict.fromkeys(ids))
        if not ids:
            raise InvalidRequest("Name at least one message id or a conversation with messages in scope.")
        result = await self._apply(
            "read" if is_read else "unread",
            ids,
            lambda m: m.is_read is is_read,
            lambda chunk: self.writer.set_read(chunk, is_read),
            continue_on_error=continue_on_error,
            known=conversation_summaries,
        )
        if conversation_ids and len(ids) > MAX_MUTATION_ITEMS:
            result.results = [
                r for r in result.results if r.id in explicit or r.status not in ("done", "unchanged")
            ]
            notes.append(
                "Ordinary conversation-expanded done/unchanged results are summarized in counts; "
                "explicit-id and error results are retained."
            )
        result.notes = notes
        return result

    async def set_flag(
        self, message_ids: list[str], flagged: bool, *, continue_on_error: bool = True
    ) -> MutationResult:
        """Entry point: validate explicit ids (1–100 unique messages) before changing flags."""
        return await self._apply(
            "flag" if flagged else "unflag",
            _message_ids(message_ids),
            lambda m: m.flagged is flagged,
            lambda chunk: self.writer.set_flag(chunk, flagged),
            continue_on_error=continue_on_error,
        )

    async def move(
        self, message_ids: list[str], folder: str, *, continue_on_error: bool = True
    ) -> MutationResult:
        """Entry point: validate explicit ids and resolve a permitted destination."""
        message_ids = _message_ids(message_ids)
        target = await self.mailbox.resolve_folder(folder)  # hidden folders are refused
        if await self.mailbox.under(target.id, "deleteditems"):
            raise InvalidRequest("To delete messages, use delete_messages (it moves them to Deleted Items).")
        destination = FolderTarget(target.id, target.well_known)
        return await self._apply(
            f"move to {target.path or target.name}",
            message_ids,
            lambda m: m.folder_id == target.id,
            lambda chunk: self.writer.move(chunk, destination),
            continue_on_error=continue_on_error,
        )

    async def delete(self, message_ids: list[str], *, continue_on_error: bool = True) -> MutationResult:
        """Entry point: validate explicit ids; move to Deleted Items, leaving messages already there alone."""
        deleted: set[str] = set()

        async def classify() -> None:  # after _apply has refreshed the folder list for these messages
            folders = await self.mailbox.folder_map()
            for folder_id in folders:
                if await self.mailbox.under(folder_id, "deleteditems"):
                    deleted.add(folder_id)

        return await self._apply(
            "delete (move to Deleted Items)",
            _message_ids(message_ids),
            lambda m: m.folder_id in deleted,
            self.writer.delete,
            prepare=classify,
            continue_on_error=continue_on_error,
        )

    # ---------------------------------------------------------------- shared flow

    async def _conversation(self, conversation_id: str, *, scope: Scope) -> tuple[list[MessageSummary], bool]:
        """Every message of the conversation in scope (all copies, not merged).

        Assumes (not re-checked here): ``conversation_id`` is taken as given (from this connector's own
        results) and ``scope`` was validated by ``set_read``.
        """
        items, truncated = await self.mailbox.reader.conversation(conversation_id)
        skip = await self.mailbox.exclusions(scope)
        folders, hidden = await self.mailbox.reach(m.folder_id for m in items)
        return [
            m
            for m in items
            if m.folder_id in folders and m.folder_id not in hidden and m.folder_id not in skip
        ], truncated

    async def _apply(
        self,
        action: str,
        message_ids: list[str],
        already: Wanted,
        send: Send,
        *,
        prepare: Callable[[], Awaitable[None]] | None = None,
        continue_on_error: bool = True,
        known: dict[str, MessageSummary] | None = None,
    ) -> MutationResult:
        """Read state, send the change once per chunk, report a result per message.

        Assumes (not re-checked here): ids are unique, nonempty and validated by the public method;
        only explicit inputs are bounded to MAX_MUTATION_ITEMS. Expanded conversations may be larger.
        The writer trusts the 20-message chunks it is given.
        """
        ids = message_ids
        self.check_account()
        known = known or {}
        unread = [mid for mid in ids if mid not in known]
        before = await self.mailbox.reader.get_summaries(unread) if unread else FetchedSummaries()
        summaries = known | before.summaries
        folders, hidden = await self.mailbox.reach(m.folder_id for m in summaries.values() if m)
        if prepare:
            await prepare()
        results: dict[str, ItemResult] = {}
        pending: list[str] = []
        for mid in ids:
            summary = summaries.get(mid)
            if mid in before.failed:
                results[mid] = ItemResult(id=mid, status="failed", detail=describe(before.failed[mid]))
            elif summary is None:
                results[mid] = ItemResult(id=mid, status="not_found", detail="Not on the server.")
            elif summary.folder_id not in folders or summary.folder_id in hidden:
                results[mid] = ItemResult(id=mid, status="failed", detail="Out of reach (hidden folder).")
            elif already(summary):
                results[mid] = ItemResult(id=mid, status="unchanged")
            else:
                pending.append(mid)
        stopped = False
        for start in range(0, len(pending), CHUNK):
            chunk = pending[start : start + CHUNK]
            if stopped:
                results.update({mid: ItemResult(id=mid, status="failed", detail="not sent") for mid in chunk})
                continue
            try:
                outcomes = await send(chunk)
            except WriteOutcomeUnknown as exc:
                results |= await self._recheck(chunk, already, str(exc))
            except ConnectorError as exc:
                results.update({mid: ItemResult(id=mid, status="failed", detail=str(exc)) for mid in chunk})
            else:
                for mid in chunk:
                    code = outcomes.get(mid)
                    if code is None:
                        results[mid] = ItemResult(id=mid, status="done")
                    elif code == "ErrorItemNotFound":
                        results[mid] = ItemResult(id=mid, status="not_found", detail=code)
                    else:
                        results[mid] = ItemResult(id=mid, status="failed", detail=code)
            if not continue_on_error and any(results[mid].status != "done" for mid in chunk):
                stopped = True
        ordered = [results[mid] for mid in ids]
        return MutationResult(action=action, results=ordered, counts=dict(Counter(r.status for r in ordered)))

    async def _recheck(self, chunk: list[str], already: Wanted, reason: str) -> dict[str, ItemResult]:
        """After an unclear answer: done where the change is visible, unknown elsewhere.

        Assumes (not re-checked here): ``chunk`` is one ``_apply`` already sent, with no clear answer.
        """
        try:
            after = await self.mailbox.reader.get_summaries(chunk)
        except ConnectorError as exc:
            return {
                mid: ItemResult(id=mid, status="unknown", detail=f"{reason} Read-back failed: {exc}")
                for mid in chunk
            }
        out = {}
        for mid in chunk:
            summary = after.summaries.get(mid)
            if summary is not None and already(summary):
                out[mid] = ItemResult(id=mid, status="done", detail="Confirmed by reading it back.")
            else:
                out[mid] = ItemResult(id=mid, status="unknown", detail=reason)
        return out


def _message_ids(message_ids: list[str], *, allow_empty: bool = False) -> list[str]:
    """Validate explicit message selection once at each mutation's public entry point."""
    ids = list(dict.fromkeys(message_ids))
    if not ids and not allow_empty:
        raise InvalidRequest("Name at least one message id.")
    if len(ids) > MAX_MUTATION_ITEMS:
        raise InvalidRequest(f"At most {MAX_MUTATION_ITEMS} messages per call.")
    return ids

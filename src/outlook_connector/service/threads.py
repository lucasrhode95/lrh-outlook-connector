"""Conversation retrieval (requirements v4 §8, research §3.4).

A thread is an Exchange conversation: every message with one conversationId across all folders.
Graph cannot sort that query, so messages are sorted here. Retained messages deleted on the
server are merged in and labelled. Bodies are optional and bounded, with a cursor.
"""

from __future__ import annotations

import re
from typing import Literal

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.domain.models import (
    Coverage,
    Message,
    MessageSummary,
    Thread,
    ThreadMessage,
    ThreadSize,
)
from outlook_connector.remote.ports import FetchedMessages
from outlook_connector.service import cursors
from outlook_connector.service.mailbox import Mailbox

EXCLUDED_BY_DEFAULT = ("deleteditems", "junkemail")
BODY_BATCH = 10
MAX_SIZE_LOOKUPS = 200
_PREFIX = re.compile(r"^\s*((re|res|fw|fwd|enc|aw|wg|sv|tr|rv)\s*:\s*)+", re.IGNORECASE)


def base_subject(subject: str | None) -> str | None:
    return _PREFIX.sub("", subject).strip() if subject else subject


def oldest_first(m: MessageSummary) -> float:
    stamp = m.received_at or m.sent_at
    return stamp.timestamp() if stamp else 0.0


def _has_content(message: Message) -> bool:
    return any(
        value is not None
        for value in (
            message.body_text,
            message.unique_body_text,
            message.body_html,
            message.unique_body_html,
        )
    )


class Threads:
    def __init__(self, mailbox: Mailbox) -> None:
        self.mailbox = mailbox

    async def messages(
        self, conversation_id: str, *, include_deleted_items: bool = False
    ) -> tuple[list[MessageSummary], int, bool]:
        """All messages of the conversation, oldest first, how many were excluded by folder, and
        whether the server listing was truncated (more than MAX_CONVERSATION messages)."""
        reader, store = self.mailbox.reader, self.mailbox.store
        remote, truncated = await reader.conversation(conversation_id)
        store.upsert_summaries(remote)
        await self.mailbox.reconciler.after_conversation(conversation_id, remote)
        retained_deleted = [m for m in store.conversation(conversation_id) if m.is_deleted]
        if not remote and not retained_deleted:
            raise NotFound(f"No conversation {conversation_id} on the server or locally.")
        excluded_ids = await self._excluded_folder_ids()
        kept = [m for m in remote if include_deleted_items or m.folder_id not in excluded_ids]
        everything = sorted(kept + retained_deleted, key=oldest_first)
        return await self.mailbox.decorate(everything), len(remote) - len(kept), truncated

    async def _excluded_folder_ids(self) -> set[str]:
        folders = await self.mailbox.folder_map()
        return {f.id for f in folders.values() if f.well_known in EXCLUDED_BY_DEFAULT}

    async def sizes(
        self, conversation_ids: list[str], *, include_deleted_items: bool = False
    ) -> list[ThreadSize]:
        """How many messages each conversation has, counted the way get_thread lists them.

        One batched server listing of ids and folders per conversation, plus retained messages
        deleted on the server. Lets a list view tell single messages from real threads.
        """
        ids = list(dict.fromkeys(conversation_ids))
        if len(ids) > MAX_SIZE_LOOKUPS:
            raise InvalidRequest(f"At most {MAX_SIZE_LOOKUPS} conversations per request.")
        if not ids:
            return []
        remote = await self.mailbox.reader.conversation_folders(ids)
        excluded_ids = await self._excluded_folder_ids()
        out = []
        for cid in ids:
            folders, more = remote.get(cid, ([], False))
            kept = sum(1 for f in folders if include_deleted_items or f not in excluded_ids)
            retained = sum(1 for m in self.mailbox.store.conversation(cid) if m.is_deleted)
            out.append(ThreadSize(conversation_id=cid, messages=kept + retained, at_least=more))
        return out

    async def get_thread(
        self,
        conversation_id: str,
        *,
        include_bodies: bool = True,
        body: Literal["unique", "full"] = "unique",
        include_deleted_items: bool = False,
        max_chars: int = 40000,
        cursor: str | None = None,
    ) -> Thread:
        start = 0
        if cursor:  # the cursor carries the original selection, so a continuation never drifts
            state = cursors.decode(cursor, "get_thread")
            if state.get("conversation_id") != conversation_id:
                raise InvalidRequest("This cursor belongs to a different conversation.")
            start = int(state["start"])
            include_bodies = bool(state["include_bodies"])
            body = state["body"]
            include_deleted_items = bool(state["include_deleted_items"])
            max_chars = int(state["max_chars"])
        if not 1 <= max_chars <= 400_000:
            raise InvalidRequest("max_chars must be between 1 and 400000.")
        items, excluded, truncated = await self.messages(
            conversation_id, include_deleted_items=include_deleted_items
        )
        notes = ["Messages are sorted oldest first (sorted locally)."]
        if truncated:
            notes.append(
                "The conversation has more messages than the server listing limit; only the first ones "
                "the server returned are included. Use search_messages or list_messages for the rest."
            )
        if excluded:
            notes.append(
                f"{excluded} message(s) in Deleted Items or Junk Email omitted; "
                "pass include_deleted_items=true to include them."
            )
        if any(m.is_deleted for m in items):
            notes.append(
                "Messages marked is_deleted=true were deleted on the server; shown from local retention."
            )

        entries: list[ThreadMessage] = []
        next_start: int | None = None
        unavailable = 0
        if include_bodies:
            budget = max_chars
            index = start
            while index < len(items) and next_start is None:
                chunk = items[index : index + BODY_BATCH]
                bodies, failed = await self.bodies(chunk)
                unavailable += len(failed)
                for offset, summary in enumerate(chunk):
                    if summary.id in failed:
                        text = f"(Body could not be fetched: {failed[summary.id]})"
                    else:
                        text = bodies[summary.id].body(body) if bodies.get(summary.id) else ""
                    if len(text) > budget and entries:
                        next_start = index + offset
                        break
                    truncated = len(text) > budget
                    entries.append(ThreadMessage(message=summary, text=text[:budget], truncated=truncated))
                    budget -= min(len(text), budget)
                    if budget <= 0 and index + offset + 1 < len(items):
                        next_start = index + offset + 1
                        break
                index += len(chunk)
        else:
            entries = [ThreadMessage(message=m) for m in items[start:]]
        if unavailable:
            notes.append(f"{unavailable} message body(ies) could not be fetched and are marked in the text.")

        return Thread(
            conversation_id=conversation_id,
            subject=base_subject(items[0].subject) if items else None,
            messages=entries,
            cursor=cursors.encode(
                "get_thread",
                start=next_start,
                conversation_id=conversation_id,
                include_bodies=include_bodies,
                body=body,
                include_deleted_items=include_deleted_items,
                max_chars=max_chars,
            )
            if next_start is not None
            else None,
            coverage=Coverage(
                source="remote+local" if any(m.is_deleted for m in items) else "remote",
                complete=next_start is None and not truncated and not unavailable,
                more_available=next_start is not None,
                notes=notes,
            ),
        )

    async def bodies(
        self, summaries: list[MessageSummary]
    ) -> tuple[dict[str, Message | None], dict[str, str]]:
        """Text bodies for a set of messages: from the server, or the retained copy for deleted ones.

        Messages the server no longer has are marked deleted, in the store and on ``summaries``.
        The second map holds the reason for each body that could not be fetched (e.g. still
        throttled after retries) and has no retained copy either.
        """
        store = self.mailbox.store
        live = [m.id for m in summaries if not m.is_deleted]
        fetched = await self.mailbox.reader.get_messages(live) if live else FetchedMessages()
        found = [m for m in fetched.messages.values() if m is not None]
        store.save_messages(found)
        gone = [mid for mid, m in fetched.messages.items() if m is None]
        if gone:
            store.mark_deleted(gone)
            for summary in summaries:
                if summary.id in gone:
                    summary.is_deleted = True
        out: dict[str, Message | None] = dict(fetched.messages)
        failed: dict[str, str] = {}
        for summary in summaries:
            if out.get(summary.id) is None:
                retained = store.message(summary.id)
                # a row without retained content is only a summary, not a body
                out[summary.id] = retained if retained and _has_content(retained) else None
                if out[summary.id] is None and summary.id in fetched.failed:
                    failed[summary.id] = fetched.failed[summary.id]
        return out, failed

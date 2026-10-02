"""Conversation retrieval (requirements v4 §8, research §3.4).

A thread is an Exchange conversation: every message with one conversationId across all folders.
Graph cannot sort that query, so messages are sorted here. Retained messages deleted on the
server are merged in and labelled, and copies of one message are shown once (mailbox.py).
Bodies are optional and bounded, with a cursor.
"""

from __future__ import annotations

import re
from typing import Literal

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.domain.models import Coverage, Message, MessageSummary, Thread, ThreadMessage
from outlook_connector.remote.ports import FetchedMessages
from outlook_connector.service import cursors
from outlook_connector.service.mailbox import NEVER_RETAINED, Mailbox, has_content, unavailable

BODY_BATCH = 10
_PREFIX = re.compile(r"^\s*((re|res|fw|fwd|enc|aw|wg|sv|tr|rv)\s*:\s*)+", re.IGNORECASE)


def base_subject(subject: str | None) -> str | None:
    return _PREFIX.sub("", subject).strip() if subject else subject


def oldest_first(m: MessageSummary) -> float:
    stamp = m.received_at or m.sent_at
    return stamp.timestamp() if stamp else 0.0


class Threads:
    def __init__(self, mailbox: Mailbox) -> None:
        self.mailbox = mailbox

    async def messages(
        self, conversation_id: str, *, include_deleted_items: bool = False
    ) -> tuple[list[MessageSummary], dict[str, int], bool]:
        """All messages of the conversation, oldest first, what was left out by folder, and whether
        the server listing was truncated (more than MAX_CONVERSATION messages)."""
        reader, store = self.mailbox.reader, self.mailbox.store
        remote, truncated = await reader.conversation(conversation_id)
        store.upsert_summaries(remote)
        await self.mailbox.reconciler.after_conversation(conversation_id, remote)
        retained_deleted = [m for m in store.conversation(conversation_id) if m.is_deleted]
        if not remote and not retained_deleted:
            raise NotFound(f"No conversation {conversation_id} on the server or locally.")
        skip = await self.mailbox.exclusions(include_deleted_items=include_deleted_items)
        items, excluded = await self.mailbox.finish(sorted(remote + retained_deleted, key=oldest_first), skip)
        return items, excluded, truncated

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
        items, excluded, listing_truncated = await self.messages(
            conversation_id, include_deleted_items=include_deleted_items
        )
        notes = ["Messages are sorted oldest first (sorted locally)."]
        if listing_truncated:
            notes.append(
                "The conversation has more messages than the server listing limit; only the first ones "
                "the server returned are included. Use search_messages or list_messages for the rest."
            )
        if excluded:
            notes.append(
                f"{sum(excluded.values())} message(s) in Deleted Items or Junk Email omitted; "
                "pass include_deleted_items=true to include them."
            )
        if any(m.is_deleted for m in items):
            notes.append(
                "Messages marked is_deleted=true were deleted on the server; shown from local retention."
            )
        if any(m.also_in for m in items):
            notes.append("Copies of one message are shown once; also_in names the other folders.")

        entries: list[ThreadMessage] = []
        next_start: int | None = None
        retryable = 0  # bodies the server could not deliver now (unlike ones never retained)
        if include_bodies:
            budget = max_chars
            index = start
            while index < len(items) and next_start is None:
                chunk = items[index : index + BODY_BATCH]
                bodies, missing = await self.bodies(chunk)
                retryable += sum(1 for reason in missing.values() if reason != NEVER_RETAINED)
                for offset, summary in enumerate(chunk):
                    found = bodies.get(summary.id)
                    text = found.body(body) if found else unavailable(missing[summary.id])
                    if len(text) > budget and entries:
                        next_start = index + offset
                        break
                    cut = len(text) > budget
                    entries.append(ThreadMessage(message=summary, text=text[:budget], truncated=cut))
                    budget -= min(len(text), budget)
                    if budget <= 0 and index + offset + 1 < len(items):
                        next_start = index + offset + 1
                        break
                index += len(chunk)
        else:
            entries = [ThreadMessage(message=m) for m in items[start:]]
        if retryable:
            notes.append(
                f"{retryable} message body(ies) could not be fetched now and are marked in the text."
            )

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
                complete=next_start is None and not listing_truncated and not retryable,
                excluded=excluded,
                notes=notes,
            ),
        )

    async def bodies(
        self, summaries: list[MessageSummary], *, known: dict[str, Message] | None = None
    ) -> tuple[dict[str, Message], dict[str, str]]:
        """Text bodies: from the server, or the retained copy for messages deleted there.

        ``known``: messages already fetched with bodies, used as they are. Messages the server no
        longer has are marked deleted, in the store and on ``summaries``. Returns (bodies by id,
        reason per message without a body): every summary is in exactly one of the two.
        """
        store = self.mailbox.store
        known = known or {}
        live = [m.id for m in summaries if not m.is_deleted and m.id not in known]
        fetched = await self.mailbox.reader.get_messages(live) if live else FetchedMessages()
        found = [m for m in fetched.messages.values() if m is not None]
        store.save_messages(found)
        gone = {mid for mid, m in fetched.messages.items() if m is None}
        if gone:
            store.mark_deleted(gone)
            for summary in summaries:
                if summary.id in gone:
                    summary.is_deleted = True
        out: dict[str, Message] = {m.id: m for m in found} | {
            mid: m for mid, m in known.items() if mid in {s.id for s in summaries}
        }
        retained = store.messages(s.id for s in summaries if s.id not in out)
        missing: dict[str, str] = {}
        for summary in summaries:
            if summary.id in out:
                continue
            copy = retained.get(summary.id)
            if copy is not None and has_content(copy):  # a row without content is only a summary
                out[summary.id] = copy
            else:
                missing[summary.id] = fetched.failed.get(summary.id, NEVER_RETAINED)
        return out, missing

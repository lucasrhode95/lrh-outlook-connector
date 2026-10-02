"""Conversation retrieval (requirements v4 §8, research §3.4).

A thread is an Exchange conversation: every message with one conversationId across all folders.
Graph cannot sort that query, so messages are sorted here. Copies of one message are shown once
(mailbox.py).
Bodies are optional and bounded, with a cursor.
"""

from __future__ import annotations

import re
from typing import Literal

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.domain.models import (
    EXCLUSION_TEXT,
    Coverage,
    Message,
    MessageSummary,
    Thread,
    ThreadMessage,
)
from outlook_connector.remote.ports import FetchedMessages
from outlook_connector.service import cursors
from outlook_connector.service.mailbox import GONE, Mailbox, unavailable

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
        remote, truncated = await self.mailbox.reader.conversation(conversation_id)
        if not remote:
            raise NotFound(f"No conversation {conversation_id} on the server.")
        self.mailbox.store.upsert_summaries(remote)
        skip = await self.mailbox.exclusions(include_deleted_items=include_deleted_items)
        items, excluded = await self.mailbox.finish(sorted(remote, key=oldest_first), skip)
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
        for reason, count in excluded.items():
            notes.append(f"{count} message(s) left out: {EXCLUSION_TEXT[reason]}.")
        if any(m.also_in for m in items):
            notes.append("Copies of one message are shown once; also_in names the other folders.")

        entries: list[ThreadMessage] = []
        next_start: int | None = None
        retryable = 0  # bodies the server could not deliver now (not deleted ones)
        if include_bodies:
            budget = max_chars
            index = start
            while index < len(items) and next_start is None:
                chunk = items[index : index + BODY_BATCH]
                bodies, missing = await self.bodies(chunk)
                retryable += sum(1 for reason in missing.values() if reason != GONE)
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
                complete=next_start is None and not listing_truncated and not retryable,
                excluded=excluded,
                notes=notes,
            ),
        )

    async def bodies(
        self, summaries: list[MessageSummary], *, known: dict[str, Message] | None = None
    ) -> tuple[dict[str, Message], dict[str, str]]:
        """Text bodies from the server, in batches.

        ``known``: messages already fetched with bodies, used as they are. Returns (bodies by id,
        reason per message without a body): every summary is in exactly one of the two.
        """
        known = {s.id: known[s.id] for s in summaries if known and s.id in known}
        wanted = [s.id for s in summaries if s.id not in known]
        fetched = await self.mailbox.reader.get_messages(wanted) if wanted else FetchedMessages()
        out = known | {mid: m for mid, m in fetched.messages.items() if m is not None}
        missing = {s.id: fetched.failed.get(s.id, GONE) for s in summaries if s.id not in out}
        return out, missing

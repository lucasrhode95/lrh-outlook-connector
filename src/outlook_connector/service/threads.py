"""Conversation retrieval (requirements v4 §8, research §3.4).

A thread is an Exchange conversation: every message with one conversationId across all folders.
Graph cannot sort that query, so messages are sorted here. Retained messages deleted on the
server are merged in and labelled. Bodies are optional and bounded, with a cursor.
"""

from __future__ import annotations

import re
from typing import Literal

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.domain.models import Coverage, Message, MessageSummary, Thread, ThreadMessage
from outlook_connector.service import cursors
from outlook_connector.service.mailbox import Mailbox

EXCLUDED_BY_DEFAULT = ("deleteditems", "junkemail")
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
    ) -> tuple[list[MessageSummary], int]:
        """All messages of the conversation, oldest first, and how many were excluded by folder."""
        reader, store = self.mailbox.reader, self.mailbox.store
        remote = await reader.conversation(conversation_id)
        store.upsert_summaries(remote)
        await self.mailbox.reconciler.after_conversation(conversation_id, remote)
        retained_deleted = [m for m in store.conversation(conversation_id) if m.is_deleted]
        if not remote and not retained_deleted:
            raise NotFound(f"No conversation {conversation_id} on the server or locally.")
        folders = await self.mailbox.folder_map()
        excluded_ids = {f.id for f in folders.values() if f.well_known in EXCLUDED_BY_DEFAULT}
        kept = [m for m in remote if include_deleted_items or m.folder_id not in excluded_ids]
        everything = sorted(kept + retained_deleted, key=oldest_first)
        return await self.mailbox.decorate(everything), len(remote) - len(kept)

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
        if not 1 <= max_chars <= 400_000:
            raise InvalidRequest("max_chars must be between 1 and 400000.")
        start = int(cursors.decode(cursor, "get_thread")["start"]) if cursor else 0
        items, excluded = await self.messages(conversation_id, include_deleted_items=include_deleted_items)
        notes = ["Messages are sorted oldest first (sorted locally)."]
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
        if include_bodies:
            budget = max_chars
            index = start
            while index < len(items) and next_start is None:
                chunk = items[index : index + BODY_BATCH]
                bodies = await self.bodies(chunk)
                for offset, summary in enumerate(chunk):
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

        return Thread(
            conversation_id=conversation_id,
            subject=base_subject(items[0].subject) if items else None,
            messages=entries,
            cursor=cursors.encode("get_thread", start=next_start) if next_start is not None else None,
            coverage=Coverage(
                source="remote+local" if any(m.is_deleted for m in items) else "remote",
                complete=next_start is None,
                more_available=next_start is not None,
                notes=notes,
            ),
        )

    async def bodies(self, summaries: list[MessageSummary]) -> dict[str, Message | None]:
        """Text bodies for a set of messages: from the server, or the retained copy for deleted ones."""
        store = self.mailbox.store
        live = [m.id for m in summaries if not m.is_deleted]
        fetched = await self.mailbox.reader.get_messages(live) if live else {}
        found = [m for m in fetched.values() if m is not None]
        store.save_messages(found)
        gone = [mid for mid, m in fetched.items() if m is None]
        if gone:
            store.mark_deleted(gone)
        out: dict[str, Message | None] = dict(fetched)
        for summary in summaries:
            if out.get(summary.id) is None:
                out[summary.id] = store.message(summary.id)
        return out

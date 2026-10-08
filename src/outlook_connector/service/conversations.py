"""Conversation retrieval (requirements v4 §8, research §3.4).

A conversation is an Exchange conversation: every message with one conversationId across all folders.
Graph cannot sort that query, so messages are sorted here. Copies of one message are shown once
(mailbox.py).
Bodies are optional and bounded, with a cursor.
"""

from __future__ import annotations

import re
from typing import Literal

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.domain.models import (
    DEFAULT_SCOPE,
    EXCLUSION_TEXT,
    Conversation,
    ConversationMessage,
    Coverage,
    ExportError,
    ExportStep,
    Message,
    MessageSummary,
    Scope,
)
from outlook_connector.remote.ports import FetchedMessages
from outlook_connector.service import cursors
from outlook_connector.service.failures import error_block, export_error, gone
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.scope import validate_scope

BODY_BATCH = 10
BODY_STEP: ExportStep = "fetching message bodies"
BODY_MISSING = "The body of this message could not be fetched."
_PREFIX = re.compile(r"^\s*((re|res|fw|fwd|enc|aw|wg|sv|tr|rv)\s*:\s*)+", re.IGNORECASE)


def base_subject(subject: str | None) -> str | None:
    return _PREFIX.sub("", subject).strip() if subject else subject


def oldest_first(m: MessageSummary) -> float:
    stamp = m.received_at or m.sent_at
    return stamp.timestamp() if stamp else 0.0


class Conversations:
    def __init__(self, mailbox: Mailbox) -> None:
        self.mailbox = mailbox

    async def messages(
        self, conversation_id: str, *, scope: Scope
    ) -> tuple[list[MessageSummary], dict[str, int], bool]:
        """All messages of the conversation, oldest first, what was left out by folder, and whether
        the server listing was truncated (more than MAX_CONVERSATION messages).

        Assumes (not re-checked here): ``conversation_id`` is taken as given (from this connector's own
        results) and ``scope`` was validated at the public entry point; an unknown id raises NotFound.
        """
        remote, truncated = await self.mailbox.reader.conversation(conversation_id)
        if not remote:
            raise NotFound(f"No conversation {conversation_id} on the server.")
        skip = await self.mailbox.exclusions(scope)
        items, excluded = await self.mailbox.finish(sorted(remote, key=oldest_first), skip)
        return items, excluded, truncated

    async def get_conversation(
        self,
        conversation_id: str,
        *,
        include_bodies: bool = True,
        body: Literal["unique", "full"] = "unique",
        scope: Scope = DEFAULT_SCOPE,
        max_chars: int = 40000,
        cursor: str | None = None,
    ) -> Conversation:
        """A whole conversation across folders, oldest first, with bounded bodies and a cursor.

        Entry point: the authoritative check of ``max_chars`` and the cursor (the web routes pass them
        through unchecked).
        """
        start = 0
        if cursor:  # the cursor carries the original selection, so a continuation never drifts
            state = cursors.decode(cursor, "get_conversation")
            if state.get("conversation_id") != conversation_id:
                raise InvalidRequest("This cursor belongs to a different conversation.")
            start = int(state["start"])
            include_bodies = bool(state["include_bodies"])
            body = state["body"]
            scope = Scope.model_validate(state["scope"])
            max_chars = int(state["max_chars"])
        validate_scope(scope, sent_items=False, meeting_mail=False)
        if not 1 <= max_chars <= 400_000:
            raise InvalidRequest("max_chars must be between 1 and 400000.")
        items, excluded, listing_truncated = await self.messages(conversation_id, scope=scope)
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

        entries: list[ConversationMessage] = []
        next_start: int | None = None
        if include_bodies:
            budget = max_chars
            index = start
            while index < len(items) and next_start is None:
                chunk = items[index : index + BODY_BATCH]
                bodies, missing = await self.bodies(chunk)
                for offset, summary in enumerate(chunk):
                    error = missing.get(summary.id)
                    text = error_block(BODY_MISSING, error) if error else bodies[summary.id].body(body)
                    if len(text) > budget and entries:
                        next_start = index + offset
                        break
                    cut = len(text) > budget
                    entries.append(
                        ConversationMessage(
                            message=summary, text=text[:budget], truncated=cut, export_error=error
                        )
                    )
                    budget -= min(len(text), budget)
                    if budget <= 0 and index + offset + 1 < len(items):
                        next_start = index + offset + 1
                        break
                index += len(chunk)
        else:
            entries = [ConversationMessage(message=m) for m in items[start:]]
        errors = [e.export_error for e in entries if e.export_error]
        retryable = sum(1 for error in errors if error.retry)
        if errors:
            notes.append(
                f"{len(errors)} message body(ies) could not be fetched ({retryable} may work on a retry); "
                "those messages carry export_error and an [EXPORT ERROR] block in their text."
            )

        return Conversation(
            conversation_id=conversation_id,
            subject=base_subject(items[0].subject) if items else None,
            messages=entries,
            cursor=cursors.encode(
                "get_conversation",
                start=next_start,
                conversation_id=conversation_id,
                include_bodies=include_bodies,
                body=body,
                scope=scope.model_dump(mode="json"),
                max_chars=max_chars,
            )
            if next_start is not None
            else None,
            coverage=Coverage(
                complete=next_start is None and not listing_truncated and not retryable,
                excluded=excluded,
                notes=notes,
            ),
            body_errors=len(errors),
        )

    async def bodies(
        self, summaries: list[MessageSummary], *, known: dict[str, Message] | None = None
    ) -> tuple[dict[str, Message], dict[str, ExportError]]:
        """Text bodies from the server, in batches.

        ``known``: messages already fetched with bodies, used as they are. Returns (bodies by id,
        the error for each message without a body): every summary is in exactly one of the two.

        Assumes (not re-checked here): ``summaries`` is the caller's final selection (in scope, copies
        merged), and ``known`` messages were fetched as text bodies.
        """
        known = {s.id: known[s.id] for s in summaries if known and s.id in known}
        wanted = [s.id for s in summaries if s.id not in known]
        fetched = await self.mailbox.reader.get_messages(wanted) if wanted else FetchedMessages()
        out = known | {mid: m for mid, m in fetched.messages.items() if m is not None}
        missing = {
            s.id: export_error(BODY_STEP, fetched.failed[s.id]) if s.id in fetched.failed else gone(BODY_STEP)
            for s in summaries
            if s.id not in out
        }
        return out, missing

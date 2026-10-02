"""Rendering of exported messages (requirements v4 §10.1): TXT for people, JSONL for agents."""

from __future__ import annotations

import json
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any

from outlook_connector.domain.models import Message, MessageSummary, Recipient

RULE = "=" * 78
THIN = "-" * 78


@dataclass
class RenderedMessage:
    message: MessageSummary
    text: str
    attachment_lines: list[str] = field(default_factory=list)
    unavailable: str | None = None  # why the body is missing (``text`` then holds the marker)
    attachments: list[dict[str, Any]] = field(default_factory=list)  # JSONL attachment records


def stamp(value: datetime | None) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%d %H:%M UTC") if value else "(no date)"


def people(values: Sequence[Recipient]) -> str:
    # "; " as in Outlook: display names are often "Last, First", so a comma would be ambiguous
    return "; ".join(r.display() for r in values)


def body_text(message: Message, kind: str) -> str:
    return message.body(kind).strip() or "(No text content.)"  # type: ignore[arg-type]


def render_message(item: RenderedMessage, *, position: str) -> str:
    m = item.message
    lines = [RULE, f"[{position}] {stamp(m.received_at or m.sent_at)}"]
    lines.append(f"From:    {m.sender.display() if m.sender else '(unknown)'}")
    if m.to:
        lines.append(f"To:      {people(m.to)}")
    if m.cc:
        lines.append(f"Cc:      {people(m.cc)}")
    lines.append(f"Subject: {m.subject or '(no subject)'}")
    if m.folder:
        lines.append(f"Folder:  {m.folder}")
    lines.append(f"Message id:      {m.id}")
    if m.conversation_id:
        lines.append(f"Conversation id: {m.conversation_id}")
    if m.internet_message_id:
        lines.append(f"Internet id:     {m.internet_message_id}")
    if m.also_in:
        lines.append(f"Also in: {'; '.join(m.also_in)} (same message, exported once)")
    if m.is_deleted:
        deleted = f" on {stamp(m.deleted_at)}" if m.deleted_at else ""
        lines.append(f"!! DELETED on the server{deleted}. This is the copy retained by outlook-connector.")
    for line in item.attachment_lines:
        lines.append(f"Attachment: {line}")
    lines.append(THIN)
    lines.append(item.text)
    return "\n".join(lines)


def render_file(
    title: str,
    items: list[RenderedMessage],
    *,
    body_kind: str,
    sections: bool,
    notes: Sequence[str] = (),
) -> str:
    """One TXT file. With ``sections``, consecutive messages of different conversations get a header.

    ``notes`` (what the export left out or could not fetch) go in the header. Every message carries
    its ids, which get_message/get_thread accept, so an export can be traced back to its source.
    """
    dates = [m.message.received_at or m.message.sent_at for m in items]
    known = [d for d in dates if d]
    span = f"{stamp(min(known))} to {stamp(max(known))}" if known else "no dates"
    head = [
        title,
        f"Messages: {len(items)} ({span})",
        f"Exported: {stamp(datetime.now(UTC))} by outlook-connector; body: "
        + ("without quoted history" if body_kind == "unique" else "full, including quoted history"),
        *notes,
        "",
    ]
    parts = ["\n".join(head)]
    current: str | None = None
    for index, item in enumerate(items, start=1):
        conversation = item.message.conversation_id or item.message.id
        if sections and conversation != current:
            parts.append(f"\n### Conversation: {item.message.subject or '(no subject)'}\n")
            current = conversation
        parts.append(render_message(item, position=f"{index}/{len(items)}"))
    return "\n\n".join(parts).rstrip() + "\n"


def jsonl_record(item: RenderedMessage, *, body_kind: str) -> str:
    """One JSON line per message: source ids, dates, people and the body, for agents and trackers."""
    m = item.message

    def person(r: Recipient | None) -> dict[str, str | None] | None:
        return {"name": r.name, "address": r.address} if r else None

    record = {
        "id": m.id,
        "conversation_id": m.conversation_id,
        "internet_message_id": m.internet_message_id,
        "folder": m.folder,
        "also_in": m.also_in,
        "received_at": m.received_at.isoformat() if m.received_at else None,
        "sent_at": m.sent_at.isoformat() if m.sent_at else None,
        "subject": m.subject,
        "from": person(m.sender),
        "to": [person(r) for r in m.to],
        "cc": [person(r) for r in m.cc],
        "is_deleted": m.is_deleted,
        "deleted_at": m.deleted_at.isoformat() if m.deleted_at else None,
        "body_kind": body_kind,
        "body": None if item.unavailable else item.text,
        "body_unavailable": item.unavailable,
        "attachments": item.attachments,
    }
    return json.dumps(record, ensure_ascii=False)

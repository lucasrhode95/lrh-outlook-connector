"""Attachment markers in message text (requirements §10): name files from attachment metadata.

Image references retain their source position as typed body parts. Attachment identity comes from the
source listing. Plain text and rendered markers are never searched to recover either fact.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence

from outlook_connector.domain.body import BodyPart, ImagePart, TextPart
from outlook_connector.domain.models import Attachment
from outlook_connector.service.export.attachments import dedupe, safe_name, wanted

NAMED_KINDS = ("file", "item", "reference")
DOWNLOADABLE_KINDS = ("file", "item")


def file_names(attachments: Sequence[Attachment]) -> dict[str, str]:
    """Attachment id -> the name its marker shows, unique within the message, in listing order."""
    taken: set[str] = set()
    return {
        a.id: dedupe(safe_name(a.name, fallback="attachment", eml=a.kind == "item"), taken)
        for a in attachments
        if a.kind in NAMED_KINDS
    }


def named(attachments: Sequence[Attachment]) -> list[Attachment]:
    """The attachments with ``file_name`` set to their marker name."""
    names = file_names(attachments)
    return [a.model_copy(update={"file_name": names.get(a.id)}) for a in attachments]


def with_markers(
    text: str,
    attachments: Sequence[Attachment],
    *,
    parts: Sequence[BodyPart] | None = None,
    names: Mapping[str, str] | None = None,
) -> str:
    """Render source body parts with inline markers, then list remaining eligible attachments.

    ``parts`` carries the source body order and image references. When unavailable, ``text`` stays the
    body and inline files are explicitly listed without a known body position. ``names`` overrides
    marker names (an export uses the names of the files it wrote).

    Assumes (not re-checked here): ``text``, ``parts`` and ``attachments`` belong to the same message and
    body kind. Metadata decisions use typed parts and attachment fields only.
    """
    names = dict(names) if names is not None else file_names(attachments)
    out, placed = _body(parts, attachments, names) if parts is not None else (text, set())
    trailing = [
        a
        for a in attachments
        if a.id in names and a.id not in placed and (a.kind != "file" or wanted(a, parts=parts))
    ]
    if not trailing:
        return out
    blocks = [out.rstrip("\r\n")] if out.strip() else []
    files = [_marker(a, names[a.id]) for a in trailing if not a.is_inline]
    inline = [_marker(a, names[a.id]) for a in trailing if a.is_inline]
    if files:
        blocks.append("\n".join(files))
    if inline:
        blocks.append("Inline attachments without a body position:\n" + "\n".join(inline))
    return "\n\n".join(blocks)


def _body(
    parts: Sequence[BodyPart], attachments: Sequence[Attachment], names: Mapping[str, str]
) -> tuple[str, set[str]]:
    by_cid: dict[str, list[Attachment]] = {}
    for attachment in attachments:
        if attachment.kind == "file" and attachment.content_id and attachment.id in names:
            by_cid.setdefault(attachment.content_id, []).append(attachment)
    output: list[str] = []
    placed: set[str] = set()
    for index, part in enumerate(parts):
        if isinstance(part, TextPart):
            output.append(part.text)
            continue
        matches = by_cid.get(part.content_id, []) if part.content_id else []
        if len(matches) == 1:
            attachment = matches[0]
            marker = f"[Image: {names[attachment.id]}]"
            placed.add(attachment.id)
        else:
            label = "Inline image unavailable" if part.content_id else "External image"
            marker = f"[{label}: {part.alt}]" if part.alt else f"[{label}]"
        if output and output[-1] and not output[-1][-1].isspace():
            output.append(" ")
        output.append(marker)
        following = parts[index + 1] if index + 1 < len(parts) else None
        if isinstance(following, ImagePart) or (
            isinstance(following, TextPart) and following.text and not following.text[0].isspace()
        ):
            output.append(" ")
    return "".join(output).strip("\r\n"), placed


def _marker(attachment: Attachment, name: str) -> str:
    if attachment.kind == "item":
        return f"[Message: {name}]"
    if attachment.kind == "reference":
        return f"[Link: {name}]"
    if attachment.is_inline:
        return f"[Image: {name}]"
    return f"[File: {name}]"

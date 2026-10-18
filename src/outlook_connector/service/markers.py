"""Attachment markers in message text (requirements §10): name files from attachment metadata.

Every known attachment is named after the message text using its source kind and attachment name. The
body text and HTML do not decide which attachments exist or where their markers appear.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence

from outlook_connector.domain.models import Attachment
from outlook_connector.service.export.attachments import dedupe, safe_name

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
    names: Mapping[str, str] | None = None,
) -> str:
    """``text`` followed by a marker for every known attachment.

    ``names`` overrides marker names (an export uses the names of the files it wrote).

    Assumes (not re-checked here): ``attachments`` is the listing for the message that supplied ``text``.
    """
    names = dict(names) if names is not None else file_names(attachments)
    trailing = [a for a in attachments if a.id in names]
    if not trailing:
        return text
    lines = "\n".join(_marker(a, names[a.id]) for a in trailing)
    return f"{text.rstrip()}\n\n{lines}" if text.strip() else lines


def _marker(attachment: Attachment, name: str) -> str:
    if attachment.kind == "item":
        return f"[Message: {name}]"
    if attachment.kind == "reference":
        return f"[Link: {name}]"
    if attachment.is_inline:
        return f"[Image: {name}]"
    return f"[File: {name}]"

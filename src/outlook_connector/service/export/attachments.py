"""Attachment selection policy and safe file names (requirements v4 §10.1, research §3.5).

- Non-inline file attachments are exported.
- Item attachments (forwarded mail) are exported as .eml.
- Inline images only when the rendered body (unique or full) references their cid:.
  Most inline images are signatures and quoted history and are skipped.
- Reference (cloud) attachments are listed, never downloaded.
"""

from __future__ import annotations

import re
from pathlib import PurePath

from outlook_connector.domain.models import Attachment

_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {
    "con",
    "prn",
    "aux",
    "nul",
    *(f"com{i}" for i in range(1, 10)),
    *(f"lpt{i}" for i in range(1, 10)),
}
MAX_NAME = 120


def wanted(attachment: Attachment, *, content_id: str | None, body_html: str | None) -> bool:
    """Whether the export includes this attachment.

    Assumes (not re-checked here): ``content_id`` and ``body_html`` belong to the attachment's message
    (the orchestrator looks them up).
    """
    if attachment.kind == "item":
        return True
    if attachment.kind != "file":
        return False
    if not attachment.is_inline:
        return True
    return bool(content_id and body_html and f"cid:{content_id}" in body_html)


def listed(attachment: Attachment) -> bool:
    """Attachments mentioned in the TXT when files are not included: the ones a reader cares about."""
    return attachment.kind in ("item", "reference") or (
        attachment.kind == "file" and not attachment.is_inline
    )


def safe_name(name: str | None, *, fallback: str, eml: bool = False) -> str:
    cleaned = _UNSAFE.sub("_", (name or "").strip()).strip(" .") or fallback
    if eml and not cleaned.lower().endswith(".eml"):
        cleaned += ".eml"
    stem, suffix = PurePath(cleaned).stem, PurePath(cleaned).suffix
    if stem.lower() in _RESERVED:
        stem = f"_{stem}"
    if len(stem) + len(suffix) > MAX_NAME:
        stem = stem[: MAX_NAME - len(suffix)]
    return stem + suffix


def dedupe(name: str, taken: set[str]) -> str:
    """Case-insensitive de-duplication: report.pdf, report (2).pdf, ..."""
    stem, suffix = PurePath(name).stem, PurePath(name).suffix
    candidate, counter = name, 2
    while candidate.lower() in taken:
        candidate = f"{stem} ({counter}){suffix}"
        counter += 1
    taken.add(candidate.lower())
    return candidate


def size_label(size: int | None) -> str:
    if size is None:
        return "size unknown"
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.0f} KB"
    return f"{size / (1024 * 1024):.1f} MB"

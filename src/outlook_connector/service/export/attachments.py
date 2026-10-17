"""Attachment selection policy and safe file names (requirements §11, research §3.6).

- Non-inline file attachments are exported.
- Item attachments (forwarded mail) are exported as .eml.
- Inline images only when the rendered body (unique or full) references their cid:.
  If the content id or rendered body is unavailable, include the image rather than drop it.
- Reference (cloud) attachments are listed, never downloaded.
"""

from __future__ import annotations

import re
from pathlib import PurePath

from outlook_connector.domain.models import Attachment

_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
# Win32 device names, superscript digits included ("Naming Files, Paths, and Namespaces").
_RESERVED = {
    "con",
    "prn",
    "aux",
    "nul",
    *(f"{port}{digit}" for port in ("com", "lpt") for digit in "0123456789¹²³"),
}
MAX_NAME = 120


def wanted(attachment: Attachment, *, body_html: str | None) -> bool:
    """Whether the export includes this attachment.

    Assumes (not re-checked here): body_html belongs to the attachment's message. If either it or
    the attachment content id is unavailable, inline inclusion fails open.
    """
    if attachment.kind == "item":
        return True
    if attachment.kind != "file":
        return False
    if not attachment.is_inline:
        return True
    if not attachment.content_id or body_html is None:
        return True
    return f"cid:{attachment.content_id}" in body_html


def listed(attachment: Attachment) -> bool:
    """Attachments mentioned in the TXT when files are not included: the ones a reader cares about."""
    return attachment.kind in ("item", "reference") or (
        attachment.kind == "file" and not attachment.is_inline
    )


def safe_name(name: str | None, *, fallback: str, eml: bool = False) -> str:
    cleaned = _UNSAFE.sub("_", (name or "").strip()).strip(" .") or fallback
    if eml and not cleaned.lower().endswith(".eml"):
        cleaned += ".eml"
    # Win32 reserves a device name followed by any extension (CON.backup.txt), so check before the first dot.
    if cleaned.split(".", 1)[0].rstrip(" ").lower() in _RESERVED:
        cleaned = f"_{cleaned}"
    stem, suffix = PurePath(cleaned).stem, PurePath(cleaned).suffix
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

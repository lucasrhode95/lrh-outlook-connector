"""Attachment selection policy and safe file names (requirements §11, research §3.6).

- File and item attachments (forwarded mail) are exported. Inline files follow typed source image
  references; unknown source structure or content ids retain the file with no assumed body position.
- Reference (cloud) attachments are listed, never downloaded.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from pathlib import PurePath

from outlook_connector.domain.body import BodyPart, ImagePart
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


def wanted(attachment: Attachment, *, parts: Sequence[BodyPart] | None) -> bool:
    """Whether source metadata and typed body references include this file.

    Assumes (not re-checked here): ``parts`` comes from the selected source body of the same message.
    Unavailable structure or a missing content id cannot establish absence, so the file is kept.
    """
    if attachment.kind == "item":
        return True
    if attachment.kind != "file":
        return False
    if not attachment.is_inline or not attachment.content_id or parts is None:
        return True
    return any(isinstance(part, ImagePart) and part.content_id == attachment.content_id for part in parts)


def listed(attachment: Attachment) -> bool:
    """Known attachments listed in JSONL when their file bytes are not included."""
    return attachment.kind in ("file", "item", "reference")


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

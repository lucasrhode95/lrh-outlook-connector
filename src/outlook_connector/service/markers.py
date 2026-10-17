"""Attachment markers in message text (requirements §10): every file a message carries is named in its text.

One rule serves get_message, get_conversation and exports:

- An inline image the HTML body references by ``cid:`` reads ``[Image: name]`` where Exchange's text body
  shows it, which is its alt text in brackets (research §3.6). A placeholder is replaced only when the text
  holds exactly as many ``[alt]`` as the HTML has images with that alt text, so the k-th occurrence is the
  k-th image. Otherwise (no alt text, the same bracketed words in the prose, or text that differs) the
  marker is listed after the text, so a name never lands on the wrong place.
- Files, forwarded messages and cloud links follow the text: ``[File: name]``, ``[Message: name.eml]`` and
  ``[Link: name]``. Graph v1.0 gives no address for a cloud link.
- Names come from the attachment name, made safe as a file name and unique within the message
  (``report.pdf``, ``report (2).pdf``). ``download_attachment`` takes these names.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from html.parser import HTMLParser

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
    html: str | None,
    names: Mapping[str, str] | None = None,
) -> str:
    """``text`` (Exchange's text body) with a marker for each file the message carries.

    ``html`` is the HTML of the same body (full or unique), or None when it was not read; then inline images
    keep Exchange's placeholder and only the other markers are added. ``names`` overrides marker names (an
    export uses the names of the files it wrote). Inline images this body does not reference get no marker.

    Assumes (not re-checked here): ``attachments`` is the listing of the message ``text`` and ``html`` come
    from.
    """
    names = dict(names) if names is not None else file_names(attachments)
    by_cid = {
        a.content_id: a
        for a in attachments
        if a.kind == "file" and a.is_inline and a.content_id and a.id in names
    }
    images = _images(html) if html is not None else []
    by_alt: dict[str, list[str | None]] = {}  # alt text -> content id of each image carrying it, in order
    for cid, alt in images:
        if alt:
            by_alt.setdefault(alt, []).append(cid)
    edits: list[tuple[int, int, str]] = []  # (start, end, marker) in ``text``
    placed: set[str] = set()
    for alt, cids in by_alt.items():
        found = _occurrences(text, f"[{alt}]")
        if len(found) != len(cids):
            continue  # ambiguous: these images are listed after the text
        for start, cid in zip(found, cids, strict=True):
            attachment = by_cid.get(cid) if cid else None
            if attachment is not None:
                edits.append((start, start + len(alt) + 2, _marker(attachment, names[attachment.id])))
                placed.add(attachment.id)
    out = text
    for start, end, marker in sorted(edits, reverse=True):
        out = out[:start] + marker + out[end:]
    unplaced = [by_cid[cid] for cid, _ in images if cid in by_cid and by_cid[cid].id not in placed]
    trailing = list({a.id: a for a in unplaced}.values())
    trailing += [a for a in attachments if a.id in names and not (a.kind == "file" and a.is_inline)]
    if not trailing:
        return out
    lines = "\n".join(_marker(a, names[a.id]) for a in trailing)
    return f"{out.rstrip()}\n\n{lines}" if out.strip() else lines


def _marker(attachment: Attachment, name: str) -> str:
    if attachment.kind == "item":
        return f"[Message: {name}]"
    if attachment.kind == "reference":
        return f"[Link: {name}]"
    if attachment.is_inline:
        return f"[Image: {name}]"
    return f"[File: {name}]"


def _occurrences(text: str, needle: str) -> list[int]:
    """Start of each non-overlapping occurrence of ``needle`` in ``text``."""
    found: list[int] = []
    start = text.find(needle)
    while start >= 0:
        found.append(start)
        start = text.find(needle, start + len(needle))
    return found


class _ImageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.images: list[tuple[str | None, str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag != "img":
            return
        values = {key: value or "" for key, value in attrs}
        src = values.get("src", "").strip()
        cid = src[4:] if src[:4].lower() == "cid:" else None
        self.images.append((cid, " ".join(values.get("alt", "").split())))


def _images(html: str) -> list[tuple[str | None, str]]:
    """(content id, or None for another source, and alt text) of each image, in document order. Images
    of any source count, since Exchange shows each one's alt text."""
    parser = _ImageParser()
    parser.feed(html)
    parser.close()
    return parser.images

"""Attachment markers in message text (requirements §10): every file a message carries is named in its text.

One rule serves get_message, get_conversation and exports:

- An inline image the HTML body references by ``cid:`` reads ``[Image: name]`` where Exchange's text body
  shows it, which is its alt text in brackets (research §3.6). Placeholders are matched in document order;
  an image whose placeholder is not found (no alt text, or text that differs) is listed after the text, so
  a name never lands on the wrong image.
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
    out, position = text, 0
    placed: set[str] = set()
    unplaced: list[Attachment] = []
    for cid, alt in _images(html) if html is not None else []:
        attachment = by_cid.get(cid)
        if attachment is None:
            continue
        marker = _marker(attachment, names[attachment.id])
        found = out.find(f"[{alt}]", position) if alt else -1
        if found >= 0:
            out = out[:found] + marker + out[found + len(alt) + 2 :]
            position = found + len(marker)
            placed.add(attachment.id)
        elif attachment.id not in placed and all(a.id != attachment.id for a in unplaced):
            unplaced.append(attachment)
    trailing = [a for a in unplaced if a.id not in placed]
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


class _ImageParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.images: list[tuple[str, str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag != "img":
            return
        values = {key: value or "" for key, value in attrs}
        src = values.get("src", "").strip()
        if src[:4].lower() == "cid:":
            self.images.append((src[4:], " ".join(values.get("alt", "").split())))


def _images(html: str) -> list[tuple[str, str]]:
    """(content id, alt text) of each ``cid:`` image, in document order."""
    parser = _ImageParser()
    parser.feed(html)
    parser.close()
    return parser.images

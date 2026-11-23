"""One inline-image conversion for native signatures and draft bodies.

Data URIs in quoted ``src`` attributes — and, in a draft body, also absolute local file paths and
``file:///`` URLs — become ``InlineImage`` values with ``cid:`` references, saved with the draft as
IsInline attachments. Source rules and error wording are stated here once for both callers; the
write itself stays in ``writes.py`` and ``ows_mapping.py``.
"""

from __future__ import annotations

import base64
import binascii
import mimetypes
import re
import uuid
from html import unescape
from pathlib import Path
from typing import Literal
from urllib.parse import unquote, unquote_to_bytes, urlsplit

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import MAX_SEND_BYTES, InlineImage
from outlook_connector.service.passive_html import StartTag, passive_start_tags

Kind = Literal["signature", "body"]
_SUBJECT: dict[Kind, str] = {"signature": "A native signature", "body": "The draft body"}
_SOURCE: dict[Kind, str] = {"signature": "Signature", "body": "Draft body"}
_IMAGE_EXTENSIONS = {
    "image/jpeg": "jpg",
    "image/png": "png",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/webp": "webp",
    "image/tiff": "tif",
    "image/svg+xml": "svg",
    "image/x-icon": "ico",
}
_DATA_IMAGE = re.compile(
    r"""(?P<prefix>\bsrc\s*=\s*)(?P<quote>["'])(?P<uri>data:image/[^"'<> \t\r\n]+)(?P=quote)""",
    re.IGNORECASE,
)
_HTML_ATTRIBUTE = re.compile(r"""(?P<name>[^\s/=>]+)\s*=\s*(?P<value>"[^"]*"|'[^']*'|[^\s>]+)""")
# A quoted src naming a local file: a file:/// URL, a Windows drive path or a UNC path.
_LOCAL_IMAGE = re.compile(
    r"""(?P<prefix>\bsrc\s*=\s*)(?P<quote>["'])"""
    r"""(?P<path>file:///[^"'<>\r\n]+|\\\\[^"'<>\r\n]+|[A-Za-z]:[\\/][^"'<>\r\n]*)"""
    r"""(?P=quote)""",
    re.IGNORECASE,
)
# Magic bytes of the image types a data URI or a local file may carry; svg is the one text format,
# taken from the file name because its text has no magic bytes.
_MAGIC: tuple[tuple[bytes, str], ...] = (
    (b"\x89PNG\r\n\x1a\n", "image/png"),
    (b"\xff\xd8\xff", "image/jpeg"),
    (b"GIF87a", "image/gif"),
    (b"GIF89a", "image/gif"),
    (b"BM", "image/bmp"),
    (b"II*\x00", "image/tiff"),
    (b"MM\x00*", "image/tiff"),
    (b"\x00\x00\x01\x00", "image/x-icon"),
)


def check_image_sources(tags: list[StartTag], kind: Kind) -> None:
    """Inline conversion needs a quoted data URI. HTMLParser accepts unquoted ones, so the source
    text is read; whole attribute values are matched, so text inside another quoted attribute is
    not a src. ``kind`` names the HTML in the error message.
    """
    for tag in tags:
        for attribute in _HTML_ATTRIBUTE.finditer(tag.source):
            if attribute.group("name").lower() != "src":
                continue
            source = attribute.group("value")
            decoded = unescape(source.strip("\"'"))
            if (
                decoded.strip().lower().startswith("data:image/")
                and _DATA_IMAGE.fullmatch(f"src={source}") is None
            ):
                raise InvalidRequest(
                    f"{_SOURCE[kind]} image data URIs must use a quoted src attribute without whitespace."
                )


def data_uri_images(page: str, kind: Kind) -> tuple[str, list[InlineImage]]:
    """Convert every quoted data:image URI in a page into an inline image referenced by cid:.

    Assumes (not re-checked here): the page passed the passive-HTML checker and
    ``check_image_sources`` for its kind, so every data URI matches ``_DATA_IMAGE``.
    """
    images: list[InlineImage] = []

    def replace(match: re.Match[str]) -> str:
        data, media_type = _decode_data_uri(match.group("uri"), kind)
        images.append(_inline_image(data, media_type, kind, len(images) + 1))
        return _cid_source(match, images[-1].content_id)

    return _DATA_IMAGE.sub(replace, page), images


def body_images(page: str) -> tuple[str, list[InlineImage]]:
    """Entry point for a draft body: convert its own image sources into inline images.

    Data URIs, absolute local paths (Windows drive, UNC or ``file:///``) and nothing else become
    inline images; other sources are left as written. A source that is missing, unreadable, over
    the send limit alone or not a recognized image raises before any draft is saved, naming the
    file. A local image keeps its file name; a data URI gets a synthetic one.

    Assumes (not re-checked here): ``page`` already passed ``passive_start_tags``.
    """
    check_image_sources(passive_start_tags(page), "body")
    page, images = data_uri_images(page, "body")

    def replace(match: re.Match[str]) -> str:
        path = _local_path(unescape(match.group("path")))
        data, media_type = _read_image_file(path)
        images.append(
            InlineImage(
                name=path.name,
                content_type=media_type,
                content_id=f"body-{uuid.uuid4().hex}@outlook-connector",
                content=data,
            )
        )
        return _cid_source(match, images[-1].content_id)

    return _LOCAL_IMAGE.sub(replace, page), images


def _decode_data_uri(uri: str, kind: Kind) -> tuple[bytes, str]:
    """Decode one data URI to its bytes and its normalized media type."""
    header, separator, payload = uri[5:].partition(",")
    media_type, *parameters = header.split(";")
    if not separator or not media_type.lower().startswith("image/"):
        raise InvalidRequest(f"{_SUBJECT[kind]} contains an invalid image data URI.")
    try:
        data = (
            base64.b64decode(payload, validate=True)
            if any(parameter.lower() == "base64" for parameter in parameters)
            else unquote_to_bytes(payload)
        )
    except (ValueError, binascii.Error):
        raise InvalidRequest(f"{_SUBJECT[kind]} contains an invalid image data URI.") from None
    if not data:
        raise InvalidRequest(f"{_SUBJECT[kind]} contains an empty image.")
    return data, media_type.lower()


def _inline_image(data: bytes, media_type: str, kind: Kind, number: int) -> InlineImage:
    """One converted data-URI image: a synthetic name and a cid unique to this draft."""
    return InlineImage(
        name=f"{kind}-image-{number}.{_IMAGE_EXTENSIONS.get(media_type, 'img')}",
        content_type=media_type,
        content_id=f"{kind}-{uuid.uuid4().hex}@outlook-connector",
        content=data,
    )


def _cid_source(match: re.Match[str], content_id: str) -> str:
    """The matched src attribute with its value replaced by the image's cid: reference."""
    return f"{match.group('prefix')}{match.group('quote')}cid:{content_id}{match.group('quote')}"


def _local_path(source: str) -> Path:
    """A ``file:///`` URL to its percent-decoded path; a plain Windows or UNC path is used as written."""
    if source[:8].lower() == "file:///":
        path = unquote(urlsplit(source).path)
        # The Windows drive form carries a slash before the colon (file:///C:/dir/x.png).
        return Path(path[1:] if re.match(r"/[A-Za-z]:", path) else path)
    return Path(source)


def _read_image_file(path: Path) -> tuple[bytes, str]:
    """Read one local image file and detect its type from its bytes; only svg is taken from the
    file name. Errors name the file."""
    try:
        if not path.is_file():
            raise InvalidRequest(f"Cannot read body image '{path}': it is not a file.")
        size = path.stat().st_size
        if size > MAX_SEND_BYTES:
            raise InvalidRequest(
                f"Body image '{path}' is {size:,} bytes; a draft over 35 MiB cannot be sent, "
                "so none was created."
            )
        data = path.read_bytes()
    except OSError as exc:
        raise InvalidRequest(f"Cannot read body image '{path}': {exc}") from None
    media_type = _sniffed_type(data)
    if media_type is None and mimetypes.guess_type(path.name)[0] == "image/svg+xml":
        media_type = "image/svg+xml"
    if media_type is None:
        raise InvalidRequest(f"Body image '{path}' is not a recognized image file.")
    return data, media_type


def _sniffed_type(data: bytes) -> str | None:
    """The media type the image's magic bytes identify, or None when they match no known type."""
    for magic, media_type in _MAGIC:
        if data.startswith(magic):
            return media_type
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    return None

"""Native signature domain behavior and fresh default resolution (W8)."""

from __future__ import annotations

import base64
import binascii
import re
import uuid
from dataclasses import dataclass
from html import escape, unescape
from html.parser import HTMLParser
from typing import Literal
from urllib.parse import unquote_to_bytes

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import AccountMismatch, InvalidRequest, NotFound, Upstream
from outlook_connector.domain.models import (
    MAX_BODY_CHARS,
    InlineImage,
    SignatureDetails,
    SignatureInfo,
    SignatureList,
    SignatureWriteResult,
)
from outlook_connector.remote.ports import SignatureContents, SignatureSettings, SignatureStore

_DATA_IMAGE = re.compile(
    r"""(?P<prefix>\bsrc\s*=\s*)(?P<quote>["'])(?P<uri>data:image/[^"'<> \t\r\n]+)(?P=quote)""",
    re.IGNORECASE,
)
_HTML_ATTRIBUTE = re.compile(r"""(?P<name>[^\s/=>]+)\s*=\s*(?P<value>"[^"]*"|'[^']*'|[^\s>]+)""")
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


@dataclass(frozen=True)
class ResolvedSignature:
    html: str
    images: list[InlineImage]


class Signatures:
    """Account-bound signature management; settings and contents are never cached."""

    def __init__(self, store: SignatureStore, account: Account) -> None:
        self._store = store
        self._account = account

    async def list_signatures(self) -> SignatureList:
        """Entry point: list exact names, defaults and whether each has readable contents."""
        self._check_account()
        settings = await self._store.settings()
        entries = []
        contents_by_name: dict[str, SignatureContents | None] = {}
        for name in settings.names:
            contents = await self._store.contents(name, settings)
            contents_by_name[name] = contents
            entries.append(SignatureInfo(name=name, readable=_readable(contents)))
        await self._unchanged(settings, contents_by_name)
        return SignatureList(
            signatures=entries,
            new_default=settings.new_default,
            reply_default=settings.reply_default,
        )

    async def get_signature(self, name: str) -> SignatureDetails:
        """Entry point: validate one public name and return its current HTML/text."""
        validate_name(name)
        self._check_account()
        settings = await self._store.settings()
        if name not in settings.names:
            raise NotFound("That native signature does not exist.")
        contents = await self._store.contents(name, settings)
        await self._unchanged(settings, {name: contents})
        html, text = _readable_contents(contents)
        return SignatureDetails(name=name, html=html, text=text)

    async def create_signature(self, name: str, html: str) -> SignatureWriteResult:
        """Entry point: validate exact name and passive HTML, then send one write."""
        validate_name(name)
        safe_html = _validate_html(html)
        self._check_account()
        settings = await self._store.settings()
        if name in settings.names:
            raise InvalidRequest("That native signature already exists; use update_signature.")
        await self._store.create(name, safe_html, _html_text(safe_html), settings)
        return SignatureWriteResult(status="created", name=name)

    async def update_signature(self, name: str, html: str) -> SignatureWriteResult:
        """Entry point: validate exact name and passive HTML, then send one write."""
        validate_name(name)
        safe_html = _validate_html(html)
        self._check_account()
        settings = await self._store.settings()
        if name not in settings.names:
            raise NotFound("That native signature does not exist; use create_signature.")
        await self._store.update(name, safe_html, _html_text(safe_html), settings)
        return SignatureWriteResult(status="updated", name=name)

    async def delete_signature(self, name: str) -> SignatureWriteResult:
        """Entry point: validate exact name and existence, then send one write."""
        validate_name(name)
        self._check_account()
        settings = await self._store.settings()
        if name not in settings.names:
            raise NotFound("That native signature does not exist.")
        await self._store.delete(name)
        return SignatureWriteResult(status="deleted", name=name)

    async def set_default_signature(
        self, name: str | None, which: Literal["new", "reply", "both"]
    ) -> SignatureWriteResult:
        """Entry point: validate the requested default and write the selected setting once."""
        if which not in {"new", "reply", "both"}:
            raise InvalidRequest("for must be new, reply or both.")
        if name is not None:
            validate_name(name)
        self._check_account()
        settings = await self._store.settings()
        if name is not None and name not in settings.names:
            raise NotFound("That native signature does not exist.")
        await self._store.set_default(name, which, settings)
        return SignatureWriteResult(status="default_set", name=name, for_type=which)

    async def resolve_for_draft(self, name: str | None, *, is_reply: bool) -> ResolvedSignature | None:
        """Resolve the selected native signature from fresh settings before composing a draft."""
        if name is not None:
            validate_name(name)
        self._check_account()
        settings = await self._store.settings()
        selected = (
            name if name is not None else (settings.reply_default if is_reply else settings.new_default)
        )
        if selected is None:
            await self._unchanged(settings)
            return None
        if selected not in settings.names:
            if name is not None:
                raise NotFound("That native signature does not exist.")
            raise Upstream("The configured native signature default no longer exists.")
        contents = await self._store.contents(selected, settings)
        await self._unchanged(settings, {selected: contents})
        html, _text = _readable_contents(contents)
        validated = _validate_html(html)
        block, images = _inline_images(validated)
        signature = f'<div id="Signature" data-signature-name="{escape(selected, quote=True)}">{block}</div>'
        return ResolvedSignature(signature, images)

    def _check_account(self) -> None:
        claims = self._store.account()
        if (claims.get("tid"), claims.get("oid")) != (self._account.tenant_id, self._account.object_id):
            raise AccountMismatch(
                "The write sign-in belongs to a different Microsoft account than its read account."
            )

    async def _unchanged(
        self,
        before: SignatureSettings,
        contents: dict[str, SignatureContents | None] | None = None,
    ) -> None:
        after = await self._store.settings()
        if before.revision != after.revision:
            raise Upstream(
                "Native signature settings changed while they were being read; reread them before composing."
            )
        for name, original in (contents or {}).items():
            current = await self._store.contents(name, before)
            if _content_revision(original) != _content_revision(current):
                raise Upstream("Native signature contents changed while being read; reread before composing.")


def validate_name(name: str) -> None:
    """Validate a public signature name once at its service entry point; preserve its exact spelling."""
    if not name or not name.strip():
        raise InvalidRequest("A native signature name must not be empty.")
    if "," in name:
        raise InvalidRequest("Native signature names cannot contain commas.")


def _content_revision(contents: SignatureContents | None) -> str | None:
    return contents.revision if contents is not None else None


def _readable(contents: SignatureContents | None) -> bool:
    return contents is not None and bool((contents.html or "").strip() or (contents.text or "").strip())


def _readable_contents(contents: SignatureContents | None) -> tuple[str, str]:
    if not _readable(contents):
        raise Upstream("That native signature has no readable HTML or text contents.")
    assert contents is not None
    html = contents.html or ""
    text = contents.text or ""
    if not html.strip():
        html = _text_html(text)
    if not text.strip():
        text = _html_text(html)
    return html, text


def _validate_html(value: str) -> str:
    if not value.strip():
        raise InvalidRequest("A native signature cannot be empty.")
    if len(value) > MAX_BODY_CHARS:
        raise InvalidRequest(f"A native signature must be at most {MAX_BODY_CHARS} characters.")
    parser = _PassiveSignatureHtml()
    parser.feed(value)
    parser.close()
    return value


def _html_text(value: str) -> str:
    parser = _TextFromHtml()
    parser.feed(value)
    parser.close()
    return re.sub(r"\n{3,}", "\n\n", "".join(parser.parts)).strip()


def _text_html(value: str) -> str:
    lines = escape(value).replace("\u00a0", "&nbsp;").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return "<div>" + "<br>".join(_keep_spaces(line) for line in lines) + "</div>"


def _keep_spaces(line: str) -> str:
    line = line.replace("\t", "&nbsp;" * 4)
    line = re.sub(r" {2,}", lambda run: "&nbsp;" * (len(run.group()) - 1) + " ", line)
    return "&nbsp;" + line[1:] if line.startswith(" ") else line


def _inline_images(value: str) -> tuple[str, list[InlineImage]]:
    images: list[InlineImage] = []

    def replace(match: re.Match[str]) -> str:
        uri = unescape(match.group("uri"))
        header, separator, payload = uri[5:].partition(",")
        media_type, *parameters = header.split(";")
        if not separator or not media_type.lower().startswith("image/"):
            raise InvalidRequest("A native signature contains an invalid image data URI.")
        try:
            data = (
                base64.b64decode(payload, validate=True)
                if any(parameter.lower() == "base64" for parameter in parameters)
                else unquote_to_bytes(payload)
            )
        except (ValueError, binascii.Error):
            raise InvalidRequest("A native signature contains an invalid image data URI.") from None
        if not data:
            raise InvalidRequest("A native signature contains an empty image.")
        normalized_type = media_type.lower()
        extension = _IMAGE_EXTENSIONS.get(normalized_type, "img")
        cid = f"signature-{uuid.uuid4().hex}@outlook-connector"
        images.append(
            InlineImage(
                name=f"signature-image-{len(images) + 1}.{extension}",
                content_type=normalized_type,
                content_id=cid,
                content=data,
            )
        )
        return f"{match.group('prefix')}{match.group('quote')}cid:{cid}{match.group('quote')}"

    converted = _DATA_IMAGE.sub(replace, value)
    return converted, images


class _PassiveSignatureHtml(HTMLParser):
    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "iframe", "object", "embed", "form", "input", "button", "textarea", "select"}:
            raise InvalidRequest(f"Active HTML content is not allowed in a signature: {tag}.")
        for name, value in attrs:
            compact = re.sub(r"[\s\x00-\x1f]+", "", unescape(value or "")).lower()
            if name.startswith("on") or compact.startswith(("javascript:", "vbscript:")):
                raise InvalidRequest("Active HTML URLs and event handlers are not allowed in a signature.")
        # HTMLParser accepts unquoted attributes, but inline conversion requires a quoted data URI.
        # Match whole attribute values so text inside another quoted attribute is not treated as src.
        for attribute in _HTML_ATTRIBUTE.finditer(self.get_starttag_text() or ""):
            if attribute.group("name").lower() != "src":
                continue
            source = attribute.group("value")
            if (
                unescape(source.strip("\"'")).lower().startswith("data:image/")
                and _DATA_IMAGE.fullmatch(f"src={source}") is None
            ):
                raise InvalidRequest(
                    "Signature image data URIs must use a quoted src attribute without whitespace."
                )

    handle_startendtag = handle_starttag


class _TextFromHtml(HTMLParser):
    BLOCKS = {
        "address",
        "article",
        "blockquote",
        "br",
        "div",
        "h1",
        "h2",
        "h3",
        "li",
        "ol",
        "p",
        "pre",
        "table",
        "tr",
        "ul",
    }

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in self.BLOCKS:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in self.BLOCKS:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

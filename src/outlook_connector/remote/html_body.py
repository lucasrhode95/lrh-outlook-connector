"""Source HTML to ordered text and image parts; no rendered-text metadata recovery."""

from __future__ import annotations

import re
from html.parser import HTMLParser
from urllib.parse import unquote

from outlook_connector.domain.body import BodyPart, ImagePart, TextPart

_VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "wbr"}
_HIDDEN = {"head", "script", "style", "template", "title"}
_BLOCK = {"address", "article", "aside", "div", "footer", "header", "section", "tr"}
_PARAGRAPH = {"p", "h1", "h2", "h3", "h4", "h5", "h6"}


def parse_html(source: str) -> list[BodyPart] | None:
    """Extract body order and image references from HTML explicitly identified as such by the source.

    This is a presentation transform: semantic blocks, lists, links, tables, quotes and preformatted
    text are kept; browser/CSS layout is not reproduced. Unsupported image/layout structures or malformed
    declarations return None, so callers retain the server text and cannot assume an inline file is absent.
    No remote image or stylesheet is fetched.
    """
    parser = _BodyParser()
    try:
        parser.feed(source)
        parser.close()
    except AssertionError:  # HTMLParser rejects malformed marked sections before calling a handler.
        return None
    return None if parser.unsupported else parser.parts


class _BodyParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[BodyPart] = []
        self.unsupported = False
        self._hidden: list[str] = []
        self._pre = 0
        self._quotes = 0
        self._lists: list[int | None] = []
        self._links: list[str | None] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = dict(attrs)
        if self._hidden or tag in _HIDDEN or _is_hidden(values):
            if tag not in _VOID:
                self._hidden.append(tag)
            return
        if _unsupported_structure(tag, values):
            self.unsupported = True
        if tag == "img":
            src = (values.get("src") or "").strip()
            cid = unquote(src[4:]) if src[:4].casefold() == "cid:" else None
            self._quote_prefix()
            self.parts.append(ImagePart(cid or None, values.get("alt") or ""))
        elif tag == "br":
            self._raw_text("\n")
        elif tag == "hr":
            self._break(2)
            self._text("---")
            self._break(2)
        elif tag in _PARAGRAPH:
            self._break(2)
        elif tag in _BLOCK:
            self._break(1)
        elif tag == "blockquote":
            self._break(2)
            self._quotes += 1
        elif tag == "pre":
            self._break(2)
            self._pre += 1
        elif tag in {"ul", "ol"}:
            self._break(1)
            self._lists.append(0 if tag == "ol" else None)
        elif tag == "li":
            self._break(1)
            bullet = "- "
            if self._lists and self._lists[-1] is not None:
                self._lists[-1] += 1
                bullet = f"{self._lists[-1]}. "
            self._text("  " * max(0, len(self._lists) - 1) + bullet)
        elif tag == "a":
            href = values.get("href") or ""
            self._links.append(
                href if href.casefold().startswith(("http://", "https://", "mailto:")) else None
            )

    def handle_endtag(self, tag: str) -> None:
        if self._hidden:
            if tag in self._hidden:
                index = len(self._hidden) - 1 - self._hidden[::-1].index(tag)
                del self._hidden[index:]
            return
        if tag in _PARAGRAPH:
            self._break(2)
        elif tag in _BLOCK or tag == "li":
            self._break(1)
        elif tag in {"td", "th"}:
            self._raw_text("\t")
        elif tag == "blockquote":
            self._quotes = max(0, self._quotes - 1)
            self._break(2)
        elif tag == "pre":
            self._pre = max(0, self._pre - 1)
            self._break(2)
        elif tag in {"ul", "ol"}:
            if self._lists:
                self._lists.pop()
            self._break(1)
        elif tag == "a" and self._links:
            href = self._links.pop()
            if href:
                self._text(f" <{href}>")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.handle_starttag(tag, attrs)
        if tag not in _VOID:
            self.handle_endtag(tag)

    def handle_data(self, data: str) -> None:
        if not self._hidden:
            self._text(data if self._pre else re.sub(r"\s+", " ", data))

    def handle_comment(self, data: str) -> None:
        # Conditional Outlook/VML image branches do not have one unambiguous HTML position.
        if self._hidden:
            return
        if data.lstrip().casefold().startswith("[if") and re.search(
            r"<(?:img|v:imagedata|svg|picture|object|embed)\b", data, re.IGNORECASE
        ):
            self.unsupported = True

    def unknown_decl(self, data: str) -> None:
        self.unsupported = True

    def _text(self, value: str) -> None:
        if not self._pre and value.startswith(" "):
            previous = self.parts[-1] if self.parts else None
            if previous is None or (isinstance(previous, TextPart) and previous.text[-1:].isspace()):
                value = value.lstrip(" ")
        for line in value.splitlines(keepends=True):
            if line.strip():
                self._quote_prefix()
            self._raw_text(line)

    def _quote_prefix(self) -> None:
        previous = self.parts[-1] if self.parts else None
        if self._quotes and (
            previous is None or (isinstance(previous, TextPart) and previous.text.endswith("\n"))
        ):
            self._raw_text("> " * self._quotes)

    def _raw_text(self, value: str) -> None:
        if not value:
            return
        if self.parts and isinstance(self.parts[-1], TextPart):
            self.parts[-1] = TextPart(self.parts[-1].text + value)
        else:
            self.parts.append(TextPart(value))

    def _break(self, lines: int) -> None:
        if not self.parts:
            return
        previous = self.parts[-1]
        trailing = 0
        if isinstance(previous, TextPart):
            value = previous.text.rstrip(" \t")
            self.parts[-1] = TextPart(value)
            trailing = len(value) - len(value.rstrip("\n"))
        self._raw_text("\n" * max(0, lines - trailing))


def _is_hidden(attrs: dict[str, str | None]) -> bool:
    if "hidden" in attrs:
        return True
    for declaration in (attrs.get("style") or "").split(";"):
        name, _, value = declaration.partition(":")
        value = value.partition("!")[0].strip().casefold()
        if (name.strip().casefold(), value) in {("display", "none"), ("visibility", "hidden")}:
            return True
    return False


def _unsupported_structure(tag: str, attrs: dict[str, str | None]) -> bool:
    if tag in {
        "plaintext",
        "textarea",
        "xmp",
        "iframe",
        "noembed",
        "noframes",
        "svg",
        "picture",
        "object",
        "embed",
        "v:imagedata",
    }:
        return True
    if attrs.get("srcset") or re.search(r"\burl\s*\([^)]*cid:", attrs.get("style") or "", re.IGNORECASE):
        return True
    for name in ("background", "src", "href", "data"):
        if tag == "img" and name == "src":
            continue
        if (attrs.get(name) or "").lstrip().casefold().startswith("cid:"):
            return True
    return False

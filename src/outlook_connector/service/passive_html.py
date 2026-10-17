"""Passive HTML: the rules every connector applies to HTML it sends.

Each HTML input is checked once, where it enters. A complete passive fragment cannot change how another
complete fragment parses, so composing fragments needs no second check. The HTML is sent as written:
nothing is parsed into another form here. A connector's policy on top of these rules, such as which images
a Teams message may carry, lives in that connector's own modules.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from html import unescape
from html.parser import HTMLParser

from outlook_connector.domain.errors import InvalidRequest

ACTIVE_TAGS = frozenset(
    {"script", "iframe", "object", "embed", "form", "input", "button", "textarea", "select"}
)
# Elements a browser reads as text (or that never end, as plaintext does). This parser reads the markup
# after them as tags, so a fragment that ends after one would be checked as tags but rendered as text.
STATE_CHANGING_TAGS = frozenset({"plaintext", "title", "xmp", "noembed", "noframes", "noscript"})
_WHITESPACE = re.compile(r"[\s\x00-\x1f]+")


@dataclass(frozen=True)
class StartTag:
    """One start tag: its lower-case name, its attributes as the parser reads them, its source text, and
    where it starts (line, column) in the HTML."""

    name: str
    attributes: tuple[tuple[str, str | None], ...]
    source: str
    position: tuple[int, int]


def passive_start_tags(value: str) -> list[StartTag]:
    """Entry point for HTML the connector sends: refuse declarations, active content, elements that
    change how the rest of the markup is read, and unfinished markup.

    Returns the start tags, so a caller with more rules need not parse the HTML again. HTMLParser
    never reports an unfinished tag, quote, comment or script; it keeps them in ``rawdata``. The
    trailing space lets the parser flush text that ends in ``&`` (such as "Tom&Jerry") without hiding
    an unfinished construct.
    """
    parser = _StartTags()
    parser.feed(value + " ")
    if parser.declarations:
        # The parser swallows a declaration whole, while a browser ends an unknown one at its first ">":
        # markup inside it would be checked as text but rendered as tags.
        raise InvalidRequest("HTML declarations, CDATA sections and processing instructions are not allowed.")
    for tag in parser.tags:
        _refuse_tag(tag)
    if parser.rawdata:
        raise InvalidRequest("The HTML ends inside a tag, attribute, comment or script; complete it.")
    return parser.tags


def _refuse_tag(tag: StartTag) -> None:
    if tag.name in ACTIVE_TAGS:
        raise InvalidRequest(f"Active HTML content is not allowed: {tag.name}.")
    if tag.name in STATE_CHANGING_TAGS:
        raise InvalidRequest(
            f"HTML that changes how the rest of the markup is read is not allowed: {tag.name}."
        )
    for name, value in tag.attributes:
        compact = _WHITESPACE.sub("", unescape(value or "")).lower()
        if name.startswith("on") or compact.startswith(("javascript:", "vbscript:")):
            raise InvalidRequest("Active HTML URLs and event handlers are not allowed.")


class _StartTags(HTMLParser):
    """Collects the start tags the parser completes; ``rawdata`` keeps whatever it could not finish."""

    def __init__(self) -> None:
        super().__init__()
        self.tags: list[StartTag] = []
        self.declarations = 0  # doctype, CDATA and other <!...> declarations, and <?...> instructions

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.tags.append(StartTag(tag, tuple(attrs), self.get_starttag_text() or "", self.getpos()))

    handle_startendtag = handle_starttag

    def handle_decl(self, decl: str) -> None:
        self.declarations += 1

    def unknown_decl(self, data: str) -> None:
        self.declarations += 1

    def handle_pi(self, data: str) -> None:
        self.declarations += 1

"""Passive HTML: the rule set for HTML the connector sends in drafts and native signatures.

Each HTML input is checked once, where it enters: a draft body in ``writes`` and a signature in
``signatures``. A complete passive fragment cannot change how another complete fragment parses, so
composing them needs no second check.
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
    """One start tag: its lower-case name, its attributes as the parser reads them, and its source text."""

    name: str
    attributes: tuple[tuple[str, str | None], ...]
    source: str


def passive_start_tags(value: str) -> list[StartTag]:
    """Entry point for HTML the connector sends: refuse active content, elements that change how the rest of
    the markup is read, and unfinished markup.

    Returns the start tags, so a caller with more rules (signature image sources) need not parse the HTML
    again. HTMLParser never reports an unfinished tag, quote, comment or script; it keeps them in
    ``rawdata``. The trailing space lets the parser flush text that ends in ``&`` (such as "Tom&Jerry")
    without hiding an unfinished construct.
    """
    parser = _StartTags()
    parser.feed(value + " ")
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

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.tags.append(StartTag(tag, tuple(attrs), self.get_starttag_text() or ""))

    handle_startendtag = handle_starttag

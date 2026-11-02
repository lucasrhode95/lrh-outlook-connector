"""Shared primitive normalizers for Search v2, OWS and Graph calendar responses."""

from __future__ import annotations

import re
from datetime import UTC, datetime
from typing import Any
from urllib.parse import unquote, urlsplit

from outlook_connector.domain.models import Recipient

_TEAMS_JOIN = re.compile(r"https://teams\.microsoft\.com/l/meetup-join/\S+", re.IGNORECASE)


def text(value: Any) -> str | None:
    return value if isinstance(value, str) and value else None


def boolean(value: Any) -> bool | None:
    return value if isinstance(value, bool) else None


def datetime_value(value: Any) -> datetime | None:
    if isinstance(value, dict):
        value = value.get("DateTime") or value.get("dateTime")
    if not isinstance(value, str) or not value:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    return parsed.replace(tzinfo=UTC) if parsed.tzinfo is None else parsed


def recipient(value: Any) -> Recipient | None:
    if not isinstance(value, dict):
        return None
    email = value.get("EmailAddress")
    if isinstance(email, dict):
        return Recipient(name=text(email.get("Name")), address=text(email.get("EmailAddress")))
    return Recipient(name=text(value.get("Name")), address=text(email))


def location(value: Any) -> str | None:
    if isinstance(value, dict):
        return text(value.get("DisplayName")) or text(value.get("Name"))
    return text(value)


def teams_join_url(body: str | None) -> str | None:
    """The first Teams meeting join link in a description, without trailing punctuation."""
    if not body:
        return None
    match = _TEAMS_JOIN.search(body)
    return match.group(0).rstrip(".,);]") if match else None


def teams_chat_id_from_join_link(join_url: str | None) -> str | None:
    """Parse a meeting chat id from the join-link path; Microsoft does not document this format."""
    if not join_url:
        return None
    try:
        segments = urlsplit(join_url).path.split("/")
    except ValueError:
        return None
    for index, segment in enumerate(segments[:-1]):
        if segment == "meetup-join":
            candidate = unquote(segments[index + 1])
            if candidate.startswith("19:meeting_") and candidate.endswith("@thread.v2"):
                return candidate
    return None

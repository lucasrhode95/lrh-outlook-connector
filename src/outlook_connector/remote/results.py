"""Values the remote adapters and the service exchange: result pages, counts, signature settings,
move targets and the shared request cap. Wire formats stay in the adapters."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Literal

from outlook_connector.domain.errors import Failure
from outlook_connector.domain.models import (
    CalendarEvent,
    Message,
    MessageSummary,
)

BodyFormat = Literal["text", "html"]

MAX_CONCURRENT_REQUESTS = 4  # Shared Exchange Online request cap for service scheduling and transport


@dataclass
class FetchedMessages:
    """A batch fetch: every requested id is in exactly one of the two maps."""

    messages: dict[str, Message | None] = field(default_factory=dict)  # None: no longer on the server
    failed: dict[str, Failure] = field(default_factory=dict)  # not fetched (throttled, error): why


@dataclass(frozen=True)
class FolderCount:
    """A folder's date-window count and newest received timestamp, returned by one count request."""

    count: int
    newest_received_at: datetime | None


@dataclass
class CalendarSearchPage:
    """Normalized one-page Exchange event discovery response."""

    events: list[CalendarEvent]
    rows_received: int
    more_results_available: bool | None
    result_set_found: bool


@dataclass
class CalendarViewPage:
    """One page of the calendar view. ``next_link`` continues it; None means the window is exhausted."""

    events: list[CalendarEvent]
    next_link: str | None = None


@dataclass(frozen=True)
class FolderTarget:
    """A move target: a well-known folder by its alias, or any folder by its Graph id."""

    folder_id: str
    well_known: str | None = None


@dataclass
class FetchedSummaries:
    """A batch lookup: every requested id is in exactly one of the two maps."""

    summaries: dict[str, MessageSummary | None] = field(default_factory=dict)  # None: not on the server
    failed: dict[str, Failure] = field(default_factory=dict)  # not fetched (throttled, error): why


@dataclass(frozen=True)
class SignatureSettings:
    """Fresh native settings; opaque scope is passed back unchanged, or None if no records exist."""

    names: tuple[str, ...]
    new_default: str | None
    reply_default: str | None
    scope: Any
    revision: str


@dataclass(frozen=True)
class SignatureContents:
    """HTML and text contents plus a read-time revision for one exact signature name."""

    html: str | None
    text: str | None
    revision: str

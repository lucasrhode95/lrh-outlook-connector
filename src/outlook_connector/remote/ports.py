"""The ports the service depends on (architecture §6.1).

Adapters: remote/graph_mail.py implements MailReader, remote/ows.py implements MailWriter.
Swapping an adapter (another tenant, a policy change) must not change anything above this file.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Literal, Protocol

from outlook_connector.domain.errors import Failure
from outlook_connector.domain.models import (
    Attachment,
    DraftMessage,
    Folder,
    InboxRule,
    Message,
    MessageSummary,
    RuleChange,
    UserProfile,
)

BodyFormat = Literal["text", "html"]


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


class MailReader(Protocol):
    async def list_folders(self) -> list[Folder]: ...

    async def list_messages(
        self,
        *,
        folder_id: str | None,
        since: datetime | None,
        until: datetime | None,
        page_size: int,
        page: str | None,
        skip: int = 0,
    ) -> tuple[list[MessageSummary], str | None]:
        """One page of messages, newest first. ``page`` is an opaque continuation from a previous
        call; without one, ``skip`` leaves out that many of the newest messages first."""
        ...

    async def search(
        self, *, query: str, folder_id: str | None, page_size: int, page: str | None
    ) -> tuple[list[MessageSummary], str | None]: ...

    async def conversation(self, conversation_id: str) -> tuple[list[MessageSummary], bool]:
        """Every message of the conversation in every folder (unsorted), and whether it was truncated."""
        ...

    async def conversation_folders(
        self, conversation_ids: list[str]
    ) -> dict[str, tuple[list[tuple[str | None, str | None]], bool]]:
        """Per conversation: (folder id, Internet message id) of each message, and whether there are
        more than listed."""
        ...

    async def count_messages(
        self, *, folder_ids: list[str], since: datetime | None, until: datetime | None
    ) -> dict[str, FolderCount]:
        """Count messages and find the newest date in the window per folder, excluding subfolders.
        A failed folder is left out, so one failure never hides the others' results."""
        ...

    async def get_message(self, message_id: str, *, body_format: BodyFormat = "text") -> Message: ...

    async def get_messages(
        self, message_ids: list[str], *, body_format: BodyFormat = "text"
    ) -> FetchedMessages:
        """Batch fetch with bodies. Per-item failures are reported, not raised."""
        ...

    async def list_attachments(self, message_id: str) -> list[Attachment]: ...

    async def list_attachments_many(
        self, message_ids: list[str]
    ) -> tuple[dict[str, list[Attachment]], dict[str, Failure]]:
        """Attachments per message, and why listing failed for the others (per item, not raised)."""
        ...

    async def download_attachment(self, message_id: str, attachment_id: str, dest: Path) -> int: ...

    async def download_mime(self, message_id: str, dest: Path) -> int: ...

    async def profile(self) -> UserProfile:
        """The signed-in user's display name and address."""
        ...

    async def profile_photo(self) -> bytes | None:
        """The signed-in user's photo (small), or None when none is set."""
        ...

    async def get_summaries(self, message_ids: list[str]) -> FetchedSummaries:
        """Batch lookup of message summaries (folder, read state, flag)."""
        ...


@dataclass(frozen=True)
class FolderTarget:
    """A move target: a well-known folder by its alias, or any folder by its Graph id."""

    folder_id: str
    well_known: str | None = None


@dataclass
class FetchedSummaries:
    """A batch lookup: every requested id is in exactly one of the two maps."""

    summaries: dict[str, MessageSummary | None] = field(default_factory=dict)  # None: not on the server
    failed: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class SignatureSettings:
    """Fresh native signature settings; scope is opaque and passed back unchanged."""

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


class SignatureStore(Protocol):
    """Account-scoped native signature settings; each write is sent once."""

    def account(self) -> dict[str, Any]: ...
    async def settings(self) -> SignatureSettings: ...
    async def contents(self, name: str, settings: SignatureSettings) -> SignatureContents | None: ...
    async def create(self, name: str, html: str, text: str, settings: SignatureSettings) -> None: ...
    async def update(self, name: str, html: str, text: str, settings: SignatureSettings) -> None: ...
    async def delete(self, name: str) -> None: ...
    async def set_default(self, name: str | None, which: str, settings: SignatureSettings) -> None: ...


class MailWriter(Protocol):
    """Writes as the signed-in account. Each call is sent once and never retried."""

    def account(self) -> dict[str, Any]:
        """Identity claims (tid, oid, upn) of the credential the writes use."""
        ...

    async def create_draft(self, message: DraftMessage) -> str | None:
        """Save into Drafts without sending; the draft's id when the backend reports it."""
        ...

    async def send_draft(self, draft_id: str, revision: str) -> None:
        """Send an existing draft without changing its content, once. ``revision`` is the version
        it was read at: if the draft changed since, nothing is sent."""
        ...

    # Mutations: {message id: None when done, else the backend's response code}.

    async def set_read(self, message_ids: list[str], is_read: bool) -> dict[str, str | None]: ...

    async def set_flag(self, message_ids: list[str], flagged: bool) -> dict[str, str | None]: ...

    async def move(self, message_ids: list[str], folder: FolderTarget) -> dict[str, str | None]: ...

    async def delete(self, message_ids: list[str]) -> dict[str, str | None]:
        """Move to Deleted Items. There is no hard delete."""
        ...


class RuleWriter(Protocol):
    """OWS rule reads and single-attempt writes; wire formats stay in the adapter."""

    async def list_rules(self) -> list[InboxRule]: ...
    async def create_rule(self, changes: RuleChange, folder: Folder | None) -> str | None: ...
    async def update_rule(self, rule: InboxRule, changes: RuleChange, folder: Folder | None) -> None: ...
    async def reorder_rules(self, rules: list[InboxRule]) -> None: ...
    async def delete_rule(self, rule: InboxRule) -> None: ...

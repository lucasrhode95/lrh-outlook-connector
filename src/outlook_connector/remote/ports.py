"""The ports the service depends on (architecture §6.1).

Adapters: remote/graph_mail.py implements MailReader, remote/ows.py implements MailWriter.
Swapping an adapter (another tenant, a policy change) must not change anything above this file.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Literal, Protocol

from outlook_connector.domain.models import (
    Attachment,
    EmailProposal,
    Folder,
    Message,
    MessageSummary,
)

BodyFormat = Literal["text", "html"]


@dataclass
class FetchedMessages:
    """A batch fetch: every requested id is in exactly one of the two maps."""

    messages: dict[str, Message | None] = field(default_factory=dict)  # None: no longer on the server
    failed: dict[str, str] = field(default_factory=dict)  # not fetched (throttled, error): the reason


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
    ) -> tuple[list[MessageSummary], str | None]:
        """One page of messages, newest first. ``page`` is an opaque continuation from a previous call."""
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
    ) -> dict[str, int] | None:
        """Server count of messages in the window per folder (not its subfolders); None if any
        count is unavailable."""
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
    ) -> tuple[dict[str, list[Attachment]], dict[str, str]]:
        """Attachments per message, and the reason for each message whose listing failed."""
        ...

    async def attachment_content_ids(
        self, attachments: Mapping[str, list[str]]
    ) -> dict[str, dict[str, str | None]]: ...

    async def download_attachment(self, message_id: str, attachment_id: str, dest: Path) -> int: ...

    async def download_mime(self, message_id: str, dest: Path) -> int: ...

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


class MailWriter(Protocol):
    """Writes as the signed-in account. Each call is sent once and never retried."""

    def account(self) -> dict[str, Any]:
        """Identity claims (tid, oid, upn) of the credential the writes use."""
        ...

    async def create_draft(self, message: EmailProposal) -> str | None:
        """Save into Drafts without sending; the draft's id when the backend reports it."""
        ...

    async def send(self, message: EmailProposal) -> None:
        """Send and keep a copy in Sent Items. Raises WriteOutcomeUnknown when it cannot tell."""
        ...

    # Mutations: {message id: None when done, else the backend's response code}.

    async def set_read(self, message_ids: list[str], is_read: bool) -> dict[str, str | None]: ...

    async def set_flag(self, message_ids: list[str], flagged: bool) -> dict[str, str | None]: ...

    async def move(self, message_ids: list[str], folder: FolderTarget) -> dict[str, str | None]: ...

    async def delete(self, message_ids: list[str]) -> dict[str, str | None]:
        """Move to Deleted Items. There is no hard delete."""
        ...

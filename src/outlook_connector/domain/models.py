"""Domain models: the single schema for the service, the MCP tools and the web API.

These never mirror a Microsoft wire format. remote/graph_mapping.py converts Graph JSON into
them, and nothing else knows Graph field names.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, Literal

from pydantic import BaseModel, Field, SerializerFunctionWrapHandler, model_serializer

BodyKind = Literal["unique", "full", "html"]
CombineMode = Literal["per_thread", "all", "none"]
ExportFormat = Literal["txt", "jsonl"]
Detail = Literal["compact", "full"]

# Why messages were left out of a result (Coverage.excluded, ExportArtifact.messages_excluded keys).
ExclusionReason = Literal["deleted_or_junk", "sync_issues", "outgoing", "hidden"]
EXCLUSION_TEXT: dict[str, str] = {
    "deleted_or_junk": "in Deleted Items or Junk Email (include_deleted_items=false)",
    "sync_issues": "in Sync Issues, Outlook's conflict and failure copies (include_deleted_items=false)",
    "outgoing": "in Sent Items, Drafts or Outbox (received_only=true)",
    "hidden": "in hidden folders or outside the mail folders (out of reach)",
}


class Compact(BaseModel):
    """Serializes optional fields only when set (not null, not an empty list): MCP results stay small,
    and a missing field means its default. Required fields are always present."""

    @model_serializer(mode="wrap")
    def _drop_empty(self, handler: SerializerFunctionWrapHandler) -> Any:
        data = handler(self)
        if not isinstance(data, dict):
            return data
        fields = type(self).model_fields
        return {
            k: v
            for k, v in data.items()
            if (k in fields and fields[k].is_required()) or (v is not None and v != [] and v != {})
        }


class Recipient(Compact):
    name: str | None = None
    address: str | None = None

    def display(self) -> str:
        if self.name and self.address and self.name != self.address:
            return f"{self.name} <{self.address}>"
        return self.address or self.name or "(unknown)"


class Folder(Compact):
    id: str
    parent_id: str | None = None
    name: str
    well_known: str | None = None  # inbox, sentitems, archive, ... when the folder is a well-known one
    path: str = ""  # "Inbox/Projects/RIE" style, filled by the service
    total: int | None = None
    unread: int | None = None
    child_count: int | None = None
    hidden: bool = False  # Graph's isHidden. Hidden folders are out of reach and never listed, except
    # Sync Issues, which Outlook hides from its mail view but which stays reachable like Deleted Items.


class MessageSummary(Compact):
    id: str  # Graph immutable id
    conversation_id: str | None = None
    folder_id: str | None = None
    folder: str | None = None  # folder path, filled by the service
    subject: str | None = None
    sender: Recipient | None = None
    to: list[Recipient] = Field(default_factory=list)
    cc: list[Recipient] = Field(default_factory=list)
    received_at: datetime | None = None
    sent_at: datetime | None = None
    is_read: bool | None = None
    is_draft: bool | None = None
    has_attachments: bool = False
    importance: str | None = None
    categories: list[str] = Field(default_factory=list)
    flagged: bool = False
    preview: str | None = None
    internet_message_id: str | None = None
    is_deleted: bool = False  # retained locally after the server copy disappeared
    deleted_at: datetime | None = None
    also_in: list[str] = Field(default_factory=list)  # folders holding another copy (same Internet id)


# Filled by the service per result, never stored with the message.
DERIVED_FIELDS = frozenset({"folder", "is_deleted", "deleted_at", "also_in"})


class Attachment(Compact):
    id: str
    message_id: str
    name: str | None = None
    content_type: str | None = None
    size: int | None = None
    is_inline: bool = False
    kind: Literal["file", "item", "reference", "unknown"] = "unknown"
    content_id: str | None = None


class Message(MessageSummary):
    bcc: list[Recipient] = Field(default_factory=list)
    body_text: str | None = None  # full body as plain text
    unique_body_text: str | None = None  # body without quoted history, plain text
    body_html: str | None = None
    unique_body_html: str | None = None
    attachments: list[Attachment] = Field(default_factory=list)

    def body(self, kind: BodyKind) -> str:
        if kind == "html":
            return self.body_html or ""
        if kind == "unique":
            return self.unique_body_text if self.unique_body_text is not None else (self.body_text or "")
        return self.body_text or ""


class Coverage(Compact):
    """What a result actually covers. Agents must read this before treating results as complete.

    ``complete``: nothing more is retrievable for this request (otherwise follow ``cursor``, or read
    ``notes``). ``excluded``: messages left out by folder, per reason.
    """

    source: Literal["remote", "local", "remote+local"]
    complete: bool
    server_total: int | None = None
    excluded: dict[str, int] = Field(default_factory=dict)
    notes: list[str] = Field(default_factory=list)


class MessagePage(Compact):
    items: list[MessageSummary]
    cursor: str | None = None
    coverage: Coverage


class ConversationHit(Compact):
    conversation_id: str | None
    subject: str | None
    last_received_at: datetime | None
    matching_messages: list[MessageSummary]
    message_count: int | None = None  # the whole conversation, counted like get_thread
    message_count_at_least: bool = False


class SearchResult(Compact):
    query: str
    conversations: list[ConversationHit]
    cursor: str | None = None
    coverage: Coverage


class MessageContent(Compact):
    message: MessageSummary
    body_kind: BodyKind
    text: str
    offset: int
    total_chars: int
    next_offset: int | None
    attachments: list[Attachment] = Field(default_factory=list)


class ThreadMessage(Compact):
    message: MessageSummary
    text: str | None = None  # bounded body when requested
    truncated: bool = False


class Thread(Compact):
    conversation_id: str
    subject: str | None
    messages: list[ThreadMessage]
    cursor: str | None = None
    coverage: Coverage


class ThreadSize(Compact):
    conversation_id: str
    messages: int  # what get_thread would list with the same include_deleted_items
    at_least: bool = False  # the conversation is larger than the server listed in one request


EXPORT_MAX_MESSAGES = 2000  # hard cap per export


class ExportRequest(BaseModel):
    """What to export: conversations, messages, and/or every message in a range. All are combined."""

    conversation_ids: list[str] = Field(default_factory=list)
    message_ids: list[str] = Field(default_factory=list)
    # range selection: active when any of these is set
    since: datetime | None = None
    until: datetime | None = None
    folder: str | None = None  # path, alias or id; None = whole mailbox
    received_only: bool = False  # leave out Sent Items, Drafts and Outbox (range selection only)
    limit: int = Field(default=EXPORT_MAX_MESSAGES, ge=1, le=EXPORT_MAX_MESSAGES)
    format: ExportFormat = "txt"  # jsonl: one JSON record per message, for agents
    include_attachments: bool = False
    combine: CombineMode = "per_thread"
    body: Literal["unique", "full"] = "unique"
    include_deleted_items: bool = False

    @property
    def by_range(self) -> bool:
        return bool(self.since or self.until or self.folder or self.received_only)


class ExportArtifact(Compact):
    path: str
    filename: str
    content_type: str
    size: int
    message_count: int
    text_files: int
    attachment_files: int
    attachments_unavailable: int  # attachment files that could not be downloaded
    attachment_listing_failures: int = 0  # messages whose attachments could not be listed
    messages_excluded: dict[str, int] = Field(default_factory=dict)  # ExclusionReason -> count
    duplicates_merged: int = 0  # copies of the same message (same Internet id) exported once
    messages_unavailable: int = 0  # selected, but the body could not be fetched (marked in the file)
    unavailable_message_ids: list[str] = Field(default_factory=list)

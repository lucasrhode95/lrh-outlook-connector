"""Domain models: the single schema for the service, the MCP tools and the web API.

These never mirror a Microsoft wire format. remote/graph_mapping.py converts Graph JSON into
them, and nothing else knows Graph field names.
"""

from __future__ import annotations

from datetime import datetime
from typing import Literal

from pydantic import BaseModel, Field

BodyKind = Literal["unique", "full", "html"]
CombineMode = Literal["per_thread", "all", "none"]


class Recipient(BaseModel):
    name: str | None = None
    address: str | None = None

    def display(self) -> str:
        if self.name and self.address and self.name != self.address:
            return f"{self.name} <{self.address}>"
        return self.address or self.name or "(unknown)"


class Folder(BaseModel):
    id: str
    parent_id: str | None = None
    name: str
    well_known: str | None = None  # inbox, sentitems, archive, ... when the folder is a well-known one
    path: str = ""  # "Inbox/Projects/RIE" style, filled by the service
    total: int | None = None
    unread: int | None = None
    child_count: int | None = None
    hidden: bool = False


class MessageSummary(BaseModel):
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


class Attachment(BaseModel):
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


class Coverage(BaseModel):
    """What a result actually covers. Agents must read this before treating results as complete."""

    source: Literal["remote", "local", "remote+local"]
    complete: bool
    more_available: bool = False
    server_total: int | None = None
    notes: list[str] = Field(default_factory=list)


class MessagePage(BaseModel):
    items: list[MessageSummary]
    cursor: str | None = None
    coverage: Coverage


class ConversationHit(BaseModel):
    conversation_id: str | None
    subject: str | None
    last_received_at: datetime | None
    matching_messages: list[MessageSummary]


class SearchResult(BaseModel):
    query: str
    conversations: list[ConversationHit]
    cursor: str | None = None
    coverage: Coverage


class MessageContent(BaseModel):
    message: MessageSummary
    body_kind: BodyKind
    text: str
    offset: int
    total_chars: int
    next_offset: int | None
    attachments: list[Attachment] = Field(default_factory=list)


class ThreadMessage(BaseModel):
    message: MessageSummary
    text: str | None = None  # bounded body when requested
    truncated: bool = False


class Thread(BaseModel):
    conversation_id: str
    subject: str | None
    messages: list[ThreadMessage]
    cursor: str | None = None
    coverage: Coverage


class ExportRequest(BaseModel):
    conversation_ids: list[str] = Field(default_factory=list)
    message_ids: list[str] = Field(default_factory=list)
    include_attachments: bool = False
    combine: CombineMode = "per_thread"
    body: Literal["unique", "full"] = "unique"
    include_deleted_items: bool = False


class ExportArtifact(BaseModel):
    path: str
    filename: str
    content_type: str
    size: int
    message_count: int
    text_files: int
    attachment_files: int
    attachments_unavailable: int

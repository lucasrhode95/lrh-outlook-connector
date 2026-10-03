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
    also_in: list[str] = Field(default_factory=list)  # folders holding another copy (same Internet id)


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


ExportStep = Literal["fetching message bodies", "listing attachments", "downloading an attachment"]


class ExportError(Compact):
    """Why part of an export (or a thread body) is missing: the step, Microsoft's answer, the
    likely cause, whether retrying can help, and what to do."""

    step: ExportStep
    status: int | None = None  # HTTP status; None: no response (or no HTTP failure at all)
    code: str | None = None  # Microsoft's error code
    message: str | None = None  # Microsoft's message, shortened to one line
    request_id: str | None = None
    likely_cause: str
    retry: bool  # retrying later can succeed
    fix: str


class ThreadMessage(Compact):
    message: MessageSummary
    text: str | None = None  # bounded body when requested
    truncated: bool = False
    export_error: ExportError | None = None  # the body could not be fetched (``text`` holds its block)


class Thread(Compact):
    conversation_id: str
    subject: str | None
    messages: list[ThreadMessage]
    cursor: str | None = None
    coverage: Coverage
    body_errors: int = 0  # messages on this page whose body could not be fetched (see export_error)


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
    messages_excluded: dict[str, int] = Field(default_factory=dict)  # ExclusionReason -> count
    unavailable_message_ids: list[str] = Field(default_factory=list)  # bodies not fetched (marked in file)
    export_errors: dict[str, int] = Field(default_factory=dict)  # ExportStep -> failures, marked in file
    error_summary: str | None = None  # the file header's "Export errors: ..." line


# ---------------------------------------------------------------- writes (requirements v4 §11)

MAX_RECIPIENTS = 100
MAX_BODY_CHARS = 100_000
ADDRESS_PATTERN = r"^[^@\s<>,;\"]+@[^@\s<>,;\"]+\.[^@\s<>,;\"]+$"


class OutgoingMessage(BaseModel):
    """A plain-text message to draft or send from the signed-in account (no attachments, no Send As).

    As a reply (``reply_to_message_id``), Outlook appends the quoted original below ``body``, and
    omitted recipients and subject default to Outlook's: the sender (and, with ``reply_all``, the
    other recipients, minus you unless nobody else is left) and "RE: <subject>".
    """

    to: list[str] = Field(default_factory=list, description="Email addresses.")
    cc: list[str] = Field(default_factory=list, description="Email addresses.")
    bcc: list[str] = Field(default_factory=list, description="Email addresses.")
    subject: str | None = Field(default=None, max_length=255)
    body: str = Field(max_length=MAX_BODY_CHARS, description="Plain text.")
    reply_to_message_id: str | None = Field(default=None, description="Reply to this message.")
    reply_all: bool = False


class EmailProposal(Compact):
    """Exactly what would be sent, resolved and validated, with the code that confirms it."""

    sender: str  # the signed-in account; there is no Send As
    to: list[str] = Field(default_factory=list)
    cc: list[str] = Field(default_factory=list)
    bcc: list[str] = Field(default_factory=list)
    subject: str
    body: str
    reply_to_message_id: str | None = None
    reply_all: bool = False
    quotes_original: bool = False  # a reply: Outlook appends the quoted original below the body
    confirmation: str  # changes with any field above


class DraftResult(Compact):
    id: str  # the draft's Graph immutable id (readable with get_message)
    folder: str = "Drafts"
    proposal: EmailProposal  # what the draft holds (its confirmation code is not needed to save)
    verified: bool  # the draft was read back from the mailbox


SendStatus = Literal["sent", "unknown"]


class SendResult(Compact):
    status: SendStatus  # "unknown": no clear answer and no copy found in Sent Items yet
    detail: str
    sent_item_id: str | None = None  # the Sent Items copy, when it was looked for and found


MAX_MUTATION_ITEMS = 100
MutationStatus = Literal["done", "unchanged", "not_found", "failed", "unknown"]


class ItemResult(Compact):
    id: str
    status: MutationStatus  # unchanged: already so, nothing sent; unknown: no clear answer, not confirmed
    detail: str | None = None


class MutationResult(Compact):
    """One result per message, in request order. Partial failure is reported, never hidden."""

    action: str
    results: list[ItemResult]
    counts: dict[str, int] = Field(default_factory=dict)  # MutationStatus -> messages

"""MCP server over stdio: thin tools over the shared service (architecture §8).

One process per agent session. Nothing touches the network until the first tool call.
Reads never change the mailbox (not even read state). The write tools save drafts, send mail
(only with the user's confirmation of the exact message, requirements v4 §11.1) and change
messages named by id (§11.2), each with a result per message.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Annotated, Literal

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from outlook_connector.bootstrap import AppContext, Services
from outlook_connector.domain.models import (
    EXPORT_MAX_MESSAGES,
    Attachment,
    BodyKind,
    CombineMode,
    Conversation,
    Detail,
    DraftResult,
    EmailProposal,
    ExportArtifact,
    ExportFormat,
    ExportRequest,
    Folder,
    MessageContent,
    MessagePage,
    MutationResult,
    OutgoingMessage,
    SearchResult,
    SendResult,
)
from outlook_connector.service.files import SavedFile

INSTRUCTIONS = """\
The signed-in user's own Outlook mailbox: reads through Microsoft Graph, drafts and sending through \
Outlook Web. Reading never changes the mailbox, not even read state.

- Finding mail: search_messages for topics (server-side search, accent-insensitive, supports \
subject:/from:/to: terms) and list_messages for recent mail or a date window (folder optional; \
omit it for the whole mailbox). Both return message ids and conversation ids; search hits also \
carry the conversation's message_count. Results are compact by default (detail="full" adds \
recipients, categories and Internet ids).
- Scope, the same for every tool: Deleted Items and Junk Email are left out unless \
include_deleted_items=true (a folder you name is always included; subfolders count with their \
parent). include_sent_items=false also leaves out Sent \
Items, Drafts and Outbox: use it for "the latest mail I received", which includes mail that rules \
filed into other folders. Both flags point the same way: true shows more mail, false filters more. \
coverage.excluded counts what was left out.
- Out of reach: hidden folders, and items outside the mail folders (Teams meeting records, settings \
and other non-mail items), are never listed, searched, counted, grouped into conversations or exported, and \
list_folders does not show them; coverage.excluded.hidden counts any that were dropped. Search \
covers mail only.
- Meeting mail (invitations and their updates, cancellations, replies to invitations) carries \
meeting: kind (invite, update, cancelled, accepted, tentative, declined), start, end, location, \
out_of_date; ordinary mail has none. A reply written to an invitation stays in its conversation. \
include_meeting_mail=false (list, search, range export) leaves meeting mail out: a conversation \
that is only invitations, RSVPs and cancellations disappears; one with real replies shows \
through them (get_conversation still returns the whole conversation).
- Copies of one message (mail sent to yourself or to a list you are on) are shown once; also_in \
names the folders of the other copies.
- Reading: get_conversation returns a whole conversation across folders, oldest first, with bodies \
without quoted history by default; a body that could not be fetched sets export_error on its \
message, and body_errors counts them. get_message reads one message with offset/max_chars continuation.
- Always read `coverage` before treating results as complete; follow `cursor` for more. \
list_messages(include_total=true) gives the server's count for the window, to plan large reads. \
When Junk Email and Deleted Items hold most of the mailbox, a mailbox-wide list_messages reads folder \
by folder (full pages, a bit slower) and coverage.notes says so: pass that note on to the user.
- Mail deleted on the server is gone: nothing is kept locally.
- Attachments: list_attachments, then download_attachment saves the raw file and returns its local \
path for you to read with your own file tools. save_message_mime saves the original .eml.
- export_messages writes one local file and returns its path. Select conversations, message ids \
and/or a range (since/until/folder/include_sent_items) in one call; at most 2,000 messages (`limit` \
lowers that). For a large period, export the range rather than enumerating ids. format="jsonl" \
writes one JSON record per message (ids, dates, folder, people, body): use it to analyse mail; \
"txt" is for people. Read messages_excluded and error_summary in the result: parts that could not \
be exported are marked [EXPORT ERROR] in TXT and carry an export_error object in JSONL (step, \
status, likely_cause, retry, fix).
- Throttling: Microsoft Graph limits each mailbox to about 4 concurrent requests and 10,000 requests \
per 10 minutes (a $batch counts each of its up to 20 items). This connector paces and retries for you. \
Do not call these tools in parallel, and prefer one large call (a range export, a bigger limit) over \
many small ones. On a throttling error, wait at least a minute before retrying.
- Drafts: create_draft saves a plain-text message or reply into Drafts and never sends it (a reply \
reports history_intact: whether Outlook quoted the original exactly as received); the user \
reviews and sends it from Outlook. Prefer it whenever the user has not explicitly asked you to send.
- Sending, only when the user explicitly asks to send: (1) propose_email with the message; (2) show \
the user the whole proposal (from, to, cc, bcc, subject, full body; a reply also carries the quoted \
original, added by Outlook; a reply is sent only if its draft keeps the original exactly as \
received, otherwise nothing is sent) and ask them to confirm it; (3) only after they confirm, \
send_email with the same message and the proposal's confirmation code as user_confirmation. \
Never confirm on the user's behalf. A message changed after confirmation is refused: propose and \
confirm again. If send_email returns status "unknown", do not send again; ask the user to check \
Sent Items and Outbox.
- Changing messages: set_read_state, set_flag, move_messages and delete_messages take \
explicit message ids (from list, search or get_conversation), at most 100 per call, never a query; \
set_read_state also takes conversation ids. Each returns a result per message: done, unchanged \
(already so; nothing sent), not_found, failed (with Outlook's code) or unknown (no clear answer; \
check before repeating). delete_messages moves \
to Deleted Items; messages already there are left alone (there is no permanent delete). Act only \
on messages the user asked about, and say which ones before changing many.
- Writes need the write sign-in (`outlook-connector auth write`).
- If a tool says sign-in is required, ask the user to run the quoted `outlook-connector auth` command \
in a terminal; never attempt to sign in yourself. An "access denied" error is about that item, \
not the sign-in: do not ask the user to sign in again for it.
"""

READ_ONLY = ToolAnnotations(
    readOnlyHint=True, destructiveHint=False, idempotentHint=True, openWorldHint=False
)
LOCAL_FILE = ToolAnnotations(
    readOnlyHint=True, destructiveHint=False, idempotentHint=False, openWorldHint=False
)
DRAFT = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=False)
SEND = ToolAnnotations(readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=True)
CHANGE = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=True, openWorldHint=False)
RELOCATE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, idempotentHint=True, openWorldHint=False)

MessageIds = Annotated[
    list[str], Field(description="Explicit message ids (at most 100), from list, search or get_conversation.")
]


IncludeDeleted = Annotated[
    bool,
    Field(description="Include Deleted Items and Junk Email (a folder you name is always included)."),
]
IncludeMeetings = Annotated[
    bool,
    Field(
        description="Include meeting mail: invitations, RSVPs and cancellations "
        "(false: leave them out; conversations with real replies still show those)."
    ),
]
IncludeSent = Annotated[
    bool, Field(description="Include Sent Items, Drafts and Outbox (false: only mail you received).")
]
DetailLevel = Annotated[
    Detail, Field(description="compact (default) or full: adds recipients, categories and Internet ids.")
]


def _utc(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    return value.replace(tzinfo=UTC) if value.tzinfo is None else value


def build_server(context: AppContext) -> FastMCP:
    mcp = FastMCP("lrh-outlook", instructions=INSTRUCTIONS, log_level="WARNING")

    async def services() -> Services:
        return await context.services()

    @mcp.tool(annotations=READ_ONLY)
    async def auth_status() -> dict[str, object]:
        """Offline sign-in status: the bound account and which client profiles have credentials."""
        status = context.tokens.status()
        return {
            "account": [a.username for a in status.accounts],
            "profiles": {
                p.profile: {"signed_in": p.signed_in, "purpose": p.purpose} for p in status.profiles
            },
            "sign_in_command": context.tokens.sign_in_command("read"),
        }

    @mcp.tool(annotations=READ_ONLY)
    async def list_folders(refresh: bool = False) -> list[Folder]:
        """Mail folders with paths (e.g. Inbox/Projects/RIE), well-known aliases and counts.

        Hidden folders are out of reach and not listed."""
        return await (await services()).mailbox.folders(refresh=refresh)

    @mcp.tool(annotations=READ_ONLY)
    async def list_messages(
        folder: Annotated[
            str | None,
            Field(
                description="Folder path, alias (inbox, sentitems, archive...) "
                "or id. Omit for the whole mailbox."
            ),
        ] = None,
        since: Annotated[datetime | None, Field(description="Inclusive lower bound (ISO 8601).")] = None,
        until: Annotated[datetime | None, Field(description="Inclusive upper bound (ISO 8601).")] = None,
        limit: Annotated[int, Field(ge=1, le=200)] = 25,
        cursor: str | None = None,
        include_sent_items: IncludeSent = True,
        include_meeting_mail: IncludeMeetings = True,
        include_deleted_items: IncludeDeleted = False,
        include_total: Annotated[
            bool, Field(description="Also count the server's messages in scope (first page only).")
        ] = False,
        detail: DetailLevel = "compact",
    ) -> MessagePage:
        """Messages newest first, from the server.

        With folder exclusions a page can hold fewer than limit messages; follow cursor."""
        return await (await services()).mailbox.list_messages(
            folder=folder,
            since=_utc(since),
            until=_utc(until),
            limit=limit,
            cursor=cursor,
            include_sent_items=include_sent_items,
            include_meeting_mail=include_meeting_mail,
            include_deleted_items=include_deleted_items,
            include_total=include_total,
            detail=detail,
        )

    @mcp.tool(annotations=READ_ONLY)
    async def search_messages(
        query: Annotated[
            str,
            Field(description="Words or KQL terms, e.g. 'budget review', 'from:alice', 'subject:invoice'."),
        ],
        since: Annotated[datetime | None, Field(description="Inclusive lower bound (ISO 8601).")] = None,
        until: Annotated[datetime | None, Field(description="Inclusive upper bound (ISO 8601).")] = None,
        folder: Annotated[str | None, Field(description="Folder path, alias or id.")] = None,
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        cursor: str | None = None,
        include_sent_items: IncludeSent = True,
        include_meeting_mail: IncludeMeetings = True,
        include_deleted_items: IncludeDeleted = False,
        detail: DetailLevel = "compact",
    ) -> SearchResult:
        """Server-side search of mail only (hidden folders and non-mail items are left out); hits
        grouped by conversation in rank order, with each conversation's message_count."""
        return await (await services()).mailbox.search(
            query,
            since=_utc(since),
            until=_utc(until),
            folder=folder,
            limit=limit,
            cursor=cursor,
            include_sent_items=include_sent_items,
            include_meeting_mail=include_meeting_mail,
            include_deleted_items=include_deleted_items,
            detail=detail,
        )

    @mcp.tool(annotations=READ_ONLY)
    async def get_conversation(
        conversation_id: str,
        include_bodies: bool = True,
        body: Literal["unique", "full"] = "unique",
        include_deleted_items: IncludeDeleted = False,
        max_chars: Annotated[int, Field(ge=1000, le=400_000)] = 40_000,
        cursor: Annotated[
            str | None,
            Field(description="Continuation; it restores the original options, which are then ignored."),
        ] = None,
    ) -> Conversation:
        """A whole conversation across folders, oldest first. body=unique strips quoted reply history."""
        return await (await services()).conversations.get_conversation(
            conversation_id,
            include_bodies=include_bodies,
            body=body,
            include_deleted_items=include_deleted_items,
            max_chars=max_chars,
            cursor=cursor,
        )

    @mcp.tool(annotations=READ_ONLY)
    async def get_message(
        message_id: str,
        body: BodyKind = "unique",
        offset: Annotated[int, Field(ge=0)] = 0,
        max_chars: Annotated[int, Field(ge=1000, le=200_000)] = 20_000,
    ) -> MessageContent:
        """One message's body (unique/full/html) with offset continuation, plus attachment metadata."""
        return await (await services()).mailbox.get_message(
            message_id, body=body, offset=offset, max_chars=max_chars
        )

    @mcp.tool(annotations=READ_ONLY)
    async def list_attachments(message_id: str) -> list[Attachment]:
        """Attachment metadata. kind=item is a forwarded email; kind=reference is a cloud link."""
        return await (await services()).mailbox.attachments(message_id)

    @mcp.tool(annotations=LOCAL_FILE)
    async def download_attachment(message_id: str, attachment_id: str) -> SavedFile:
        """Save one attachment as a local file and return its path (forwarded emails as .eml)."""
        return await (await services()).files.download_attachment(message_id, attachment_id)

    @mcp.tool(annotations=LOCAL_FILE)
    async def save_message_mime(message_id: str) -> SavedFile:
        """Save the original message (MIME) as a local .eml file and return its path."""
        return await (await services()).files.save_mime(message_id)

    @mcp.tool(annotations=LOCAL_FILE)
    async def export_messages(
        conversation_ids: list[str] | None = None,
        message_ids: list[str] | None = None,
        include_attachments: bool = False,
        combine: Annotated[
            CombineMode,
            Field(
                description="per_conversation: one TXT per conversation; all: one TXT; "
                "none: one TXT per message."
            ),
        ] = "per_conversation",
        body: Literal["unique", "full"] = "unique",
        include_deleted_items: IncludeDeleted = False,
        format: Annotated[
            ExportFormat,
            Field(description="txt (for people) or jsonl (one JSON record per message, for analysis)."),
        ] = "txt",
        since: Annotated[
            datetime | None, Field(description="Range selection: inclusive lower bound (ISO 8601).")
        ] = None,
        until: Annotated[
            datetime | None, Field(description="Range selection: inclusive upper bound (ISO 8601).")
        ] = None,
        folder: Annotated[
            str | None,
            Field(description="Range selection: folder path, alias or id (default: whole mailbox)."),
        ] = None,
        include_sent_items: IncludeSent = True,
        include_meeting_mail: IncludeMeetings = True,
        limit: Annotated[
            int,
            Field(
                ge=1,
                le=EXPORT_MAX_MESSAGES,
                description="Refuse the export if the selection holds more messages than this.",
            ),
        ] = EXPORT_MAX_MESSAGES,
    ) -> ExportArtifact:
        """Export conversations, messages and/or a date range to one local file; returns its path.

        A .txt or .jsonl, or a .zip when there are several files or attachments. Copies of one message
        are exported once. The result counts what was left out and lists messages without a body.
        """
        request = ExportRequest(
            conversation_ids=conversation_ids or [],
            message_ids=message_ids or [],
            since=_utc(since),
            until=_utc(until),
            folder=folder,
            include_sent_items=include_sent_items,
            include_meeting_mail=include_meeting_mail,
            limit=limit,
            format=format,
            include_attachments=include_attachments,
            combine=combine,
            body=body,
            include_deleted_items=include_deleted_items,
        )
        return await (await services()).exports.export(request)

    @mcp.tool(annotations=DRAFT)
    async def create_draft(message: OutgoingMessage) -> DraftResult:
        """Save a plain-text message, or a reply (reply_to_message_id), into Drafts. Never sends.

        The user can review, edit and send it from Outlook. Returns the draft's id and what it holds."""
        return await (await services()).writes.create_draft(message)

    @mcp.tool(annotations=READ_ONLY)
    async def propose_email(message: OutgoingMessage) -> EmailProposal:
        """Step 1 of sending: the message exactly as it would be sent, and its confirmation code.

        Changes nothing. Show the whole proposal to the user and ask them to confirm it."""
        return await (await services()).writes.propose(message)

    @mcp.tool(annotations=SEND)
    async def send_email(
        message: OutgoingMessage,
        user_confirmation: Annotated[
            str,
            Field(description="The confirmation code of the proposal the user confirmed (SEND-…)."),
        ],
    ) -> SendResult:
        """Step 2 of sending: send the message the user confirmed, once, from their account.

        Only with the user's explicit confirmation of this exact message (propose_email). Refused if
        the message differs from the confirmed proposal. Never retried: on status "unknown", ask the
        user to check Sent Items and Outbox instead of sending again."""
        return await (await services()).writes.send(message, user_confirmation)

    @mcp.tool(annotations=CHANGE)
    async def set_read_state(
        read: bool,
        message_ids: MessageIds | None = None,
        conversation_ids: Annotated[
            list[str] | None, Field(description="Also every message of these conversations, in scope.")
        ] = None,
        include_deleted_items: IncludeDeleted = False,
    ) -> MutationResult:
        """Mark messages read or unread (read receipts are never sent). A result per message."""
        return await (await services()).mutations.set_read(
            message_ids or [],
            read,
            conversation_ids=conversation_ids,
            include_deleted_items=include_deleted_items,
        )

    @mcp.tool(annotations=CHANGE)
    async def set_flag(message_ids: MessageIds, flagged: bool) -> MutationResult:
        """Flag or unflag messages. A result per message."""
        return await (await services()).mutations.set_flag(message_ids, flagged)

    @mcp.tool(annotations=RELOCATE)
    async def move_messages(
        message_ids: MessageIds,
        folder: Annotated[str, Field(description="Target folder: path, alias (archive, inbox) or id.")],
    ) -> MutationResult:
        """Move messages to a folder (not Deleted Items: use delete_messages). A result per message."""
        return await (await services()).mutations.move(message_ids, folder)

    @mcp.tool(annotations=RELOCATE)
    async def delete_messages(message_ids: MessageIds) -> MutationResult:
        """Move messages to Deleted Items. Messages already there are left alone; nothing is ever
        deleted permanently. A result per message."""
        return await (await services()).mutations.delete(message_ids)

    return mcp


def serve_mcp(*, unsecure: bool) -> None:
    """Run the MCP server over stdio until the client closes it."""
    build_server(AppContext(unsecure=unsecure)).run("stdio")

"""MCP server over stdio: thin tools over the shared service (architecture §8).

One process per agent session. Nothing touches the network until the first tool call.
Reads never change the mailbox (not even read state). The write tools save drafts, send mail
(only with the user's confirmation of the exact message, requirements v4 §11.1) and change
messages named by id (§11.2), each with a result per message.
"""

from __future__ import annotations

from datetime import UTC, datetime
from inspect import Parameter
from inspect import Signature as FunctionSignature
from typing import Annotated, Any, Literal

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from outlook_connector.bootstrap import AppContext, Services
from outlook_connector.domain.models import (
    DEFAULT_SCOPE,
    EXPORT_MAX_MESSAGES,
    Attachment,
    BodyKind,
    CombineMode,
    Conversation,
    Detail,
    DraftResult,
    ExportArtifact,
    ExportFormat,
    ExportRequest,
    Folder,
    InboxRule,
    MessageContent,
    MessagePage,
    MutationResult,
    OutgoingMessage,
    RuleChange,
    RuleWriteResult,
    Scope,
    SearchResult,
    SendResult,
    SignatureDetails,
    SignatureList,
    SignatureWriteResult,
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
- Scope is one object with independent keys: `scope.sent_items` (default true) includes Sent Items, \
Drafts and Outbox; `scope.meeting_mail` (default true) includes invitations, RSVPs and cancellations; \
`scope.deleted_items` (default false) includes Deleted Items and Junk Email. A folder you name is \
always included; subfolders count with their parent. True shows more mail, false filters more. Use \
`scope.sent_items=false` for "the latest mail I received", including mail that rules filed into \
other folders. Export scope filters folder/date windows only; selected conversations \
and message ids stay whole. \
coverage.excluded counts what was left out.
- Out of reach: hidden folders, and items outside the mail folders (Teams meeting records, settings \
and other non-mail items), are never listed, searched, counted, grouped into conversations \
or selected by folder/date-window/conversation export, and \
list_folders does not show them; coverage.excluded.hidden counts any that were dropped. Search \
covers mail only.
- Meeting mail (invitations and their updates, cancellations, replies to invitations) carries \
meeting: kind (invite, update, cancelled, accepted, tentative, declined), start, end, location, \
out_of_date; ordinary mail has none. A reply written to an invitation stays in its conversation. \
scope.meeting_mail=false (list, search, folder/date-window export) leaves meeting mail out: a conversation \
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
- Explicit export message_ids are authoritative like get_message(id); readable hidden/out-of-reach ids \
are included and merged. Folder/date-window selections follow scope; conversation selections stay whole. \
Scope alone does not select anything.
- export_messages writes one local file and returns its path. Select conversations, message ids \
and/or a folder/date window (since/until/folder, narrowed by `scope`), up to \
2,000 messages (`limit` lowers that). Scope alone is not a selection; conversation and message \
selections stay whole. For a large period, prefer \
a folder/date-window export over enumerating ids. format="jsonl" \
writes one JSON record per message (ids, dates, folder, people, body): use it to analyse mail; \
"txt" is for people. Read messages_excluded and error_summary in the result: parts that could not \
be exported are marked [EXPORT ERROR] in TXT and carry an export_error object in JSONL (step, \
status, likely_cause, retry, fix).
- Inline export images: known CID references in the rendered body are included.
  If the content id or rendered body cannot be read, include the image so a read gap
  does not silently drop it.
- Throttling: Microsoft Graph limits each mailbox to about 4 concurrent requests and 10,000 requests \
per 10 minutes (a $batch counts each of its up to 20 items). This connector paces and retries for you. \
Do not call these tools in parallel, and prefer one large folder/date-window export or a bigger limit over \
many small ones. On a throttling error, wait at least a minute before retrying.
- Native signatures: all signature tools, including list_signatures and get_signature, need the
write sign-in. list_signatures reports Outlook's exact, case-sensitive native names, the new-message
and reply/forward defaults, and whether contents are readable. get_signature returns the HTML and
text. create_signature, update_signature, delete_signature and set_default_signature change Outlook
settings once; they are not retried. Names cannot contain commas. A deleted signature can
remain selected as a dangling default, so set a valid default after deletion when needed. To copy an
Outlook draft signature, get its HTML, download its CID-referenced inline images, replace each CID
image with a data:image URI containing those bytes, then save the resulting HTML as a native
signature. create_draft freshly resolves and inserts the default for new mail or replies;
signature selects a named signature and include_signature=false suppresses all signatures.
- Drafts: create_draft accepts exactly one of text_body or html_body, for new mail and replies.
Drafts are composed once, never edited. To change one, create a replacement with the full intended
content (use the same reply_to_message_id for a reply), check its server read-back, then delete the
old draft with delete_messages and use the new id. Never delete first. The old version remains in
Deleted Items; tell the user. Outlook edits and attachments are not carried over. There is no version
check against the old draft; read it first if needed.
- Sending: only after the user explicitly asks, call send_draft with the existing Microsoft draft id.
Never confirm on the user's behalf. It sends the stored draft without changing it, once.
On status "unknown", ask the user to check Sent Items and Outbox before any further send.
- Changing messages: set_read_state, set_flag, move_messages and delete_messages take \
explicit message ids (from list, search or get_conversation), at most 100 per call, never a query; \
set_read_state also takes conversation ids (up to 1,000 listed messages per conversation). \
For large expansions, counts summarize ordinary done/unchanged results; explicit-id and error \
results remain detailed. Read notes for truncation. Each returns results and counts: done, unchanged \
(already so; nothing sent), not_found, failed (with Outlook's code) or unknown (no clear answer; \
check before repeating). continue_on_error defaults to true: later chunks are attempted after errors. \
With false, later messages are failed with detail "not sent". Always report results and counts. \
delete_messages moves \
to Deleted Items; messages already there are left alone (there is no permanent delete). Act only \
on messages the user asked about, and say which ones before changing many.
- Rules: list_rules lists unsupported rules read-only. For create_rule, update_rule, reorder_rules and \
 delete_rule, first omit user_confirmation and show the persistent change and returned RULE code. \
Only after explicit human confirmation repeat the exact request with that code. \
Never confirm on the user's behalf. \
Each write is sent once and read back; unknown means check Outlook before repeating. \
Enable/disable is a separate update. Reordering is refused while any unsupported rule is present.
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
RULE_WRITE = ToolAnnotations(
    readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=False
)
SIGNATURE_WRITE = ToolAnnotations(
    readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=False
)
RELOCATE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, idempotentHint=True, openWorldHint=False)

MessageIds = Annotated[
    list[str], Field(description="Explicit message ids (at most 100), from list, search or get_conversation.")
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
        """Mail folders with paths (e.g. Inbox/Projects/Project Alpha), well-known aliases and counts.

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
        scope: Scope = DEFAULT_SCOPE,
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
            scope=scope,
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
        scope: Scope = DEFAULT_SCOPE,
        detail: DetailLevel = "compact",
    ) -> SearchResult:
        """Server-side search of mail only (hidden folders and non-mail items are left out); hits
        grouped by conversation in rank order, with each conversation's message_count."""
        return await (await services()).mailbox.search(
            query,
            since=since,
            until=until,
            folder=folder,
            limit=limit,
            cursor=cursor,
            scope=scope,
            detail=detail,
        )

    @mcp.tool(annotations=READ_ONLY)
    async def get_conversation(
        conversation_id: str,
        include_bodies: bool = True,
        body: Literal["unique", "full"] = "unique",
        scope: Scope = DEFAULT_SCOPE,
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
            scope=scope,
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
        """List attachment metadata, including inline content ids when available."""
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
        scope: Scope = DEFAULT_SCOPE,
        format: Annotated[
            ExportFormat,
            Field(description="txt (for people) or jsonl (one JSON record per message, for analysis)."),
        ] = "txt",
        since: Annotated[
            datetime | None,
            Field(description="Folder/date-window selection: inclusive lower bound (ISO 8601)."),
        ] = None,
        until: Annotated[
            datetime | None,
            Field(description="Folder/date-window selection: inclusive upper bound (ISO 8601)."),
        ] = None,
        folder: Annotated[
            str | None,
            Field(description="Folder path, alias or id; omit for a whole-mailbox date-window selection."),
        ] = None,
        limit: Annotated[
            int,
            Field(
                ge=1,
                le=EXPORT_MAX_MESSAGES,
                description="Refuse the export if the selection holds more messages than this.",
            ),
        ] = EXPORT_MAX_MESSAGES,
    ) -> ExportArtifact:
        """Export conversations, messages and/or a folder/date window to one local file; returns its path.

        A .txt or .jsonl, or a .zip when there are several files or attachments. Copies of one message
        are exported once. Scope narrows only folder/date-window selections; selected conversations and
        messages stay whole. Scope alone is not a selection. The result counts exclusions and body gaps.
        With attachments, inline images are included when their CID is found in the rendered body;
        if the content id or body cannot be read, the image is included rather than silently dropped.
        """
        request = ExportRequest(
            conversation_ids=conversation_ids or [],
            message_ids=message_ids or [],
            since=_utc(since),
            until=_utc(until),
            folder=folder,
            scope=scope,
            limit=limit,
            format=format,
            include_attachments=include_attachments,
            combine=combine,
            body=body,
        )
        return await (await services()).exports.export(request)

    @mcp.tool(annotations=READ_ONLY)
    async def list_rules() -> list[InboxRule]:
        """Current inbox rules in order. Unsupported rules are read-only; needs the write sign-in."""
        return await (await services()).rules.list_rules()

    @mcp.tool(annotations=RULE_WRITE)
    async def create_rule(changes: RuleChange, user_confirmation: str | None = None) -> RuleWriteResult:
        """Propose a persistent rule; only after human approval repeat with the returned confirmation code."""
        return await (await services()).rules.create_rule(changes, user_confirmation)

    @mcp.tool(annotations=RULE_WRITE)
    async def update_rule(
        rule_id: str, changes: RuleChange, user_confirmation: str | None = None
    ) -> RuleWriteResult:
        """Propose/confirm partial edits. Null conditions clear them; enabled must be a separate update."""
        return await (await services()).rules.update_rule(rule_id, changes, user_confirmation)

    @mcp.tool(annotations=RULE_WRITE)
    async def reorder_rules(rule_ids: list[str], user_confirmation: str | None = None) -> RuleWriteResult:
        """Propose/confirm every rule in the new order. Refused if any rule is unsupported/read-only."""
        return await (await services()).rules.reorder_rules(rule_ids, user_confirmation)

    @mcp.tool(annotations=RULE_WRITE)
    async def delete_rule(rule_id: str, user_confirmation: str | None = None) -> RuleWriteResult:
        """Propose/confirm removing a supported inbox rule; changes future mail handling, never retries."""
        return await (await services()).rules.delete_rule(rule_id, user_confirmation)

    @mcp.tool(annotations=READ_ONLY)
    async def list_signatures() -> SignatureList:
        """List exact native Outlook signature names, defaults and content readability."""
        return await (await services()).signatures.list_signatures()

    @mcp.tool(annotations=READ_ONLY)
    async def get_signature(name: str) -> SignatureDetails:
        """Read the current HTML and text of one native Outlook signature."""
        return await (await services()).signatures.get_signature(name)

    @mcp.tool(annotations=SIGNATURE_WRITE)
    async def create_signature(name: str, html: str) -> SignatureWriteResult:
        """Create one native Outlook signature from passive HTML, including embedded data images."""
        return await (await services()).signatures.create_signature(name, html)

    @mcp.tool(annotations=SIGNATURE_WRITE)
    async def update_signature(name: str, html: str) -> SignatureWriteResult:
        """Replace the contents of an existing native Outlook signature; this does not rename it."""
        return await (await services()).signatures.update_signature(name, html)

    @mcp.tool(annotations=SIGNATURE_WRITE)
    async def delete_signature(name: str) -> SignatureWriteResult:
        """Delete one native Outlook signature; Outlook may leave its default pointer dangling."""
        return await (await services()).signatures.delete_signature(name)

    async def set_default_signature(**arguments: Any) -> SignatureWriteResult:
        """Set the required new, reply or both native defaults; name=null clears the selected default."""
        return await (await services()).signatures.set_default_signature(arguments["name"], arguments["for"])

    # MCP exposes the required field as the exact contract name "for"; Python accepts it through
    # **arguments because "for" is a reserved keyword.
    setattr(  # noqa: B010
        set_default_signature,
        "__signature__",
        FunctionSignature(
            [
                Parameter("name", Parameter.POSITIONAL_ONLY, annotation=str | None),
                Parameter(
                    "for",
                    Parameter.POSITIONAL_ONLY,
                    annotation=Literal["new", "reply", "both"],
                ),
            ],
            return_annotation=SignatureWriteResult,
        ),
    )
    mcp.tool(annotations=SIGNATURE_WRITE)(set_default_signature)

    @mcp.tool(annotations=DRAFT)
    async def create_draft(
        to: list[str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        subject: str | None = None,
        text_body: str | None = None,
        html_body: str | None = None,
        reply_to_message_id: str | None = None,
        reply_all: bool = False,
        signature: str | None = None,
        include_signature: bool = True,
    ) -> DraftResult:
        """Save a new message or reply and return server read-back.

        By default, freshly resolves and inserts the Outlook native default for new mail or replies.
        Set signature to select a native signature by exact name, or include_signature=False to
        suppress all signatures. A missing or unreadable selected signature raises instead of falling back.
        Drafts are composed once. To change one, create a full replacement (using the same
        ``reply_to_message_id`` for a reply), verify its read-back, then delete the old draft with
        ``delete_messages`` and use the new id. Never delete first. Outlook edits and attachments
        are not carried over; the old version remains in Deleted Items. There is no version check
        against the old draft, so read it first if needed.
        """
        return await (await services()).writes.create_draft(
            OutgoingMessage(
                to=to or [],
                cc=cc or [],
                bcc=bcc or [],
                subject=subject,
                text_body=text_body,
                html_body=html_body,
                reply_to_message_id=reply_to_message_id,
                reply_all=reply_all,
                signature=signature,
                include_signature=include_signature,
            )
        )

    @mcp.tool(annotations=SEND)
    async def send_draft(draft_id: str) -> SendResult:
        """Send this existing server draft unchanged, once, only on the user's explicit send request."""
        return await (await services()).writes.send_draft(draft_id)

    @mcp.tool(annotations=CHANGE)
    async def set_read_state(
        read: bool,
        message_ids: MessageIds | None = None,
        conversation_ids: Annotated[
            list[str] | None, Field(description="Also every message of these conversations, in scope.")
        ] = None,
        scope: Scope = DEFAULT_SCOPE,
        continue_on_error: bool = True,
    ) -> MutationResult:
        """Mark messages read or unread (read receipts are never sent). Results and counts; when
        conversations expand past 100 messages, their done/unchanged ones are only counted."""
        return await (await services()).mutations.set_read(
            message_ids or [],
            read,
            conversation_ids=conversation_ids,
            scope=scope,
            continue_on_error=continue_on_error,
        )

    @mcp.tool(annotations=CHANGE)
    async def set_flag(
        message_ids: MessageIds, flagged: bool, continue_on_error: bool = True
    ) -> MutationResult:
        """Flag or unflag messages. A result per message."""
        return await (await services()).mutations.set_flag(
            message_ids, flagged, continue_on_error=continue_on_error
        )

    @mcp.tool(annotations=RELOCATE)
    async def move_messages(
        message_ids: MessageIds,
        folder: Annotated[str, Field(description="Target folder: path, alias (archive, inbox) or id.")],
        continue_on_error: bool = True,
    ) -> MutationResult:
        """Move messages to a folder (not Deleted Items: use delete_messages). A result per message."""
        return await (await services()).mutations.move(
            message_ids, folder, continue_on_error=continue_on_error
        )

    @mcp.tool(annotations=RELOCATE)
    async def delete_messages(message_ids: MessageIds, continue_on_error: bool = True) -> MutationResult:
        """Move messages to Deleted Items. Messages already there are left alone; nothing is ever
        deleted permanently. Also remove an old draft after creating and verifying its replacement.
        A result per message."""
        return await (await services()).mutations.delete(message_ids, continue_on_error=continue_on_error)

    return mcp


def serve_mcp(*, unsecure: bool) -> None:
    """Run the MCP server over stdio until the client closes it."""
    build_server(AppContext(unsecure=unsecure)).run("stdio")

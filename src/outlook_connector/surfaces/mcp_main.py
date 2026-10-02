"""MCP server over stdio: thin tools over the shared service (architecture §8).

One process per agent session. Nothing touches the network until the first tool call.
This MVP surface is read-only: it never changes the mailbox (not even read state).
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Annotated, Literal

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from outlook_connector.bootstrap import AppContext, Services
from outlook_connector.domain.models import (
    Attachment,
    BodyKind,
    CombineMode,
    ExportArtifact,
    ExportRequest,
    Folder,
    MessageContent,
    MessagePage,
    SearchResult,
    Thread,
)
from outlook_connector.service.files import SavedFile

INSTRUCTIONS = """\
Read access to the signed-in user's own Outlook mailbox (Microsoft Graph). These tools never change \
the mailbox, not even read state.

- Finding mail: use search_messages for topics (server-side search, accent-insensitive, supports \
subject:/from:/to: terms) and list_messages for recent mail or a date window (folder optional; \
omit it for the whole mailbox). Both return message ids and conversation ids.
- Reading: get_thread returns a whole conversation across folders, oldest first, with bodies \
without quoted history by default. get_message reads one message with offset/max_chars continuation.
- Always read `coverage` before treating results as complete; follow `cursor` for more.
- Messages with is_deleted=true were deleted on the server and come from local retention.
- Attachments: list_attachments, then download_attachment saves the raw file and returns its local \
path for you to read with your own file tools. save_message_mime saves the original .eml.
- export_messages writes one .txt or .zip export and returns its local path.
- If a tool says sign-in is required, ask the user to run the quoted `outlook-connector auth` command \
in a terminal; never attempt to sign in yourself.
"""

READ_ONLY = ToolAnnotations(
    readOnlyHint=True, destructiveHint=False, idempotentHint=True, openWorldHint=False
)
LOCAL_FILE = ToolAnnotations(
    readOnlyHint=True, destructiveHint=False, idempotentHint=False, openWorldHint=False
)


def _utc(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    return value.replace(tzinfo=UTC) if value.tzinfo is None else value


def build_server(context: AppContext) -> FastMCP:
    mcp = FastMCP("outlook", instructions=INSTRUCTIONS, log_level="WARNING")

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
        """Mail folders with paths (e.g. Inbox/Projects/RIE), well-known aliases and counts."""
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
        refresh: Annotated[bool, Field(description="false = local cache only (no network).")] = True,
        cursor: str | None = None,
    ) -> MessagePage:
        """Messages newest first, from the server, merged with retained messages deleted on the server."""
        return await (await services()).mailbox.list_messages(
            folder=folder, since=_utc(since), until=_utc(until), limit=limit, refresh=refresh, cursor=cursor
        )

    @mcp.tool(annotations=READ_ONLY)
    async def search_messages(
        query: Annotated[str, Field(description="Words or KQL terms, e.g. 'relatório be' or 'from:alice'.")],
        since: datetime | None = None,
        until: datetime | None = None,
        folder: str | None = None,
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        cursor: str | None = None,
    ) -> SearchResult:
        """Server-side mailbox search; hits grouped by conversation in rank order."""
        return await (await services()).mailbox.search(
            query, since=_utc(since), until=_utc(until), folder=folder, limit=limit, cursor=cursor
        )

    @mcp.tool(annotations=READ_ONLY)
    async def get_thread(
        conversation_id: str,
        include_bodies: bool = True,
        body: Literal["unique", "full"] = "unique",
        include_deleted_items: Annotated[
            bool, Field(description="Include messages in Deleted Items/Junk.")
        ] = False,
        max_chars: Annotated[int, Field(ge=1000, le=400_000)] = 40_000,
        cursor: str | None = None,
    ) -> Thread:
        """A whole conversation across folders, oldest first. body=unique strips quoted reply history."""
        return await (await services()).threads.get_thread(
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
        return await (await services()).mailbox.reader.list_attachments(message_id)

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
                description="per_thread: one TXT per conversation; all: one TXT; none: one TXT per message."
            ),
        ] = "per_thread",
        body: Literal["unique", "full"] = "unique",
        include_deleted_items: bool = False,
    ) -> ExportArtifact:
        """Export conversations and/or messages to one local .txt or .zip and return its path."""
        request = ExportRequest(
            conversation_ids=conversation_ids or [],
            message_ids=message_ids or [],
            include_attachments=include_attachments,
            combine=combine,
            body=body,
            include_deleted_items=include_deleted_items,
        )
        return await (await services()).exports.export(request)

    return mcp


def run(*, unsecure: bool) -> None:
    build_server(AppContext(unsecure=unsecure)).run("stdio")

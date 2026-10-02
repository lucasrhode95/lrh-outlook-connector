"""Folders, message listing, message content and search (requirements v4 §7–§9).

Remote first: every call asks Outlook (Graph) and only falls back to retained local data where
the server can no longer provide it (messages deleted on the server).
"""

from __future__ import annotations

import asyncio
import logging
from datetime import datetime

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.domain.models import (
    BodyKind,
    ConversationHit,
    Coverage,
    Folder,
    Message,
    MessageContent,
    MessagePage,
    MessageSummary,
    SearchResult,
)
from outlook_connector.remote.ports import MailReader
from outlook_connector.service import cursors
from outlook_connector.service.reconcile import Reconciler
from outlook_connector.store.db import Store

log = logging.getLogger(__name__)

FOLDER_TTL_SECONDS = 600
NOT_RECEIVED = ("sentitems", "drafts", "outbox", "deleteditems", "junkemail")
RECEIVED_ONLY_PAGES = 10  # server pages scanned at most for one received_only page
LOCAL_ONLY_NOTE = "Local cache only: messages this app has seen before. It is not a mirror of the mailbox."


class Mailbox:
    def __init__(self, reader: MailReader, store: Store) -> None:
        self.reader = reader
        self.store = store
        self.reconciler = Reconciler(reader, store)
        self._folders: dict[str, Folder] | None = None
        self._folder_refresh: asyncio.Task[list[Folder]] | None = None

    # ---------------------------------------------------------------- folders

    async def folders(self, *, refresh: bool = False) -> list[Folder]:
        """Cached folders right away; a stale cache is refreshed in the background (stale-while-revalidate).

        Only an empty cache or ``refresh=True`` waits for the server.
        """
        cached, age = self.store.folders()
        if refresh or not cached:
            cached = await self._refresh_folders()
        elif (age is None or age > FOLDER_TTL_SECONDS) and self._folder_refresh is None:
            self._folder_refresh = asyncio.create_task(self._refresh_folders())
            self._folder_refresh.add_done_callback(self._folder_refresh_done)
        self._folders = {f.id: f for f in cached}
        return sorted(cached, key=lambda f: f.path.lower())

    async def _refresh_folders(self) -> list[Folder]:
        fresh = _with_paths(await self.reader.list_folders())
        self.store.save_folders(fresh)
        self._folders = {f.id: f for f in fresh}
        return fresh

    def _folder_refresh_done(self, task: asyncio.Task[list[Folder]]) -> None:
        self._folder_refresh = None
        if not task.cancelled() and task.exception() is not None:
            log.warning("Background folder refresh failed: %s", type(task.exception()).__name__)

    async def folder_map(self) -> dict[str, Folder]:
        if self._folders is None:
            await self.folders()
        assert self._folders is not None
        return self._folders

    async def _not_received_ids(self) -> set[str]:
        return {f.id for f in (await self.folder_map()).values() if f.well_known in NOT_RECEIVED}

    async def resolve_folder(self, ref: str) -> Folder:
        """A folder id, a well-known alias (inbox, archive, ...), a path (Inbox/Projects) or a name."""
        folders = await self.folder_map()
        needle = ref.strip().strip("/").lower()
        if ref in folders:
            return folders[ref]
        for matcher in (
            lambda f: (f.well_known or "") == needle,
            lambda f: f.path.lower() == needle,
            lambda f: f.name.lower() == needle,
        ):
            matches = [f for f in folders.values() if matcher(f)]
            if len(matches) == 1:
                return matches[0]
            if len(matches) > 1:
                paths = ", ".join(sorted(f.path for f in matches))
                raise InvalidRequest(f"Folder '{ref}' is ambiguous: {paths}. Use the full path.")
        raise InvalidRequest(f"Unknown folder '{ref}'. Use list_folders to see folder paths.")

    async def decorate(self, items: list[MessageSummary]) -> list[MessageSummary]:
        folders = await self.folder_map()
        for item in items:
            folder = folders.get(item.folder_id or "")
            item.folder = folder.path if folder else None
        return items

    # ---------------------------------------------------------------- listing

    async def list_messages(
        self,
        *,
        folder: str | None = None,
        since: datetime | None = None,
        until: datetime | None = None,
        limit: int = 25,
        refresh: bool = True,
        cursor: str | None = None,
        received_only: bool = False,
    ) -> MessagePage:
        """Newest first. ``received_only`` leaves out Sent Items, Drafts, Outbox, Deleted Items and Junk."""
        if not 1 <= limit <= 200:
            raise InvalidRequest("limit must be between 1 and 200.")
        if since and until and since > until:
            raise InvalidRequest("since must not be after until.")
        state = cursors.decode(cursor, "list_messages") if cursor else None
        if state:
            folder_id = state.get("folder_id")
            since = datetime.fromisoformat(state["since"]) if state.get("since") else None
            until = datetime.fromisoformat(state["until"]) if state.get("until") else None
            received_only = bool(state["received_only"])
        else:
            folder_id = (await self.resolve_folder(folder)).id if folder else None
        skip = await self._not_received_ids() if received_only else set()

        if not refresh:
            items = self.store.window(
                folder_id=folder_id, since=since, until=until, limit=None if skip else limit
            )
            items = [m for m in items if m.folder_id not in skip][:limit]
            return MessagePage(
                items=await self.decorate(items),
                coverage=Coverage(source="local", complete=False, notes=[LOCAL_ONLY_NOTE]),
            )

        link = state["link"] if state else None
        items: list[MessageSummary] = []
        for _ in range(RECEIVED_ONLY_PAGES if skip else 1):
            fetched, link = await self.reader.list_messages(
                folder_id=folder_id, since=since, until=until, page_size=limit, page=link
            )
            self.store.upsert_summaries(fetched)
            await self.reconciler.after_window(
                folder_id=folder_id, since=since, until=until, remote=fetched, complete=link is None
            )
            items += [m for m in fetched if m.folder_id not in skip]
            if items or not link:  # a filtered page can be empty: keep scanning, within bounds
                break
        complete = link is None

        notes: list[str] = []
        if skip:
            notes.append(
                "received_only: Sent Items, Drafts, Outbox, Deleted Items and Junk Email are left out, "
                "so a page can hold fewer than limit messages; follow cursor for more."
            )
        if state is None:  # retained server-deleted messages are merged into the first page only
            span_since = (
                since
                if complete or not items
                else min((m.received_at for m in items if m.received_at), default=since)
            )
            deleted = self.store.window(folder_id=folder_id, since=span_since, until=until, deleted=True)
            deleted = [m for m in deleted if m.folder_id not in skip]
            if deleted:
                items = sorted(items + deleted, key=_newest_first)
                notes.append(
                    f"{len(deleted)} message(s) deleted on the server are included from local retention "
                    "(is_deleted=true)."
                )
        next_cursor = None
        if link:
            next_cursor = cursors.encode(
                "list_messages",
                link=link,
                folder_id=folder_id,
                since=since.isoformat() if since else None,
                until=until.isoformat() if until else None,
                received_only=received_only,
            )
        retained = any(m.is_deleted for m in items)
        return MessagePage(
            items=await self.decorate(items),
            cursor=next_cursor,
            coverage=Coverage(
                source="remote+local" if retained else "remote",
                complete=complete,
                more_available=not complete,
                notes=notes,
            ),
        )

    # ---------------------------------------------------------------- content

    async def message(self, message_id: str, *, body: BodyKind = "unique") -> Message:
        """The full message from the server, or the retained copy if the server no longer has it."""
        try:
            message = await self.reader.get_message(
                message_id, body_format="html" if body == "html" else "text"
            )
        except NotFound:
            retained = self.store.message(message_id)
            if retained is None:
                raise NotFound(f"Message {message_id} exists neither on the server nor locally.") from None
            self.store.mark_deleted([message_id])
            retained.is_deleted = True
            return (await self.decorate([retained]))[0]  # type: ignore[return-value]
        self.store.save_messages([message])
        return (await self.decorate([message]))[0]  # type: ignore[return-value]

    async def get_message(
        self, message_id: str, *, body: BodyKind = "unique", offset: int = 0, max_chars: int = 20000
    ) -> MessageContent:
        if offset < 0 or not 1 <= max_chars <= 200_000:
            raise InvalidRequest("offset must be >= 0 and max_chars between 1 and 200000.")
        message = await self.message(message_id, body=body)
        text = message.body(body)
        if message.is_deleted and not text:
            text = "(Deleted on the server; this app never retained its content.)"
        end = min(len(text), offset + max_chars)
        return MessageContent(
            message=MessageSummary.model_validate(message.model_dump()),
            body_kind=body,
            text=text[offset:end],
            offset=offset,
            total_chars=len(text),
            next_offset=end if end < len(text) else None,
            attachments=message.attachments,
        )

    # ---------------------------------------------------------------- search

    async def search(
        self,
        query: str,
        *,
        since: datetime | None = None,
        until: datetime | None = None,
        folder: str | None = None,
        limit: int = 25,
        cursor: str | None = None,
    ) -> SearchResult:
        if not query.strip():
            raise InvalidRequest("query must not be empty.")
        if not 1 <= limit <= 100:
            raise InvalidRequest("limit must be between 1 and 100.")
        state = cursors.decode(cursor, "search") if cursor else None
        kql = query.strip()
        if since:
            kql += f" received>={since.date().isoformat()}"
        if until:
            kql += f" received<={until.date().isoformat()}"
        folder_id = (
            state.get("folder_id") if state else ((await self.resolve_folder(folder)).id if folder else None)
        )
        items, link = await self.reader.search(
            query=state["query"] if state else kql,
            folder_id=folder_id,
            page_size=limit,
            page=state["link"] if state else None,
        )
        self.store.upsert_summaries(items)
        total = None
        # A different engine (Microsoft Search) counts differently, so its total is only shown as an
        # approximation when this result set is known to be incomplete.
        if state is None and folder_id is None and link:
            try:
                total = await self.reader.search_total(kql)
            except Exception:  # the total is a courtesy for coverage; never fail the search for it
                total = None
        groups: dict[str | None, ConversationHit] = {}
        for item in await self.decorate(items):
            hit = groups.setdefault(
                item.conversation_id or item.id,
                ConversationHit(
                    conversation_id=item.conversation_id,
                    subject=item.subject,
                    last_received_at=item.received_at,
                    matching_messages=[],
                ),
            )
            hit.matching_messages.append(item)
            if item.received_at and (hit.last_received_at is None or item.received_at > hit.last_received_at):
                hit.last_received_at = item.received_at
        return SearchResult(
            query=query,
            conversations=list(groups.values()),
            cursor=cursors.encode("search", link=link, folder_id=folder_id, query=kql) if link else None,
            coverage=Coverage(
                source="remote",
                complete=link is None,
                more_available=link is not None,
                server_total=total,
                notes=[
                    "Server-side search (Microsoft Graph); hits grouped by conversation, in rank order.",
                    "Messages retained locally after deletion on the server are not searched.",
                    *(
                        []
                        if total is None
                        else ["server_total is an approximate count of matching messages (Microsoft Search)."]
                    ),
                ],
            ),
        )


def _newest_first(m: MessageSummary) -> float:
    return -(m.received_at.timestamp() if m.received_at else 0)


def _with_paths(folders: list[Folder]) -> list[Folder]:
    by_id = {f.id: f for f in folders}
    for folder in folders:
        parts, seen, current = [], set(), folder
        while current and current.id not in seen:
            seen.add(current.id)
            parts.append(current.name)
            current = by_id.get(current.parent_id or "")
        folder.path = "/".join(reversed(parts))
    return folders

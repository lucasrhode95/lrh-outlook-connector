"""Folders, message listing, message content and search (requirements v4 §7–§9).

Remote first: every call asks Outlook (Graph) and only falls back to retained local data where
the server can no longer provide it (messages deleted on the server).

Scope rules shared by list, search, threads and export:
- Deleted Items and Junk Email are left out unless ``include_deleted_items`` (a folder asked for by
  name is always included). ``received_only`` also leaves out Sent Items, Drafts and Outbox.
- Copies of one message (same Internet message id, e.g. mail you sent to yourself or to a list you
  are on) are shown once; ``also_in`` names the folders of the other copies.
"""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Iterable
from datetime import datetime, timedelta

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.domain.models import (
    Attachment,
    BodyKind,
    ConversationHit,
    Coverage,
    Detail,
    Folder,
    Message,
    MessageContent,
    MessagePage,
    MessageSummary,
    SearchResult,
    ThreadSize,
)
from outlook_connector.remote.ports import MailReader
from outlook_connector.service import cursors
from outlook_connector.service.reconcile import Reconciler
from outlook_connector.store.db import Store

log = logging.getLogger(__name__)

FOLDER_TTL_SECONDS = 600
DELETED_OR_JUNK_FOLDERS = ("deleteditems", "junkemail")
OUTGOING_FOLDERS = ("sentitems", "drafts", "outbox")
RECEIVED_ONLY_PAGES = 10  # server pages scanned at most for one filtered page
MAX_SIZE_LOOKUPS = 200
NEVER_RETAINED = "deleted on the server and never retained by this app"
LOCAL_ONLY_NOTE = "Local cache only: messages this app has seen before. It is not a mirror of the mailbox."
COMPACT_DROP = {
    "to": [],
    "cc": [],
    "categories": [],
    "importance": None,
    "internet_message_id": None,
    "is_draft": None,
    "sent_at": None,
    "folder_id": None,
    "deleted_at": None,
}


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

    async def resolve_folder(self, ref: str) -> Folder:
        """A folder id, a well-known alias (inbox, archive, ...), a path (Inbox/Projects) or a name.

        An unknown name refreshes the folder list once, so folders created meanwhile are found.
        """
        found = _match_folder(ref, await self.folder_map())
        if found is None:
            found = _match_folder(ref, {f.id: f for f in await self.folders(refresh=True)})
        if found is None:
            raise InvalidRequest(f"Unknown folder '{ref}'. Use list_folders to see folder paths.")
        return found

    async def exclusions(self, *, include_deleted_items: bool, received_only: bool = False) -> dict[str, str]:
        """Folder id -> reason (an ExclusionReason) for every folder this scope leaves out."""
        out: dict[str, str] = {}
        for folder in (await self.folder_map()).values():
            if not include_deleted_items and folder.well_known in DELETED_OR_JUNK_FOLDERS:
                out[folder.id] = "deleted_or_junk"
            elif received_only and folder.well_known in OUTGOING_FOLDERS:
                out[folder.id] = "outgoing"
        return out

    async def decorate(self, items: list[MessageSummary]) -> list[MessageSummary]:
        folders = await self.folder_map()
        for item in items:
            folder = folders.get(item.folder_id or "")
            item.folder = folder.path if folder else None
        return items

    async def finish(
        self, items: Iterable[MessageSummary], skip: dict[str, str] | None = None
    ) -> tuple[list[MessageSummary], dict[str, int]]:
        """Apply folder exclusions, fill folder paths and merge copies. Returns (items, excluded counts)."""
        kept: list[MessageSummary] = []
        excluded: dict[str, int] = {}
        for item in items:
            reason = (skip or {}).get(item.folder_id or "")
            if reason:
                excluded[reason] = excluded.get(reason, 0) + 1
            else:
                kept.append(item)
        outgoing = {f.id for f in (await self.folder_map()).values() if f.well_known in OUTGOING_FOLDERS}
        return merge_copies(await self.decorate(kept), outgoing), excluded

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
        include_deleted_items: bool = False,
        include_total: bool = False,
        detail: Detail = "full",
    ) -> MessagePage:
        """Newest first. Scope rules in the module docstring; a named folder is listed as a whole."""
        if not 1 <= limit <= 200:
            raise InvalidRequest("limit must be between 1 and 200.")
        if since and until and since > until:
            raise InvalidRequest("since must not be after until.")
        state = cursors.decode(cursor, "list_messages") if cursor else None
        upper: datetime | None = None  # continuation: the previous page's oldest message (exclusive)
        if state:
            folder_id = state["folder_id"]
            since, until, upper = (_dt(state[k]) for k in ("since", "until", "upper"))
            received_only = bool(state["received_only"])
            include_deleted_items = bool(state["include_deleted_items"])
        else:
            folder_id = (await self.resolve_folder(folder)).id if folder else None
        skip = (
            {}
            if folder_id
            else await self.exclusions(
                include_deleted_items=include_deleted_items, received_only=received_only
            )
        )

        if not refresh:
            window = self.store.window(
                folder_id=folder_id, since=since, until=until, limit=None if skip else limit
            )
            items, excluded = await self.finish(window, skip)
            return MessagePage(
                items=_detail(items[:limit], detail),
                coverage=Coverage(source="local", complete=False, excluded=excluded, notes=[LOCAL_ONLY_NOTE]),
            )

        link = state["link"] if state else None
        fetched: list[MessageSummary] = []
        for _ in range(RECEIVED_ONLY_PAGES if skip else 1):
            page, link = await self.reader.list_messages(
                folder_id=folder_id, since=since, until=until, page_size=limit, page=link
            )
            self.store.upsert_summaries(page)
            await self.reconciler.after_window(
                folder_id=folder_id, since=since, until=until, remote=page, complete=link is None
            )
            fetched += page
            if not link or any(m.folder_id not in skip for m in page):
                break  # a page that exclusions empty entirely is skipped, within bounds
        complete = link is None

        # Retained messages deleted on the server, in the time span this page covered: [low, upper).
        low = since if complete else min((m.received_at for m in fetched if m.received_at), default=None)
        deleted: list[MessageSummary] = []
        if complete or low is not None:
            deleted = self.store.window(folder_id=folder_id, since=low, until=upper or until, deleted=True)
            if upper:
                deleted = [m for m in deleted if m.received_at and m.received_at < upper]
        items, excluded = await self.finish(sorted(fetched + deleted, key=_newest_first), skip)

        notes: list[str] = []
        if skip and not complete:
            notes.append("Folders are filtered after paging, so a page can hold fewer than limit messages.")
        if any(m.is_deleted for m in items):
            notes.append(
                "Messages with is_deleted=true were deleted on the server and come from local retention."
            )
        total = None
        if include_total and state is None:
            try:
                total = await self.reader.count_messages(
                    folder_id=folder_id, since=since, until=until, minus_folders=list(skip)
                )
            except Exception:  # a courtesy for planning; never fail the listing for it
                total = None
            if total is not None:
                notes.append(
                    "server_total counts the server's messages in scope (copies counted separately)."
                )
        next_cursor = None
        if link:
            next_cursor = cursors.encode(
                "list_messages",
                link=link,
                folder_id=folder_id,
                since=_iso(since),
                until=_iso(until),
                upper=_iso(low),
                received_only=received_only,
                include_deleted_items=include_deleted_items,
            )
        return MessagePage(
            items=_detail(items, detail),
            cursor=next_cursor,
            coverage=Coverage(
                source="remote+local" if any(m.is_deleted for m in items) else "remote",
                complete=complete,
                server_total=total,
                excluded=excluded,
                notes=notes,
            ),
        )

    async def conversation_sizes(
        self, conversation_ids: list[str], *, include_deleted_items: bool = False
    ) -> list[ThreadSize]:
        """How many messages each conversation has, counted the way get_thread lists them.

        One batched server listing of folders and Internet ids per conversation (copies counted
        once), plus retained messages deleted on the server.
        """
        ids = list(dict.fromkeys(conversation_ids))
        if len(ids) > MAX_SIZE_LOOKUPS:
            raise InvalidRequest(f"At most {MAX_SIZE_LOOKUPS} conversations per request.")
        if not ids:
            return []
        remote = await self.reader.conversation_folders(ids)
        skip = await self.exclusions(include_deleted_items=include_deleted_items)
        retained = self.store.conversations(ids)
        out = []
        for cid in ids:
            listed, more = remote.get(cid, ([], False))
            local = [(m.folder_id, m.internet_message_id) for m in retained.get(cid, []) if m.is_deleted]
            seen: set[str] = set()
            count = 0
            for folder_id, internet_id in listed + local:
                if folder_id in skip or (internet_id and internet_id in seen):
                    continue
                if internet_id:
                    seen.add(internet_id)
                count += 1
            out.append(ThreadSize(conversation_id=cid, messages=count, at_least=more))
        return out

    # ---------------------------------------------------------------- content

    async def message(self, message_id: str, *, body: BodyKind = "unique") -> Message:
        """The full message from the server, or the retained copy if the server no longer has it."""
        try:
            message = await self.reader.get_message(
                message_id, body_format="html" if body == "html" else "text"
            )
        except NotFound:
            retained = self._retained(message_id)
            return (await self.decorate([retained]))[0]  # type: ignore[return-value]
        self.store.save_messages([message])
        return (await self.decorate([message]))[0]  # type: ignore[return-value]

    async def attachments(self, message_id: str) -> list[Attachment]:
        """Attachment metadata from the server, or the retained list of a message deleted there."""
        try:
            return await self.reader.list_attachments(message_id)
        except NotFound:
            return self._retained(message_id).attachments

    def _retained(self, message_id: str) -> Message:
        """The retained copy of a message the server no longer has (marked deleted)."""
        retained = self.store.message(message_id)
        if retained is None:
            raise NotFound(f"Message {message_id} exists neither on the server nor locally.") from None
        self.store.mark_deleted([message_id])
        retained.is_deleted = True
        return retained

    async def get_message(
        self, message_id: str, *, body: BodyKind = "unique", offset: int = 0, max_chars: int = 20000
    ) -> MessageContent:
        if offset < 0 or not 1 <= max_chars <= 200_000:
            raise InvalidRequest("offset must be >= 0 and max_chars between 1 and 200000.")
        message = await self.message(message_id, body=body)
        text = message.body(body)
        if message.is_deleted and not has_content(message):
            text = unavailable(NEVER_RETAINED)
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
        received_only: bool = False,
        include_deleted_items: bool = False,
        detail: Detail = "full",
    ) -> SearchResult:
        if not query.strip():
            raise InvalidRequest("query must not be empty.")
        if not 1 <= limit <= 100:
            raise InvalidRequest("limit must be between 1 and 100.")
        state = cursors.decode(cursor, "search") if cursor else None
        if state:
            folder_id, kql = state["folder_id"], state["query"]
            since, until = _dt(state["since"]), _dt(state["until"])
            received_only = bool(state["received_only"])
            include_deleted_items = bool(state["include_deleted_items"])
        else:
            folder_id = (await self.resolve_folder(folder)).id if folder else None
            # KQL only takes dates (and their time zone is the server's): ask for a day more on each
            # side, then keep exactly the requested window below.
            kql = query.strip()
            if since:
                kql += f" received>={(since - timedelta(days=1)).date().isoformat()}"
            if until:
                kql += f" received<={(until + timedelta(days=1)).date().isoformat()}"
        skip = (
            {}
            if folder_id
            else await self.exclusions(
                include_deleted_items=include_deleted_items, received_only=received_only
            )
        )
        found, link = await self.reader.search(
            query=kql, folder_id=folder_id, page_size=limit, page=state["link"] if state else None
        )
        self.store.upsert_summaries(found)
        total = None
        # A different engine (Microsoft Search) counts differently, so its total is only shown as an
        # approximation when this result set is known to be incomplete.
        if state is None and folder_id is None and link:
            try:
                total = await self.reader.search_total(kql)
            except Exception:  # the total is a courtesy for coverage; never fail the search for it
                total = None
        in_window = [m for m in found if _within(m, since, until)]
        items, excluded = await self.finish(in_window, skip)

        groups: dict[str, ConversationHit] = {}
        for item in items:
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
        with_ids = [h.conversation_id for h in groups.values() if h.conversation_id]
        if with_ids:
            try:
                sizes = {
                    s.conversation_id: s
                    for s in await self.conversation_sizes(
                        with_ids, include_deleted_items=include_deleted_items
                    )
                }
            except Exception:  # counts are a courtesy; never fail the search for them
                sizes = {}
            for hit in groups.values():
                if (size := sizes.get(hit.conversation_id or "")) is not None:
                    hit.message_count, hit.message_count_at_least = size.messages, size.at_least
        for hit in groups.values():
            hit.matching_messages = _detail(hit.matching_messages, detail)
        notes = [
            "Server-side search (Microsoft Graph); hits grouped by conversation, in rank order.",
            "Messages retained locally after deletion on the server are not searched.",
        ]
        if total is not None:
            notes.append("server_total is an approximate count of matching messages (Microsoft Search).")
        if skip and link:
            notes.append("Folders are filtered after paging, so a page can hold fewer than limit hits.")
        return SearchResult(
            query=query,
            conversations=list(groups.values()),
            cursor=cursors.encode(
                "search",
                link=link,
                folder_id=folder_id,
                query=kql,
                since=_iso(since),
                until=_iso(until),
                received_only=received_only,
                include_deleted_items=include_deleted_items,
            )
            if link
            else None,
            coverage=Coverage(
                source="remote", complete=link is None, server_total=total, excluded=excluded, notes=notes
            ),
        )


# -------------------------------------------------------------------- helpers


def merge_copies(items: list[MessageSummary], outgoing: set[str]) -> list[MessageSummary]:
    """Show each message once: copies share an Internet message id (a self-sent mail sits in Sent
    Items and Inbox). Keeps a live copy over a retained-deleted one and a received copy over the
    sent one; ``also_in`` lists the folders of the others. Order follows the first copy."""
    groups: dict[str, list[MessageSummary]] = {}
    order: list[MessageSummary | str] = []
    for item in items:
        key = item.internet_message_id
        if not key:
            order.append(item)
        elif key in groups:
            groups[key].append(item)
        else:
            groups[key] = [item]
            order.append(key)
    out = []
    for entry in order:
        if isinstance(entry, MessageSummary):
            out.append(entry)
            continue
        copies = groups[entry]
        keep = min(copies, key=lambda m: (m.is_deleted, (m.folder_id or "") in outgoing))
        others = {m.folder or "(unknown folder)" for m in copies if m is not keep}
        others |= {folder for m in copies for folder in m.also_in}  # copies merged earlier
        keep.also_in = sorted(others - {keep.folder or ""})
        out.append(keep)
    return out


def has_content(message: Message) -> bool:
    bodies = (message.body_text, message.unique_body_text, message.body_html, message.unique_body_html)
    return any(value is not None for value in bodies)


def unavailable(reason: str) -> str:
    """The one marker for a body that cannot be shown (exports, threads, get_message)."""
    return f"(Content unavailable: {reason})"


def _detail(items: list[MessageSummary], detail: Detail) -> list[MessageSummary]:
    return items if detail == "full" else [m.model_copy(update=COMPACT_DROP) for m in items]


def _match_folder(ref: str, folders: dict[str, Folder]) -> Folder | None:
    if ref in folders:
        return folders[ref]
    needle = ref.strip().strip("/").lower()
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
    return None


def _within(m: MessageSummary, since: datetime | None, until: datetime | None) -> bool:
    stamp = m.received_at or m.sent_at
    if stamp is None:
        return since is None and until is None
    return (since is None or stamp >= since) and (until is None or stamp <= until)


def _iso(value: datetime | None) -> str | None:
    return value.isoformat() if value else None


def _dt(value: str | None) -> datetime | None:
    return datetime.fromisoformat(value) if value else None


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

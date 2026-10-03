"""Folders, message listing, message content and search (requirements v4 §7–§9).

Remote first: every call asks Outlook (Graph). The local store only caches folders and message
summaries (for the UI's instant preview and ``refresh=false``); a message deleted on the server is
gone here too.

Scope rules shared by list, search, threads, sizes and export (a folder counts with its parents):
- Deleted Items, Junk Email and Sync Issues (Outlook's own conflict and failure copies) are left out
  unless ``include_deleted_items`` (a folder asked for by name is always included).
  ``received_only`` also leaves out Sent Items, Drafts and Outbox.
- Hidden folders, and items outside the mail folders (e.g. Teams meeting records), are out of reach:
  never listed, searched, counted, threaded or exported, and list_folders does not show them.
- Copies of one message (same Internet message id, e.g. mail you sent to yourself or to a list you
  are on) are shown once; ``also_in`` names the folders of the other copies.
"""

from __future__ import annotations

import asyncio
import hashlib
import logging
from collections.abc import Iterable
from datetime import datetime, timedelta
from typing import Any

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
from outlook_connector.store.db import Store

log = logging.getLogger(__name__)

FOLDER_TTL_SECONDS = 600
DELETED_OR_JUNK_FOLDERS = ("deleteditems", "junkemail")
SYNC_ISSUES_FOLDERS = ("syncissues", "conflicts", "localfailures", "serverfailures")
OUTGOING_FOLDERS = ("sentitems", "drafts", "outbox")
RECEIVED_ONLY_PAGES = 10  # server pages scanned at most for one filtered page
MAX_SIZE_LOOKUPS = 200
SEEN_LIMIT = 400  # fingerprints a cursor carries: two pages of the largest size
GONE = "deleted on the server"
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
}


class Mailbox:
    def __init__(self, reader: MailReader, store: Store) -> None:
        self.reader = reader
        self.store = store
        self._folders: dict[str, Folder] | None = None
        self._folder_refresh: asyncio.Task[list[Folder]] | None = None
        self._outside: set[str] = set()  # folder ids a refresh confirmed are outside the mail folders

    # ---------------------------------------------------------------- folders

    async def folders(self, *, refresh: bool = False) -> list[Folder]:
        """The visible folders: cached right away, a stale cache refreshed in the background
        (stale-while-revalidate). Hidden folders are out of reach and not listed.

        Only an empty cache or ``refresh=True`` waits for the server.
        """
        cached, age = self.store.folders()
        if refresh or not cached:
            cached = await self._refresh_folders()
        elif (age is None or age > FOLDER_TTL_SECONDS) and self._folder_refresh is None:
            self._folder_refresh = asyncio.create_task(self._refresh_folders())
            self._folder_refresh.add_done_callback(self._folder_refresh_done)
        self._folders = {f.id: f for f in cached}
        categories = folder_categories(self._folders)
        visible = [f for f in cached if categories.get(f.id) != "hidden"]
        return sorted(visible, key=lambda f: f.path.lower())

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
        Hidden folders are refused: they and their items are out of reach.
        """
        found = _resolve(ref, await self.folder_map())
        if found is None:
            await self._refresh_folders()
            found = _resolve(ref, await self.folder_map())
        if found is None:
            raise InvalidRequest(f"Unknown folder '{ref}'. Use list_folders to see folder paths.")
        return found

    async def under(self, folder_id: str | None, alias: str) -> bool:
        """Whether the folder is the well-known folder ``alias`` or inside it."""
        folders = await self.folder_map()
        folder = folders.get(folder_id or "")
        return folder is not None and any(f.well_known == alias for f in _ancestry(folder, folders))

    async def exclusions(self, *, include_deleted_items: bool, received_only: bool = False) -> dict[str, str]:
        """Folder id -> reason (an ExclusionReason) for every folder this scope leaves out.

        Hidden folders are not listed here: ``finish`` always leaves them out.
        """
        left_out = set() if include_deleted_items else {"deleted_or_junk", "sync_issues"}
        if received_only:
            left_out.add("outgoing")
        categories = folder_categories(await self.folder_map())
        return {folder_id: category for folder_id, category in categories.items() if category in left_out}

    async def reach(self, folder_ids: Iterable[str | None]) -> tuple[dict[str, Folder], set[str]]:
        """The folder map and the ids of hidden folders.

        A folder id not in the map (a folder created meanwhile, or an item outside the mail folders)
        refreshes the folder list first. Ids still unknown after that are outside the mail folders;
        they are remembered and never trigger another refresh.
        """
        folders = await self.folder_map()
        unknown = {fid for fid in folder_ids if fid and fid not in folders} - self._outside
        if unknown:
            await self._refresh_folders()
            folders = await self.folder_map()
            self._outside |= {fid for fid in unknown if fid not in folders}
        hidden = {fid for fid, category in folder_categories(folders).items() if category == "hidden"}
        return folders, hidden

    async def decorate(self, items: list[MessageSummary]) -> list[MessageSummary]:
        folders = await self.folder_map()
        for item in items:
            folder = folders.get(item.folder_id or "")
            item.folder = folder.path if folder else None
        return items

    async def finish(
        self, items: Iterable[MessageSummary], skip: dict[str, str] | None = None
    ) -> tuple[list[MessageSummary], dict[str, int]]:
        """Apply folder exclusions, fill folder paths and merge copies. Returns (items, excluded counts).

        Items in hidden folders or outside the mail folders are always left out ("hidden").
        """
        items = list(items)
        folders, hidden = await self.reach(m.folder_id for m in items)
        kept: list[MessageSummary] = []
        excluded: dict[str, int] = {}
        for item in items:
            folder_id = item.folder_id or ""
            if item.folder_id and (folder_id in hidden or folder_id not in folders):
                reason: str | None = "hidden"
            else:
                reason = (skip or {}).get(folder_id)
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
        skip_returned_copies: bool = True,
    ) -> MessagePage:
        """Newest first. Scope rules in the module docstring; a named folder is listed as a whole.

        ``skip_returned_copies``: drop copies of a message an earlier page already returned. A caller
        that merges every page at the end (the export) turns it off to keep each copy's folder.
        """
        if not 1 <= limit <= 200:
            raise InvalidRequest("limit must be between 1 and 200.")
        if since and until and since > until:
            raise InvalidRequest("since must not be after until.")
        state = cursors.decode(cursor, "list_messages") if cursor else None
        if state:
            folder_id = state["folder_id"]
            since, until = _dt(state["since"]), _dt(state["until"])
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
                coverage=Coverage(complete=False, excluded=excluded, notes=[LOCAL_ONLY_NOTE]),
            )

        link = state["link"] if state else None
        fetched: list[MessageSummary] = []
        drop = set(skip) | (set() if folder_id else (await self.reach(()))[1])
        for attempt in range(RECEIVED_ONLY_PAGES if drop else 1):
            page, link = await self.reader.list_messages(
                folder_id=folder_id, since=since, until=until, page_size=limit, page=link
            )
            first, last = state is None and attempt == 0, link is None
            self._cache(page, folder_id, since=since, until=until, first=first, last=last)
            fetched += page
            if not link or any(m.folder_id not in drop for m in page):
                break  # a page that exclusions empty entirely is skipped, within bounds
        complete = link is None
        items, excluded = await self.finish(fetched, skip)
        items, seen = _skip_seen(items, state) if skip_returned_copies else (items, [])

        notes: list[str] = []
        if skip and not complete:
            notes.append("Folders are filtered after paging, so a page can hold fewer than limit messages.")
        total = None
        if include_total and state is None:
            total = await self._count(folder_id, since, until, skip)
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
                seen=seen,
                received_only=received_only,
                include_deleted_items=include_deleted_items,
            )
        return MessagePage(
            items=_detail(items, detail),
            cursor=next_cursor,
            coverage=Coverage(
                complete=complete,
                server_total=total,
                excluded=excluded,
                notes=notes,
            ),
        )

    def _cache(
        self,
        page: list[MessageSummary],
        folder_id: str | None,
        *,
        since: datetime | None,
        until: datetime | None,
        first: bool,
        last: bool,
    ) -> None:
        """Cache a listed page, and forget cached rows of the time span it covered that it does not
        hold (deleted or moved on the server). The span runs from the page's oldest to its newest
        message, extended to the window's ends on the last and the first page."""
        dated = [m.received_at for m in page if m.received_at]
        if dated or (first and last):
            low = since if last else min(dated)
            high = until if first else max(dated)
            self.store.forget(folder_id=folder_id, since=low, until=high, keep={m.id for m in page})
        self.store.upsert_summaries(page)

    async def _count(
        self, folder_id: str | None, since: datetime | None, until: datetime | None, skip: dict[str, str]
    ) -> int | None:
        """The server's count for the scope: the named folder, or every reachable folder not left out."""
        if folder_id:
            in_scope = [folder_id]
        else:
            folders, hidden = await self.reach(())
            in_scope = [fid for fid in folders if fid not in hidden and fid not in skip]
        try:
            counts = await self.reader.count_messages(folder_ids=in_scope, since=since, until=until)
        except Exception:  # a courtesy for planning; never fail the listing for it
            return None
        return sum(counts.values()) if counts is not None else None

    async def conversation_sizes(
        self, conversation_ids: list[str], *, include_deleted_items: bool = False
    ) -> list[ThreadSize]:
        """How many messages each conversation has, counted the way get_thread lists them.

        One batched server listing of folders and Internet ids per conversation (copies counted
        once).
        """
        ids = list(dict.fromkeys(conversation_ids))
        if len(ids) > MAX_SIZE_LOOKUPS:
            raise InvalidRequest(f"At most {MAX_SIZE_LOOKUPS} conversations per request.")
        if not ids:
            return []
        remote = await self.reader.conversation_folders(ids)
        skip = await self.exclusions(include_deleted_items=include_deleted_items)
        folders, hidden = await self.reach(fid for listed, _ in remote.values() for fid, _ in listed)
        out = []
        for cid in ids:
            listed, more = remote.get(cid, ([], False))
            reachable = [(f, i) for f, i in listed if f and f in folders and f not in hidden]
            seen: set[str] = set()
            count = 0
            for folder_id, internet_id in reachable:
                if folder_id in skip or (internet_id and internet_id in seen):
                    continue
                if internet_id:
                    seen.add(internet_id)
                count += 1
            out.append(ThreadSize(conversation_id=cid, messages=count, at_least=more))
        return out

    # ---------------------------------------------------------------- content

    async def message(self, message_id: str, *, body: BodyKind = "unique") -> Message:
        try:
            message = await self.reader.get_message(
                message_id, body_format="html" if body == "html" else "text"
            )
        except NotFound:
            raise NotFound(f"Message {message_id} is not on the server ({GONE}).") from None
        return (await self.decorate([message]))[0]  # type: ignore[return-value]

    async def attachments(self, message_id: str) -> list[Attachment]:
        try:
            return await self.reader.list_attachments(message_id)
        except NotFound:
            raise NotFound(f"Message {message_id} is not on the server ({GONE}).") from None

    async def get_message(
        self, message_id: str, *, body: BodyKind = "unique", offset: int = 0, max_chars: int = 20000
    ) -> MessageContent:
        if offset < 0 or not 1 <= max_chars <= 200_000:
            raise InvalidRequest("offset must be >= 0 and max_chars between 1 and 200000.")
        message = await self.message(message_id, body=body)
        text = message.body(body)
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
        in_window = [m for m in found if _within(m, since, until)]
        items, excluded = await self.finish(in_window, skip)
        items, seen = _skip_seen(items, state)

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
        notes = ["Server-side search (Microsoft Graph); hits grouped by conversation, in rank order."]
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
                seen=seen,
                since=_iso(since),
                until=_iso(until),
                received_only=received_only,
                include_deleted_items=include_deleted_items,
            )
            if link
            else None,
            coverage=Coverage(complete=link is None, excluded=excluded, notes=notes),
        )


# -------------------------------------------------------------------- helpers


def merge_copies(items: list[MessageSummary], outgoing: set[str]) -> list[MessageSummary]:
    """Show each message once: copies share an Internet message id (a self-sent mail sits in Sent
    Items and Inbox). Keeps a received copy over the sent one; ``also_in`` lists the folders of the
    others. Order follows the first copy."""
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
        keep = min(copies, key=lambda m: (m.folder_id or "") in outgoing)
        others = {m.folder or "(unknown folder)" for m in copies if m is not keep}
        others |= {folder for m in copies for folder in m.also_in}  # copies merged earlier
        keep.also_in = sorted(others - {keep.folder or ""})
        out.append(keep)
    return out


def _skip_seen(
    items: list[MessageSummary], state: dict[str, Any] | None
) -> tuple[list[MessageSummary], list[str]]:
    """Copies of one message can straddle a page boundary: drop those an earlier page returned.

    The cursor carries short fingerprints of the Internet ids returned recently, newest first, up
    to SEEN_LIMIT (copies have nearly the same timestamp, so they sit on adjacent pages).
    Returns (items, fingerprints to carry).
    """
    before: list[str] = list(state["seen"]) if state else []
    known = set(before)
    kept = [m for m in items if not (m.internet_message_id and _fingerprint(m.internet_message_id) in known)]
    current = [_fingerprint(m.internet_message_id) for m in kept if m.internet_message_id]
    return kept, list(dict.fromkeys(current + before))[:SEEN_LIMIT]


def _fingerprint(internet_id: str) -> str:
    return hashlib.blake2b(internet_id.encode(), digest_size=5).hexdigest()


def unavailable(reason: str) -> str:
    """The one marker for a body that cannot be shown (exports, threads, get_message)."""
    return f"(Content unavailable: {reason})"


def _detail(items: list[MessageSummary], detail: Detail) -> list[MessageSummary]:
    return items if detail == "full" else [m.model_copy(update=COMPACT_DROP) for m in items]


def folder_categories(folders: dict[str, Folder]) -> dict[str, str]:
    """Folder id -> category, judged from the folder and its parents (a folder deleted in Outlook
    moves into Deleted Items with its mail; Sync Issues has subfolders):
    "sync_issues" > "hidden" > "deleted_or_junk" > "outgoing". Folders without one are absent."""
    out: dict[str, str] = {}
    for folder in folders.values():
        chain = _ancestry(folder, folders)
        aliases = {f.well_known for f in chain if f.well_known}
        if aliases & set(SYNC_ISSUES_FOLDERS):
            out[folder.id] = "sync_issues"
        elif any(f.hidden for f in chain):
            out[folder.id] = "hidden"
        elif aliases & set(DELETED_OR_JUNK_FOLDERS):
            out[folder.id] = "deleted_or_junk"
        elif aliases & set(OUTGOING_FOLDERS):
            out[folder.id] = "outgoing"
    return out


def _ancestry(folder: Folder, folders: dict[str, Folder]) -> list[Folder]:
    """The folder and its parents, nearest first."""
    chain: list[Folder] = []
    seen: set[str] = set()
    current: Folder | None = folder
    while current is not None and current.id not in seen:
        seen.add(current.id)
        chain.append(current)
        current = folders.get(current.parent_id or "")
    return chain


def _resolve(ref: str, folders: dict[str, Folder]) -> Folder | None:
    """Match among the visible folders; a reference to a hidden folder is refused."""
    categories = folder_categories(folders)
    visible = {fid: f for fid, f in folders.items() if categories.get(fid) != "hidden"}
    found = _match_folder(ref, visible)
    if found is None and any(_mentions(ref, f) for fid, f in folders.items() if fid not in visible):
        raise InvalidRequest(f"'{ref}' is a hidden folder: hidden folders and their items are out of reach.")
    return found


def _mentions(ref: str, folder: Folder) -> bool:
    needle = ref.strip().strip("/").lower()
    names = {folder.well_known or "", folder.path.lower(), folder.name.lower()} - {""}
    return ref == folder.id or needle in names


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

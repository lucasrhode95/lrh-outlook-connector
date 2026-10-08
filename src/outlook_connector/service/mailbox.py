"""Folders, message listing, message content and search (requirements v4 §7–§9).

Remote first: every call asks Outlook (Graph). The local store caches the folder list only; no
message data is kept locally, so a message deleted on the server is gone here too.

Scope filters list, search, conversation reads, sizes and folder/date-window exports.
Selected export conversations and message ids stay whole. A folder counts with its
parents:
- Deleted Items and Junk Email are left out unless ``scope.deleted_items`` (a folder asked for by
  name is always included).
  Sent Items, Drafts and Outbox are included unless ``scope.sent_items`` is false. List and
  search also leave out meeting mail (invitations, replies to them, cancellations) when
  ``scope.meeting_mail`` is false: a conversation that is only meeting traffic disappears, one
  with real replies shows through them. Every flag points the same way: true shows more mail.
- Hidden folders, Sync Issues (classic Outlook's conflict and failure copies, decided 2026-10-04),
  and items outside the mail folders (e.g. Teams meeting records) are out of reach:
  never listed, searched, counted, included in conversations or folder/date-window exports, and
  list_folders does
  not show them.
- Copies of one message (same Internet message id, e.g. mail you sent to yourself or to a list you
  are on) are shown once; ``also_in`` names the folders of the other copies.
"""

from __future__ import annotations

import asyncio
import hashlib
import time
from collections.abc import Iterable
from datetime import UTC, datetime, timedelta
from math import ceil
from typing import Any

from outlook_connector.domain.errors import ConnectorError, InvalidRequest, NotFound
from outlook_connector.domain.models import (
    DEFAULT_SCOPE,
    Attachment,
    BodyKind,
    ConversationHit,
    ConversationSize,
    Coverage,
    Detail,
    Folder,
    Message,
    MessageContent,
    MessagePage,
    MessageSummary,
    Scope,
    SearchResult,
    UserProfile,
)
from outlook_connector.remote.ports import FolderCount, MailReader
from outlook_connector.service import cursors
from outlook_connector.service.scope import validate_scope
from outlook_connector.store.db import Store

FOLDER_TTL_SECONDS = 600
DELETED_OR_JUNK_FOLDERS = ("deleteditems", "junkemail")
SYNC_ISSUES_FOLDERS = ("syncissues", "conflicts", "localfailures", "serverfailures")  # out of reach
OUTGOING_FOLDERS = ("sentitems", "drafts", "outbox")
FILTERED_PAGES = 10  # server pages scanned at most for one filtered page
PER_FOLDER_SHARE = 2 / 3  # list folder by folder when left-out folders hold this share of the mailbox
MAX_SIZE_LOOKUPS = 200
SEEN_LIMIT = 400  # fingerprints a cursor carries: two pages of the largest size
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
        self._folders_at = 0.0  # time.monotonic() when _folders was last loaded
        self._folder_lock = asyncio.Lock()  # one folder refresh at a time per process
        self._outside: set[str] = set()  # folder ids a refresh confirmed are outside the mail folders
        self._profile: UserProfile | None = None
        self._photo: tuple[bytes | None] | None = None  # (photo or None,) once looked up

    # ---------------------------------------------------------------- folders

    async def folders(self, *, refresh: bool = False) -> list[Folder]:
        """The visible folders (hidden folders are out of reach and not listed).

        The cached list is used only while it is fresh (younger than FOLDER_TTL_SECONDS, whichever
        process saved it); an older or empty cache, or ``refresh=True``, waits for the server (a
        full refresh takes under a second). Processes are short-lived and often start after days
        idle, so a stale list is never used: it would miss folders created meanwhile.
        """
        async with self._folder_lock:
            cached, age = self.store.folders()
            if refresh or not cached or age is None or age > FOLDER_TTL_SECONDS:
                cached = await self._fetch_folders()
            else:
                self._keep(cached)
        categories = folder_categories(self._folders or {})
        visible = [f for f in cached if categories.get(f.id) != "hidden"]
        return sorted(visible, key=lambda f: f.path.lower())

    async def _refresh_folders(self) -> list[Folder]:
        async with self._folder_lock:
            return await self._fetch_folders()

    async def _fetch_folders(self) -> list[Folder]:
        fresh = _with_paths(await self.reader.list_folders())
        self.store.save_folders(fresh)
        self._keep(fresh)
        return fresh

    def _keep(self, folders: list[Folder]) -> None:
        self._folders = {f.id: f for f in folders}
        self._folders_at = time.monotonic()

    async def folder_map(self) -> dict[str, Folder]:
        """The folders by id, fresh: reloaded once FOLDER_TTL_SECONDS old, also within a process."""
        if self._folders is None or time.monotonic() - self._folders_at > FOLDER_TTL_SECONDS:
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

    async def exclusions(self, scope: Scope) -> dict[str, str]:
        """Folder id -> reason (an ExclusionReason) for every folder this scope leaves out.

        Hidden folders are not listed here: ``finish`` always leaves them out.
        """
        left_out = set() if scope.deleted_items else {"deleted_or_junk"}
        if not scope.sent_items:
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

        Assumes (not re-checked here): ``skip`` comes from ``exclusions()`` for the same scope as the
        items.
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
        cursor: str | None = None,
        scope: Scope = DEFAULT_SCOPE,
        include_total: bool = False,
        detail: Detail = "full",
        skip_returned_copies: bool = True,
    ) -> MessagePage:
        """Newest first. Scope rules in the module docstring; a named folder is listed as a whole.

        ``skip_returned_copies``: drop copies of a message an earlier page already returned. A caller
        that merges every page at the end (the export) turns it off to keep each copy's folder.

        Entry point: the authoritative check of its arguments (scope, limit, window, folder, cursor); the web
        routes validate the scope and pass other inputs through, and the helpers below trust them.
        """
        if not 1 <= limit <= 200:
            raise InvalidRequest("limit must be between 1 and 200.")
        validate_scope(scope)
        if since and until and since > until:
            raise InvalidRequest("since must not be after until.")
        state = cursors.decode(cursor, "list_messages") if cursor else None
        if state:
            folder_id = state["folder_id"]
            since, until = _dt(state["since"]), _dt(state["until"])
            scope = Scope.model_validate(state["scope"])
        else:
            folder_id = (await self.resolve_folder(folder)).id if folder else None
        skip = {} if folder_id else await self.exclusions(scope)

        notes: list[str] = []
        window_excluded: dict[str, int] = {}
        folder_counts: dict[str, FolderCount] | None = None
        link: str | None = None
        offsets: dict[str, int] | None = None  # per-folder listing: folder id -> messages returned
        hidden = set() if folder_id else (await self.reach(()))[1]
        if state:
            link, offsets = state.get("link"), state.get("offsets")
            if offsets:  # keep only folders still there and still in scope (a folder may be deleted)
                folders = await self.folder_map()
                offsets = {
                    f: n for f, n in offsets.items() if f in folders and f not in skip and f not in hidden
                }
        elif not folder_id:
            share, note = await self._left_out_share(set(skip) | hidden)
            if share >= PER_FOLDER_SHARE:
                offsets, window_excluded, folder_counts = await self._folders_with_mail(
                    set(skip) | hidden, since, until
                )
                notes.append(note)
        if offsets is not None:
            fetched, offsets = await self._merged_folders(offsets, since, until, limit, folder_counts)
            complete = not offsets
            filtered = not scope.meeting_mail
        else:
            fetched = []
            drop = set(skip) | hidden

            def kept(m: MessageSummary) -> bool:
                return m.folder_id not in drop and (scope.meeting_mail or m.meeting is None)

            filtered = bool(drop) or not scope.meeting_mail
            for _ in range(FILTERED_PAGES if filtered else 1):
                page, link = await self.reader.list_messages(
                    folder_id=folder_id, since=since, until=until, page_size=limit, page=link
                )
                fetched += page
                if not link or any(kept(m) for m in page):
                    break  # a page that exclusions empty entirely is skipped, within bounds
            complete = link is None
        items, excluded = await self.finish(fetched, skip)
        excluded.update(window_excluded)  # full-window counts only on the first per-folder page
        items = _without_meetings(items, excluded) if not scope.meeting_mail else items
        items, seen = _skip_seen(items, state) if skip_returned_copies else (items, [])

        if filtered and not complete:
            notes.append("Filters apply after paging, so a page can hold fewer than limit messages.")
        total = None
        if include_total and state is None:
            total = await self._count(folder_id, since, until, skip, known_counts=folder_counts)
            if total is not None:
                notes.append(
                    "server_total counts the server's messages in scope (copies counted separately). "
                    "It includes meeting mail even when scope.meeting_mail=false hides it."
                )
        next_cursor = None
        if not complete:
            next_cursor = cursors.encode(
                "list_messages",
                link=link,
                offsets=offsets,
                folder_id=folder_id,
                since=_iso(since),
                until=_iso(until),
                seen=seen,
                scope=scope.model_dump(mode="json"),
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

    async def _left_out_share(self, left_out: set[str]) -> tuple[float, str]:
        """The share of the mailbox's messages in folders a mailbox-wide listing leaves out, from
        the cached folder counts (no request), and the note that explains a per-folder listing.

        Assumes (not re-checked here): ``left_out`` is the caller's complete set of left-out folder ids
        (the scope's exclusions plus hidden folders).
        """
        folders = await self.folder_map()
        total = sum(f.total or 0 for f in folders.values())
        out = sorted(
            (f for fid, f in folders.items() if fid in left_out and f.total), key=lambda f: -(f.total or 0)
        )
        share = sum(f.total or 0 for f in out) / total if total else 0.0
        biggest = ", ".join(f"{f.path}: {f.total:,}" for f in out[:3])
        note = (
            f"Listed folder by folder: folders this listing leaves out hold {share:.0%} of the mailbox "
            f"({biggest}), which makes a whole-mailbox listing slow. Emptying Junk Email or Deleted Items "
            "makes it faster."
        )
        return share, note

    async def _folders_with_mail(
        self, left_out: set[str], since: datetime | None, until: datetime | None
    ) -> tuple[dict[str, int], dict[str, int], dict[str, FolderCount]]:
        """Initial folder offsets and excluded Deleted/Junk counts, from the same count batch.

        Assumes (not re-checked here): complete left-out folder ids and a validated listing window.
        A folder whose count is unavailable is chosen by its cached total; the excluded count is
        claimed only when every excluded folder was counted.
        """
        folders = await self.folder_map()
        in_scope = [fid for fid in folders if fid not in left_out]
        categories = folder_categories(folders)
        excluded_folders = [
            fid for fid in folders if fid in left_out and categories.get(fid) == "deleted_or_junk"
        ]
        try:
            counts: dict[str, FolderCount] = await self.reader.count_messages(
                folder_ids=in_scope + excluded_folders, since=since, until=until
            )
        except ConnectorError:
            counts = {}
        offsets = {fid: 0 for fid in in_scope if (counts[fid].count if fid in counts else folders[fid].total)}
        excluded: dict[str, int] = {}
        if all(fid in counts for fid in excluded_folders):
            excluded_count = sum(counts[fid].count for fid in excluded_folders)
            excluded = {"deleted_or_junk": excluded_count} if excluded_count else {}
        return offsets, excluded, counts

    async def _merged_folders(
        self,
        offsets: dict[str, int],
        since: datetime | None,
        until: datetime | None,
        limit: int,
        folder_counts: dict[str, FolderCount] | None = None,
    ) -> tuple[list[MessageSummary], dict[str, int]]:
        """The newest ``limit`` messages across the folders, merged newest first, and each folder's
        new position (folders with nothing left are dropped). Each folder is read from its position
        in count-weighted chunks, so busy folders often need just one read. The count batch's newest
        dates defer folders that cannot contribute; missing dates stay eligible. Reads are parallel
        within the transport's existing limit of 4.

        Assumes (not re-checked here): ``offsets`` holds only in-scope folders (``list_messages`` drops
        left-out, hidden and deleted ones), and ``limit`` and the window were validated by
        ``list_messages``.
        """
        positions = dict(offsets)
        read = dict(offsets)  # how far each folder has been read
        buffers: dict[str, list[MessageSummary]] = {fid: [] for fid in offsets}
        more = dict.fromkeys(offsets, True)
        folders = await self.folder_map()
        weights = {
            fid: max(
                1,
                folder_counts[fid].count
                if folder_counts and fid in folder_counts
                else folders[fid].total or 0,
            )
            for fid in offsets
        }
        total_weight = sum(weights.values()) or 1
        chunk = {
            fid: min(limit, max(1, ceil(limit * weight / total_weight) + 1))
            for fid, weight in weights.items()
        }
        upper_bounds: dict[str, float | None] = {}
        for fid in offsets:
            folder_count = folder_counts.get(fid) if folder_counts is not None else None
            newest = folder_count.newest_received_at if folder_count is not None else None
            upper_bounds[fid] = newest.timestamp() if newest is not None else None

        gone: set[str] = set()

        async def fill(fid: str) -> None:
            size = min(chunk[fid], limit)
            try:
                page, link = await self.reader.list_messages(
                    folder_id=fid, since=since, until=until, page_size=size, page=None, skip=read[fid]
                )
            except NotFound:  # the folder was deleted meanwhile: its mail is gone or in Deleted Items
                gone.add(fid)
                more[fid] = False
                upper_bounds[fid] = float("-inf")
                return
            buffers[fid] += page
            read[fid] += len(page)
            more[fid] = bool(link and page)
            upper_bounds[fid] = _stamp(page[-1]) if page else float("-inf")
            chunk[fid] = size * 2

        taken: list[MessageSummary] = []

        def cutoff() -> float | None:
            known = taken + [message for buffer in buffers.values() for message in buffer]
            if len(known) < limit:
                return None
            ordered = sorted(known, key=_stamp, reverse=True)
            return _stamp(ordered[limit - 1])

        while len(taken) < limit:
            empty = [fid for fid in buffers if not buffers[fid] and more[fid]]
            if empty:
                current_cutoff = cutoff()
                if current_cutoff is None:
                    pending = sorted(
                        empty,
                        key=lambda fid: (
                            upper_bounds[fid] is None,
                            upper_bounds[fid] if upper_bounds[fid] is not None else float("inf"),
                            weights[fid],
                        ),
                        reverse=True,
                    )
                    selected: list[str] = []
                    expected = len(taken) + sum(len(buffer) for buffer in buffers.values())
                    for fid in pending:
                        selected.append(fid)
                        expected += chunk[fid]
                        if expected >= limit:
                            break
                else:
                    selected = []
                    for fid in empty:
                        newest = upper_bounds[fid]
                        if newest is None or newest >= current_cutoff:
                            selected.append(fid)
                if selected:
                    await asyncio.gather(*(fill(fid) for fid in selected))
                    continue
            heads = [fid for fid in buffers if buffers[fid]]
            if not heads:
                break
            newest = max(heads, key=lambda fid: _stamp(buffers[fid][0]))
            taken.append(buffers[newest].pop(0))
            positions[newest] += 1
        left = {fid: positions[fid] for fid in positions if buffers[fid] or more[fid]}
        if gone:
            await self._refresh_folders()  # so the rest of the call sees the folder tree as it is now
        return taken, left

    async def _count(
        self,
        folder_id: str | None,
        since: datetime | None,
        until: datetime | None,
        skip: dict[str, str],
        known_counts: dict[str, FolderCount] | None = None,
    ) -> int | None:
        """The server's count for the scope: the named folder, or every reachable folder not left out.
        Complete per-folder counts are reused when available.

        Assumes (not re-checked here): ``folder_id`` was resolved and ``skip`` built by ``list_messages``
        for the same scope.
        """
        if folder_id:
            in_scope = [folder_id]
        else:
            folders, hidden = await self.reach(())
            in_scope = [fid for fid in folders if fid not in hidden and fid not in skip]
        if known_counts is not None and all(fid in known_counts for fid in in_scope):
            return sum(known_counts[fid].count for fid in in_scope)
        try:
            counts = await self.reader.count_messages(folder_ids=in_scope, since=since, until=until)
        except Exception:  # a courtesy for planning; never fail the listing for it
            return None
        return sum(count.count for count in counts.values()) if len(counts) == len(set(in_scope)) else None

    async def conversation_sizes(
        self, conversation_ids: list[str], *, scope: Scope = DEFAULT_SCOPE
    ) -> list[ConversationSize]:
        """How many messages each conversation has, counted the way get_conversation lists them.

        One batched server listing of folders and Internet ids per conversation (copies counted
        once).

        Entry point: validates scope and its own arguments (at most MAX_SIZE_LOOKUPS conversations,
        duplicates dropped).
        """
        validate_scope(scope, sent_items=False, meeting_mail=False)
        ids = list(dict.fromkeys(conversation_ids))
        if len(ids) > MAX_SIZE_LOOKUPS:
            raise InvalidRequest(f"At most {MAX_SIZE_LOOKUPS} conversations per request.")
        if not ids:
            return []
        remote = await self.reader.conversation_folders(ids)
        skip = await self.exclusions(scope)
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
            out.append(ConversationSize(conversation_id=cid, messages=count, at_least=more))
        return out

    # ---------------------------------------------------------------- content

    # ---------------------------------------------------------------- the signed-in user

    async def profile(self) -> UserProfile:
        """Display name and address, read once per process."""
        if self._profile is None:
            self._profile = await self.reader.profile()
        return self._profile

    async def photo(self) -> bytes | None:
        """The user's small profile photo (None when none is set), read once per process; kept in
        memory only."""
        if self._photo is None:
            self._photo = (await self.reader.profile_photo(),)
        return self._photo[0]

    async def message(self, message_id: str, *, body: BodyKind = "unique") -> Message:
        try:
            message = await self.reader.get_message(
                message_id, body_format="html" if body == "html" else "text"
            )
        except NotFound:
            raise NotFound(
                f"Message {message_id} was not found; it may have been deleted, moved out of reach, "
                "or the id may be wrong."
            ) from None
        return (await self.decorate([message]))[0]  # type: ignore[return-value]

    async def attachments(self, message_id: str) -> list[Attachment]:
        try:
            return await self.reader.list_attachments(message_id)
        except NotFound:
            raise NotFound(
                f"Message {message_id} was not found; it may have been deleted, moved out of reach, "
                "or the id may be wrong."
            ) from None

    async def attachments_many(self, message_ids: list[str]) -> dict[str, list[Attachment]]:
        """Attachments of many messages at once (batched), for the list's file chips. Messages whose
        listing failed are left out: the chips are a convenience.

        Entry point: validates its own arguments (at most MAX_SIZE_LOOKUPS messages, duplicates dropped).
        """
        ids = list(dict.fromkeys(message_ids))
        if len(ids) > MAX_SIZE_LOOKUPS:
            raise InvalidRequest(f"At most {MAX_SIZE_LOOKUPS} messages per request.")
        found, _ = await self.reader.list_attachments_many(ids) if ids else ({}, {})
        return found

    async def get_message(
        self,
        message_id: str,
        *,
        body: BodyKind = "unique",
        offset: int = 0,
        max_chars: int | None = 20000,
    ) -> MessageContent:
        """One message's body with offset continuation, and its attachment metadata. ``max_chars=None``
        returns the whole body from ``offset`` (the local web reader: one server read, no token budget).

        Entry point: the authoritative check of ``offset`` and ``max_chars`` (the web routes pass them
        through unchecked). The id is taken as given (ids come from this connector's own results).
        """
        if offset < 0 or (max_chars is not None and not 1 <= max_chars <= 200_000):
            raise InvalidRequest("offset must be >= 0 and max_chars between 1 and 200000.")
        message = await self.message(message_id, body=body)
        text = message.body(body)
        end = len(text) if max_chars is None else min(len(text), offset + max_chars)
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
        scope: Scope = DEFAULT_SCOPE,
        detail: Detail = "full",
    ) -> SearchResult:
        """Server-side search, hits grouped by conversation (requirements v4 §9).

        Entry point: normalize naive dates as UTC and aware dates to UTC; validate scope, query, limit,
        folder and cursor. The web and MCP search routes pass dates through unchecked.
        """
        if not query.strip():
            raise InvalidRequest("query must not be empty.")
        if not 1 <= limit <= 100:
            raise InvalidRequest("limit must be between 1 and 100.")
        validate_scope(scope)
        since, until = _search_date(since), _search_date(until)
        state = cursors.decode(cursor, "search") if cursor else None
        if state:
            folder_id, kql = state["folder_id"], state["query"]
            since, until = _dt(state["since"]), _dt(state["until"])
            scope = Scope.model_validate(state["scope"])
        else:
            folder_id = (await self.resolve_folder(folder)).id if folder else None
            # KQL only takes dates (and their time zone is the server's): ask for a day more on each
            # side, then keep exactly the requested window below.
            kql = query.strip()
            if since:
                kql += f" received>={(since - timedelta(days=1)).date().isoformat()}"
            if until:
                kql += f" received<={(until + timedelta(days=1)).date().isoformat()}"
        skip = {} if folder_id else await self.exclusions(scope)
        found, link = await self.reader.search(
            query=kql, folder_id=folder_id, page_size=limit, page=state["link"] if state else None
        )
        in_window = [m for m in found if _within(m, since, until)]
        items, excluded = await self.finish(in_window, skip)
        items = _without_meetings(items, excluded) if not scope.meeting_mail else items
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
                        with_ids, scope=Scope(deleted_items=scope.deleted_items)
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
        if (skip or not scope.meeting_mail) and link:
            notes.append("Filters apply after paging, so a page can hold fewer than limit hits.")
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
                scope=scope.model_dump(mode="json"),
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


def _without_meetings(items: list[MessageSummary], excluded: dict[str, int]) -> list[MessageSummary]:
    """Leave out meeting mail, counting it in ``excluded``."""
    kept = [m for m in items if m.meeting is None]
    if len(kept) < len(items):
        excluded["meeting_mail"] = excluded.get("meeting_mail", 0) + len(items) - len(kept)
    return kept


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


def _detail(items: list[MessageSummary], detail: Detail) -> list[MessageSummary]:
    return items if detail == "full" else [m.model_copy(update=COMPACT_DROP) for m in items]


def folder_categories(folders: dict[str, Folder]) -> dict[str, str]:
    """Folder id -> category, judged from the folder and its parents (a folder deleted in Outlook
    moves into Deleted Items with its mail; Sync Issues has subfolders):
    "hidden" (also Sync Issues, which Graph does not mark hidden) > "deleted_or_junk" > "outgoing".
    Folders without one are absent."""
    out: dict[str, str] = {}
    for folder in folders.values():
        chain = _ancestry(folder, folders)
        aliases = {f.well_known for f in chain if f.well_known}
        if aliases & set(SYNC_ISSUES_FOLDERS) or any(f.hidden for f in chain):
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


def _stamp(m: MessageSummary) -> float:
    stamp = m.received_at or m.sent_at
    return stamp.timestamp() if stamp else float("-inf")


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


def _search_date(value: datetime | None) -> datetime | None:
    """Assumes (not re-checked here): optional datetime passed to the search entry point."""
    if value is None:
        return None
    return value.replace(tzinfo=UTC) if value.tzinfo is None else value.astimezone(UTC)

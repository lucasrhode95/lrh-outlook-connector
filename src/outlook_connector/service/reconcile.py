"""Server-side deletion and move detection (research §3.6).

A message missing from a remote listing is not proof of deletion: it may have moved, and Graph
even reports soft deletes as "deleted". So every candidate is looked up by immutable id:
404 → tombstone (known content is kept); found elsewhere → folder updated.
"""

from __future__ import annotations

from collections.abc import Iterable
from datetime import datetime

from outlook_connector.domain.models import MessageSummary
from outlook_connector.remote.ports import MailReader
from outlook_connector.store.db import Store

MAX_LOOKUPS = 200  # bound the extra work a single listing can trigger


class Reconciler:
    def __init__(self, reader: MailReader, store: Store) -> None:
        self._reader = reader
        self._store = store

    async def after_window(
        self,
        *,
        folder_id: str | None,
        since: datetime | None,
        until: datetime | None,
        remote: list[MessageSummary],
        complete: bool,
    ) -> int:
        """Reconcile retained rows of a listed window. Returns how many rows changed state.

        With an incomplete listing, only the time span the remote page actually covered is checked.
        """
        if not complete:
            dated = [m.received_at for m in remote if m.received_at]
            if not dated:
                return 0
            since, until = min(dated), max(dated)
        local = self._store.window(folder_id=folder_id, since=since, until=until, deleted=False)
        return await self.resolve(_missing(local, remote))

    async def after_conversation(self, conversation_id: str, remote: list[MessageSummary]) -> int:
        local = [m for m in self._store.conversation(conversation_id) if not m.is_deleted]
        return await self.resolve(_missing(local, remote))

    async def resolve(self, message_ids: list[str]) -> int:
        if not message_ids:
            return 0
        located = await self._reader.locate(message_ids[:MAX_LOOKUPS])
        gone = [mid for mid, folder in located.items() if folder is None]
        moved = {mid: folder for mid, folder in located.items() if folder is not None}
        if gone:
            self._store.mark_deleted(gone)
        if moved:
            self._store.set_folders(moved)
        return len(located)


def _missing(local: Iterable[MessageSummary], remote: list[MessageSummary]) -> list[str]:
    seen = {m.id for m in remote}
    return [m.id for m in local if m.id not in seen]

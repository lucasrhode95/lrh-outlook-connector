"""Account-bound SQLite store (architecture §5.7).

Holds the account binding, the folder cache and a cache of message summaries this app has
listed (for the UI's instant preview and ``refresh=false``). No bodies, no attachment bytes, no
mailbox mirror: a message deleted on the server is dropped from the cache when a listing that
covered it no longer has it. Several short-lived processes may share it: WAL mode, busy timeout,
one connection per operation, short transactions.

The store is a reconstructable cache. If SQLite reports it damaged when it is opened, the damaged
files are moved to a ``corrupt-<timestamp>`` folder next to it and a fresh store is started.
"""

from __future__ import annotations

import contextlib
import logging
import os
import sqlite3
import time
from collections.abc import Iterable, Iterator
from contextlib import contextmanager
from datetime import UTC, datetime
from pathlib import Path

from outlook_connector import config
from outlook_connector.domain.errors import AccountMismatch, ConnectorError
from outlook_connector.domain.models import DERIVED_FIELDS, Folder, MessageSummary

SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS folders (
    id TEXT PRIMARY KEY,
    data TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS messages (
    id TEXT PRIMARY KEY,
    conversation_id TEXT,
    folder_id TEXT,
    received_at TEXT,
    summary TEXT NOT NULL,
    updated_at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS messages_folder_received ON messages (folder_id, received_at);
CREATE INDEX IF NOT EXISTS messages_received ON messages (received_at);
"""


log = logging.getLogger(__name__)

# The summary column holds exactly the summary fields: never bodies, never per-result fields.
SUMMARY_COLUMNS = frozenset(MessageSummary.model_fields) - DERIVED_FIELDS


def store_path(fingerprint: str) -> Path:
    return config.data_dir() / "accounts" / fingerprint / "mail.sqlite3"


def _iso(value: datetime | None) -> str | None:
    return value.astimezone(UTC).isoformat() if value else None


class Store:
    def __init__(self, path: Path, owner: str) -> None:
        self.path = path
        self.owner = owner
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists() and not self._healthy():
            self._quarantine()
        with contextlib.closing(sqlite3.connect(self.path, timeout=15)) as db:
            db.execute("PRAGMA journal_mode=WAL")  # persistent: set once per database file
        with self._tx() as db:
            db.executescript(SCHEMA)
            row = db.execute("SELECT value FROM meta WHERE key = 'owner'").fetchone()
            if row is None:
                db.execute("INSERT INTO meta (key, value) VALUES ('owner', ?)", (owner,))
            elif row[0] != owner:
                raise AccountMismatch(f"The local store at {path} belongs to a different Microsoft account.")

    def _healthy(self) -> bool:
        try:
            db = sqlite3.connect(self.path, timeout=15)
            try:
                return db.execute("PRAGMA quick_check").fetchone()[0] == "ok"
            finally:
                db.close()
        except sqlite3.DatabaseError:
            return False

    def _quarantine(self) -> None:
        """Move a damaged store (and its WAL files) aside so a fresh one can start."""
        target = self.path.parent / f"corrupt-{time.strftime('%Y%m%d-%H%M%S')}"
        target.mkdir(exist_ok=True)
        try:
            for suffix in ("", "-wal", "-shm"):
                source = Path(f"{self.path}{suffix}")
                if source.exists():
                    os.replace(source, target / source.name)
        except PermissionError:
            raise ConnectorError(
                "The local store is damaged and another outlook-connector process is using it. "
                "Close the other connector processes (UI, MCP sessions) and retry."
            ) from None
        log.warning("Local store was damaged; moved to %s and starting a fresh one.", target)

    @contextmanager
    def _tx(self) -> Iterator[sqlite3.Connection]:
        db = sqlite3.connect(self.path, timeout=15)  # timeout = busy timeout for concurrent processes
        try:
            with db:  # one transaction
                yield db
        finally:
            db.close()

    # ---------------------------------------------------------------- folders

    def save_folders(self, folders: Iterable[Folder]) -> None:
        with self._tx() as db:
            db.execute("DELETE FROM folders")
            db.executemany(
                "INSERT INTO folders (id, data) VALUES (?, ?)", [(f.id, f.model_dump_json()) for f in folders]
            )
            db.execute(
                "INSERT OR REPLACE INTO meta (key, value) VALUES ('folders_at', ?)", (str(time.time()),)
            )

    def folders(self) -> tuple[list[Folder], float | None]:
        """Cached folders and their age in seconds (None when never cached)."""
        with self._tx() as db:
            rows = db.execute("SELECT data FROM folders").fetchall()
            at = db.execute("SELECT value FROM meta WHERE key = 'folders_at'").fetchone()
        age = time.time() - float(at[0]) if at else None
        return [Folder.model_validate_json(r[0]) for r in rows], age

    # ---------------------------------------------------------------- messages

    def upsert_summaries(self, items: Iterable[MessageSummary]) -> None:
        now = time.time()
        rows = [
            (m.id, m.conversation_id, m.folder_id, _iso(m.received_at),
             m.model_dump_json(include=set(SUMMARY_COLUMNS)), now)
            for m in items
        ]  # fmt: skip
        with self._tx() as db:
            db.executemany(
                """INSERT INTO messages (id, conversation_id, folder_id, received_at, summary, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT (id) DO UPDATE SET conversation_id = excluded.conversation_id,
                       folder_id = excluded.folder_id, received_at = excluded.received_at,
                       summary = excluded.summary, updated_at = excluded.updated_at""",
                rows,
            )

    def summaries(self, ids: Iterable[str]) -> dict[str, MessageSummary]:
        values = list(dict.fromkeys(ids))
        rows: list[tuple[str]] = []
        with self._tx() as db:
            for start in range(0, len(values), 500):  # well under SQLite's bound-parameter limit
                chunk = values[start : start + 500]
                rows += db.execute(
                    f"SELECT summary FROM messages WHERE id IN ({','.join('?' * len(chunk))})", chunk
                ).fetchall()
        return {m.id: m for m in (MessageSummary.model_validate_json(r[0]) for r in rows)}

    def window(
        self,
        *,
        folder_id: str | None,
        since: datetime | None,
        until: datetime | None,
        limit: int | None = None,
    ) -> list[MessageSummary]:
        """Cached summaries newest first."""
        where, args = _where(folder_id, since, until)
        sql = f"SELECT summary FROM messages WHERE {where} ORDER BY received_at DESC"
        if limit:
            sql += " LIMIT ?"
            args.append(limit)
        with self._tx() as db:
            rows = db.execute(sql, args).fetchall()
        return [MessageSummary.model_validate_json(r[0]) for r in rows]

    def drop(self, ids: Iterable[str]) -> None:
        """Forget cached summaries (they changed on the server; the next listing caches them again)."""
        with self._tx() as db:
            db.executemany("DELETE FROM messages WHERE id = ?", [(i,) for i in ids])

    def forget(
        self, *, folder_id: str | None, since: datetime | None, until: datetime | None, keep: set[str]
    ) -> None:
        """Drop cached summaries in the window (inclusive) except ``keep``: a listing covered the
        window and no longer has them."""
        where, args = _where(folder_id, since, until)
        with self._tx() as db:
            stale = [r[0] for r in db.execute(f"SELECT id FROM messages WHERE {where}", args)]
            db.executemany("DELETE FROM messages WHERE id = ?", [(i,) for i in stale if i not in keep])


def _where(folder_id: str | None, since: datetime | None, until: datetime | None) -> tuple[str, list[object]]:
    clauses, args = ["1=1"], []
    if folder_id:
        clauses.append("folder_id = ?")
        args.append(folder_id)
    if since:
        clauses.append("received_at >= ?")
        args.append(_iso(since))
    if until:
        clauses.append("received_at <= ?")
        args.append(_iso(until))
    return " AND ".join(clauses), args

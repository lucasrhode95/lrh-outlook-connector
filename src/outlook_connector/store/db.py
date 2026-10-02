"""Account-bound SQLite store (architecture §5.7).

Holds only what Outlook cannot give back: the folder cache, retained copies of messages the
app has seen (so they survive server-side deletion), and deletion markers. No mailbox mirror,
no attachment bytes. Several short-lived processes may share it: WAL mode, busy timeout,
one connection per operation, short transactions.

The store is a reconstructable cache. If SQLite reports it damaged when it is opened, the damaged
files are moved to a ``corrupt-<timestamp>`` folder next to it and a fresh store is started.
"""

from __future__ import annotations

import contextlib
import json
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
from outlook_connector.domain.models import DERIVED_FIELDS, Folder, Message, MessageSummary

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
    is_deleted INTEGER NOT NULL DEFAULT 0,
    deleted_at TEXT,
    summary TEXT NOT NULL,
    content TEXT,
    updated_at REAL NOT NULL
);
CREATE INDEX IF NOT EXISTS messages_conversation ON messages (conversation_id);
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
        """Record messages seen on the server. Seeing one again clears any deletion marker."""
        now = time.time()
        rows = [
            (m.id, m.conversation_id, m.folder_id, _iso(m.received_at),
             m.model_dump_json(include=SUMMARY_COLUMNS), now)
            for m in items
        ]  # fmt: skip
        with self._tx() as db:
            db.executemany(
                """INSERT INTO messages (id, conversation_id, folder_id, received_at, summary, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT (id) DO UPDATE SET conversation_id = excluded.conversation_id,
                       folder_id = excluded.folder_id, received_at = excluded.received_at,
                       summary = excluded.summary, updated_at = excluded.updated_at,
                       is_deleted = 0, deleted_at = NULL""",
                rows,
            )

    def save_messages(self, messages: Iterable[Message]) -> None:
        """Retain full content (bodies, attachment metadata). Also refreshes the summary."""
        messages = list(messages)
        self.upsert_summaries(messages)
        with self._tx() as db:
            for m in messages:
                existing = db.execute("SELECT content FROM messages WHERE id = ?", (m.id,)).fetchone()
                merged = _merge_content(existing[0] if existing else None, m)
                db.execute("UPDATE messages SET content = ? WHERE id = ?", (merged, m.id))

    def message(self, message_id: str) -> Message | None:
        with self._tx() as db:
            row = db.execute(
                "SELECT summary, content, is_deleted, deleted_at FROM messages WHERE id = ?", (message_id,)
            ).fetchone()
        return _message_from_row(row) if row else None

    def messages(self, ids: Iterable[str]) -> dict[str, Message]:
        """Retained messages (with content where retained), in one query."""
        rows = self._rows("summary, content, is_deleted, deleted_at", "id", ids)
        return {m.id: m for m in (_message_from_row(r) for r in rows)}

    def summaries(self, ids: Iterable[str]) -> dict[str, MessageSummary]:
        rows = self._rows("summary, NULL, is_deleted, deleted_at", "id", ids)
        return {m.id: m for m in (_summary_from_row(r) for r in rows)}

    def conversations(self, conversation_ids: Iterable[str]) -> dict[str, list[MessageSummary]]:
        """Retained summaries per conversation, in one query."""
        out: dict[str, list[MessageSummary]] = {}
        for row in self._rows("summary, NULL, is_deleted, deleted_at", "conversation_id", conversation_ids):
            summary = _summary_from_row(row)
            out.setdefault(summary.conversation_id or "", []).append(summary)
        return out

    def _rows(self, columns: str, key: str, values: Iterable[str]) -> list[tuple]:
        values = list(dict.fromkeys(values))
        rows: list[tuple] = []
        with self._tx() as db:
            for start in range(0, len(values), 500):  # well under SQLite's bound-parameter limit
                chunk = values[start : start + 500]
                rows += db.execute(
                    f"SELECT {columns} FROM messages WHERE {key} IN ({','.join('?' * len(chunk))})", chunk
                ).fetchall()
        return rows

    def window(
        self,
        *,
        folder_id: str | None,
        since: datetime | None,
        until: datetime | None,
        deleted: bool | None = None,
        limit: int | None = None,
    ) -> list[MessageSummary]:
        """Retained messages newest first. ``deleted``: None = all, True/False = only those."""
        sql = "SELECT summary, NULL, is_deleted, deleted_at FROM messages WHERE 1=1"
        args: list[object] = []
        if folder_id:
            sql += " AND folder_id = ?"
            args.append(folder_id)
        if since:
            sql += " AND received_at >= ?"
            args.append(_iso(since))
        if until:
            sql += " AND received_at <= ?"
            args.append(_iso(until))
        if deleted is not None:
            sql += " AND is_deleted = ?"
            args.append(int(deleted))
        sql += " ORDER BY received_at DESC"
        if limit:
            sql += " LIMIT ?"
            args.append(limit)
        with self._tx() as db:
            rows = db.execute(sql, args).fetchall()
        return [_summary_from_row(r) for r in rows]

    def conversation(self, conversation_id: str) -> list[MessageSummary]:
        return self.conversations([conversation_id]).get(conversation_id, [])

    def mark_deleted(self, ids: Iterable[str], when: datetime | None = None) -> None:
        stamp = _iso(when or datetime.now(UTC))
        with self._tx() as db:
            db.executemany(
                "UPDATE messages SET is_deleted = 1, deleted_at = COALESCE(deleted_at, ?) WHERE id = ?",
                [(stamp, i) for i in ids],
            )

    def set_folders(self, moves: dict[str, str]) -> None:
        with self._tx() as db:
            for message_id, folder_id in moves.items():
                row = db.execute("SELECT summary FROM messages WHERE id = ?", (message_id,)).fetchone()
                if row is None:
                    continue
                summary = json.loads(row[0]) | {"folder_id": folder_id}
                db.execute(
                    "UPDATE messages SET folder_id = ?, summary = ?, is_deleted = 0, deleted_at = NULL "
                    "WHERE id = ?",
                    (folder_id, json.dumps(summary), message_id),
                )


def _summary_from_row(row: tuple[str, str | None, int, str | None]) -> MessageSummary:
    summary = MessageSummary.model_validate_json(row[0])
    summary.is_deleted = bool(row[2])
    summary.deleted_at = datetime.fromisoformat(row[3]) if row[3] else None
    return summary


def _message_from_row(row: tuple[str, str | None, int, str | None]) -> Message:
    summary = _summary_from_row(row)
    if not row[1]:
        return Message.model_validate(summary.model_dump())
    content = Message.model_validate_json(row[1])
    return Message.model_validate(content.model_dump() | summary.model_dump())


def _merge_content(existing: str | None, new: Message) -> str:
    """Keep previously retained bodies when a newer fetch lacks them (e.g. text vs html fetch)."""
    if not existing:
        return new.model_dump_json()
    old = Message.model_validate_json(existing)
    merged = new.model_dump()
    for key in ("body_text", "unique_body_text", "body_html", "unique_body_html", "attachments"):
        if not merged.get(key) and getattr(old, key):
            merged[key] = (
                getattr(old, key) if key != "attachments" else [a.model_dump() for a in old.attachments]
            )
    return Message.model_validate(merged).model_dump_json()

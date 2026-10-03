"""Account-bound SQLite store (architecture §5.7).

Holds the account binding and the folder cache only: no messages, no bodies, no attachment bytes.
Several short-lived processes may share it: WAL mode, busy timeout, one connection per operation,
short transactions.

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
from pathlib import Path

from outlook_connector import config
from outlook_connector.domain.errors import AccountMismatch, ConnectorError
from outlook_connector.domain.models import Folder

SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS folders (
    id TEXT PRIMARY KEY,
    data TEXT NOT NULL
);
"""


log = logging.getLogger(__name__)


def store_path(fingerprint: str) -> Path:
    return config.data_dir() / "accounts" / fingerprint / "mail.sqlite3"


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

from __future__ import annotations

import contextlib
import sqlite3
from pathlib import Path

import pytest

from outlook_connector.domain.errors import AccountMismatch
from outlook_connector.domain.models import Folder
from outlook_connector.store.db import Store, store_path


@pytest.fixture
def store(tmp_path: Path) -> Store:
    return Store(tmp_path / "mail.sqlite3", owner="fp-1")


def test_path_is_per_account(isolated_home: Path) -> None:
    assert store_path("abc").parent.name == "abc"
    assert store_path("abc").is_relative_to(isolated_home.resolve())


def test_owner_is_enforced(tmp_path: Path) -> None:
    Store(tmp_path / "mail.sqlite3", owner="fp-1")
    with pytest.raises(AccountMismatch):
        Store(tmp_path / "mail.sqlite3", owner="fp-2")


def test_folder_cache_round_trip(store: Store) -> None:
    assert store.folders() == ([], None)
    store.save_folders([Folder(id="f1", name="Inbox", well_known="inbox")])
    folders, age = store.folders()
    assert [f.name for f in folders] == ["Inbox"] and age is not None and age < 5


def test_store_holds_no_messages(store: Store) -> None:
    with contextlib.closing(sqlite3.connect(store.path)) as db:
        tables = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
    assert tables == {"meta", "folders"}


def test_two_store_instances_share_the_file(tmp_path: Path) -> None:
    one, two = Store(tmp_path / "m.sqlite3", "fp"), Store(tmp_path / "m.sqlite3", "fp")
    one.save_folders([Folder(id="f1", name="Inbox")])
    assert [f.id for f in two.folders()[0]] == ["f1"]


def test_damaged_store_is_moved_aside_and_rebuilt(tmp_path: Path) -> None:
    path = tmp_path / "mail.sqlite3"
    store = Store(path, owner="fp")
    store.save_folders([Folder(id=f"f{i}", name=f"Folder {i}" * 20) for i in range(400)])
    with path.open("r+b") as handle:  # clobber a data page in the middle of the file
        handle.seek(path.stat().st_size // 2)
        handle.write(b"\x00garbage" * 512)
    fresh = Store(path, owner="fp")
    assert fresh.folders() == ([], None)
    quarantined = [p for p in tmp_path.iterdir() if p.name.startswith("corrupt-")]
    assert len(quarantined) == 1 and (quarantined[0] / "mail.sqlite3").exists()

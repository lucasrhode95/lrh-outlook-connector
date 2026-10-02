from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path

import pytest

from outlook_connector.domain.errors import AccountMismatch
from outlook_connector.domain.models import Folder, MessageSummary, Recipient
from outlook_connector.store.db import Store, store_path


def summary(mid: str, *, day: int, folder: str = "f-inbox", conv: str = "c1") -> MessageSummary:
    return MessageSummary(
        id=mid,
        conversation_id=conv,
        folder_id=folder,
        subject=f"s-{mid}",
        received_at=datetime(2026, 9, day, 12, tzinfo=UTC),
        sender=Recipient(name="A", address="a@example.com"),
    )


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


def test_window_queries_newest_first(store: Store) -> None:
    store.upsert_summaries([summary("a", day=1), summary("b", day=3), summary("c", day=2, folder="f-sent")])
    assert [m.id for m in store.window(folder_id=None, since=None, until=None)] == ["b", "c", "a"]
    assert [
        m.id for m in store.window(folder_id="f-inbox", since=datetime(2026, 9, 2, tzinfo=UTC), until=None)
    ] == ["b"]


def test_forget_drops_rows_of_a_covered_window_except_kept(store: Store) -> None:
    store.upsert_summaries(
        [summary("a", day=1), summary("b", day=2), summary("c", day=3), summary("s", day=2, folder="f-sent")]
    )
    store.forget(folder_id="f-inbox", since=datetime(2026, 9, 2, tzinfo=UTC), until=None, keep={"c"})
    assert [m.id for m in store.window(folder_id=None, since=None, until=None)] == ["c", "s", "a"]


def test_summaries_lookup_holds_only_summary_fields(store: Store) -> None:
    labelled = summary("a", day=1).model_copy(update={"folder": "Inbox", "also_in": ["Sent Items"]})
    store.upsert_summaries([labelled, summary("b", day=2, conv="c2")])
    assert set(store.summaries(["a", "b", "zzz"])) == {"a", "b"}
    cached = store.summaries(["a"])["a"]
    assert cached.folder is None and cached.also_in == []  # derived per result, never stored


def test_two_store_instances_share_the_file(tmp_path: Path) -> None:
    one, two = Store(tmp_path / "m.sqlite3", "fp"), Store(tmp_path / "m.sqlite3", "fp")
    one.upsert_summaries([summary("a", day=1)])
    assert [m.id for m in two.window(folder_id=None, since=None, until=None)] == ["a"]


def test_damaged_store_is_moved_aside_and_rebuilt(tmp_path: Path) -> None:
    path = tmp_path / "mail.sqlite3"
    store = Store(path, owner="fp")
    store.upsert_summaries([summary(f"m{i}", day=1 + i % 28) for i in range(400)])
    with path.open("r+b") as handle:  # clobber a data page in the middle of the file
        handle.seek(path.stat().st_size // 2)
        handle.write(b"\x00garbage" * 512)
    fresh = Store(path, owner="fp")
    assert fresh.window(folder_id=None, since=None, until=None) == []
    quarantined = [p for p in tmp_path.iterdir() if p.name.startswith("corrupt-")]
    assert len(quarantined) == 1 and (quarantined[0] / "mail.sqlite3").exists()

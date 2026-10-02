from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path

import pytest

from outlook_connector.domain.errors import AccountMismatch
from outlook_connector.domain.models import Folder, Message, MessageSummary, Recipient
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


def test_deleted_rows_keep_content_and_reappearing_clears_the_marker(store: Store) -> None:
    msg = Message(**summary("a", day=1).model_dump(), body_text="full", unique_body_text="unique")
    store.save_messages([msg])
    store.mark_deleted(["a"])
    retained = store.message("a")
    assert retained and retained.is_deleted and retained.deleted_at and retained.body_text == "full"
    store.upsert_summaries([summary("a", day=1)])  # seen on the server again
    again = store.message("a")
    assert again and not again.is_deleted and again.body_text == "full"


def test_later_fetch_without_bodies_does_not_erase_retained_bodies(store: Store) -> None:
    store.save_messages([Message(**summary("a", day=1).model_dump(), body_text="text")])
    store.save_messages([Message(**summary("a", day=1).model_dump(), body_html="<p>html</p>")])
    retained = store.message("a")
    assert retained and retained.body_text == "text" and retained.body_html == "<p>html</p>"


def test_moves_update_folder_and_summary(store: Store) -> None:
    store.upsert_summaries([summary("a", day=1)])
    store.mark_deleted(["a"])
    store.set_folders({"a": "f-archive"})
    moved = store.window(folder_id="f-archive", since=None, until=None)
    assert [m.id for m in moved] == ["a"] and not moved[0].is_deleted and moved[0].folder_id == "f-archive"


def test_conversation_and_summaries_lookup(store: Store) -> None:
    store.upsert_summaries([summary("a", day=1), summary("b", day=2, conv="c2")])
    assert [m.id for m in store.conversation("c1")] == ["a"]
    assert set(store.summaries(["a", "b", "zzz"])) == {"a", "b"}


def test_two_store_instances_share_the_file(tmp_path: Path) -> None:
    one, two = Store(tmp_path / "m.sqlite3", "fp"), Store(tmp_path / "m.sqlite3", "fp")
    one.upsert_summaries([summary("a", day=1)])
    assert [m.id for m in two.window(folder_id=None, since=None, until=None)] == ["a"]

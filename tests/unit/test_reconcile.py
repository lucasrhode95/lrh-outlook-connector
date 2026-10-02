from __future__ import annotations

from pathlib import Path

import httpx

from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.service.reconcile import Reconciler
from outlook_connector.store.db import Store
from tests.fakes.graph_fake import StaticTokens, sample_mailbox


async def test_deleted_and_moved_messages_are_reconciled(tmp_path: Path) -> None:
    fake = sample_mailbox()
    reader = GraphMailReader(
        Graph(Transport(StaticTokens(), client=httpx.AsyncClient(transport=fake.transport())))
    )
    store = Store(tmp_path / "m.sqlite3", "fp")
    inbox, _ = await reader.list_messages(
        folder_id="f-inbox", since=None, until=None, page_size=50, page=None
    )
    store.upsert_summaries(inbox)  # m5, m1

    del fake.messages["m5"]  # deleted for good on the server
    fake.messages["m1"].folder = "f-archive"  # moved by the user
    remote, _ = await reader.list_messages(
        folder_id="f-inbox", since=None, until=None, page_size=50, page=None
    )
    changed = await Reconciler(reader, store).after_window(
        folder_id="f-inbox", since=None, until=None, remote=remote, complete=True
    )
    assert changed == 2
    retained = store.message("m5")
    assert retained and retained.is_deleted
    assert [m.id for m in store.window(folder_id="f-archive", since=None, until=None)] == ["m1"]


async def test_incomplete_listing_only_checks_the_covered_span(tmp_path: Path) -> None:
    fake = sample_mailbox()
    reader = GraphMailReader(
        Graph(Transport(StaticTokens(), client=httpx.AsyncClient(transport=fake.transport())))
    )
    store = Store(tmp_path / "m.sqlite3", "fp")
    everything, _ = await reader.list_messages(
        folder_id=None, since=None, until=None, page_size=50, page=None
    )
    store.upsert_summaries(everything)
    newest, _ = await reader.list_messages(folder_id=None, since=None, until=None, page_size=1, page=None)
    calls_before = len(fake.calls)
    changed = await Reconciler(reader, store).after_window(
        folder_id=None, since=None, until=None, remote=newest, complete=False
    )
    assert changed == 0 and len(fake.calls) == calls_before  # older rows were not outside the span

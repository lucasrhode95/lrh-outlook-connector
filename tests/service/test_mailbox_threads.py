from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path

import httpx
import pytest

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.threads import Threads, base_subject
from outlook_connector.store.db import Store
from tests.fakes.graph_fake import FakeGraph, FakeMessage, StaticTokens, sample_mailbox


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


@pytest.fixture
def mailbox(fake: FakeGraph, tmp_path: Path) -> Mailbox:
    transport = Transport(StaticTokens(), client=httpx.AsyncClient(transport=fake.transport()))
    return Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "m.sqlite3", "fp"))


# ---------------------------------------------------------------- folders


async def test_folders_have_paths_and_are_cached(mailbox: Mailbox, fake: FakeGraph) -> None:
    folders = {f.id: f for f in await mailbox.folders()}
    assert folders["f-rie"].path == "Inbox/Projects/RIE"
    calls = len(fake.calls)
    await mailbox.folders()
    assert len(fake.calls) == calls  # served from the folder cache


@pytest.mark.parametrize("ref", ["inbox", "INBOX", "Inbox", "f-inbox"])
async def test_resolve_folder_by_alias_path_name_or_id(mailbox: Mailbox, ref: str) -> None:
    assert (await mailbox.resolve_folder(ref)).id == "f-inbox"


async def test_resolve_folder_path_and_unknown(mailbox: Mailbox) -> None:
    assert (await mailbox.resolve_folder("inbox/projects/rie")).id == "f-rie"
    with pytest.raises(InvalidRequest, match="Unknown folder"):
        await mailbox.resolve_folder("Nope")


# ---------------------------------------------------------------- listing


async def test_list_messages_pages_with_self_contained_cursor(mailbox: Mailbox) -> None:
    first = await mailbox.list_messages(limit=2)
    assert [m.id for m in first.items] == ["m5", "m4"]
    assert first.cursor and first.coverage.more_available and not first.coverage.complete
    assert first.items[0].folder == "Inbox"
    second = await mailbox.list_messages(limit=2, cursor=first.cursor)
    assert [m.id for m in second.items] == ["m3", "m2"]


async def test_list_messages_folder_and_window(mailbox: Mailbox) -> None:
    page = await mailbox.list_messages(folder="inbox", since=datetime(2026, 9, 29, tzinfo=UTC))
    assert [m.id for m in page.items] == ["m5"] and page.coverage.complete


async def test_server_deleted_messages_are_retained_and_labelled(mailbox: Mailbox, fake: FakeGraph) -> None:
    await mailbox.list_messages(folder="inbox")
    del fake.messages["m1"]
    page = await mailbox.list_messages(folder="inbox")
    deleted = [m for m in page.items if m.is_deleted]
    assert [m.id for m in deleted] == ["m1"]
    assert page.coverage.source == "remote+local" and "deleted on the server" in page.coverage.notes[0]


async def test_received_only_leaves_out_sent_deleted_and_junk(mailbox: Mailbox) -> None:
    page = await mailbox.list_messages(received_only=True)
    assert [m.id for m in page.items] == ["m5", "m3", "m1"]
    assert any("received_only" in n for n in page.coverage.notes)


async def test_received_only_scans_past_filtered_pages_and_keeps_it_in_the_cursor(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    for i in range(3):
        fake.add(FakeMessage(f"j{i}", "spam", "f-junk", f"2026-10-01T0{i}:00:00Z", conversation=f"cj{i}"))
    first = await mailbox.list_messages(received_only=True, limit=2)
    assert [m.id for m in first.items] == ["m5"] and first.cursor  # pages of junk skipped
    second = await mailbox.list_messages(limit=2, cursor=first.cursor)
    assert [m.id for m in second.items] == ["m3"]  # m4 (junk) and m2 (sent) left out


async def test_local_only_listing_says_so(mailbox: Mailbox) -> None:
    await mailbox.list_messages()
    page = await mailbox.list_messages(refresh=False)
    assert page.coverage.source == "local" and not page.coverage.complete and len(page.items) == 5


async def test_list_messages_validation(mailbox: Mailbox) -> None:
    with pytest.raises(InvalidRequest):
        await mailbox.list_messages(limit=0)
    with pytest.raises(InvalidRequest):
        await mailbox.list_messages(cursor="garbage!")


# ---------------------------------------------------------------- content


async def test_get_message_bounded_with_continuation(mailbox: Mailbox) -> None:
    first = await mailbox.get_message("m1", body="full", max_chars=5)
    assert (
        first.text == "First"
        and first.next_offset == 5
        and first.total_chars == len("First report\n\nregards")
    )
    rest = await mailbox.get_message("m1", body="full", offset=5, max_chars=1000)
    assert rest.next_offset is None and rest.text.endswith("regards")


async def test_unique_body_is_the_default(mailbox: Mailbox) -> None:
    assert (await mailbox.get_message("m2")).text == "Thanks!"


async def test_deleted_message_falls_back_to_retained_copy(mailbox: Mailbox, fake: FakeGraph) -> None:
    await mailbox.get_message("m2")
    del fake.messages["m2"]
    content = await mailbox.get_message("m2")
    assert content.message.is_deleted and content.text == "Thanks!"
    del fake.messages["m5"]
    with pytest.raises(NotFound):
        await mailbox.get_message("m5")


# ---------------------------------------------------------------- search


async def test_search_groups_by_conversation_with_coverage(mailbox: Mailbox) -> None:
    result = await mailbox.search("relatório")
    assert [h.conversation_id for h in result.conversations] == ["c-rel"]
    assert {m.id for m in result.conversations[0].matching_messages} == {"m1", "m2", "m3"}
    assert result.coverage.server_total is None and result.coverage.complete  # no total for a complete set
    assert any("not searched" in n for n in result.coverage.notes)


async def test_search_requires_a_query(mailbox: Mailbox) -> None:
    with pytest.raises(InvalidRequest):
        await mailbox.search("  ")


# ---------------------------------------------------------------- threads


async def test_thread_spans_folders_sorted_and_excludes_junk_by_default(mailbox: Mailbox) -> None:
    thread = await Threads(mailbox).get_thread("c-rel")
    assert [t.message.id for t in thread.messages] == ["m1", "m2", "m3"]
    assert [t.message.folder for t in thread.messages] == ["Inbox", "Sent Items", "Inbox/Projects/RIE"]
    assert [t.text for t in thread.messages] == ["First report", "Thanks!", "Follow-up with numbers"]
    assert thread.subject == "Relatório BE semanal"
    assert any("omitted" in n for n in thread.coverage.notes)


async def test_thread_can_include_deleted_items_and_junk(mailbox: Mailbox) -> None:
    thread = await Threads(mailbox).get_thread("c-rel", include_deleted_items=True, include_bodies=False)
    assert [t.message.id for t in thread.messages] == ["m1", "m2", "m3", "m4"]


async def test_thread_bodies_are_bounded_with_cursor(mailbox: Mailbox) -> None:
    threads = Threads(mailbox)
    first = await threads.get_thread("c-rel", max_chars=15)  # "First report" fits, "Thanks!" does not
    assert [t.message.id for t in first.messages] == ["m1"] and first.cursor
    # the cursor restores the original options: max_chars=1000 here is ignored
    second = await threads.get_thread("c-rel", max_chars=1000, cursor=first.cursor)
    assert [t.message.id for t in second.messages] == ["m2"] and second.cursor
    third = await threads.get_thread("c-rel", cursor=second.cursor)
    assert [t.message.id for t in third.messages] == ["m3"] and third.cursor is None


async def test_thread_cursor_keeps_include_deleted_items(mailbox: Mailbox) -> None:
    threads = Threads(mailbox)
    first = await threads.get_thread("c-rel", include_deleted_items=True, max_chars=15)
    rest = []
    cursor = first.cursor
    while cursor:  # continue without repeating include_deleted_items
        page = await threads.get_thread("c-rel", cursor=cursor)
        rest += [t.message.id for t in page.messages]
        cursor = page.cursor
    assert [t.message.id for t in first.messages] + rest == ["m1", "m2", "m3", "m4"]


async def test_thread_cursor_belongs_to_its_conversation(mailbox: Mailbox) -> None:
    first = await Threads(mailbox).get_thread("c-rel", max_chars=15)
    assert first.cursor
    with pytest.raises(InvalidRequest, match="different conversation"):
        await Threads(mailbox).get_thread("c-lunch", cursor=first.cursor)


async def test_truncated_conversation_is_not_reported_complete(
    mailbox: Mailbox, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("outlook_connector.remote.graph_mail.MAX_CONVERSATION", 2)
    thread = await Threads(mailbox).get_thread("c-rel", include_bodies=False)
    assert not thread.coverage.complete and any("listing limit" in n for n in thread.coverage.notes)


async def test_thread_sizes_count_like_get_thread(mailbox: Mailbox, fake: FakeGraph) -> None:
    threads = Threads(mailbox)
    sizes = {s.conversation_id: s.messages for s in await threads.sizes(["c-rel", "c-lunch"])}
    assert sizes == {"c-rel": 3, "c-lunch": 1}  # junk m4 left out, as in get_thread
    with_junk = await threads.sizes(["c-rel"], include_deleted_items=True)
    assert with_junk[0].messages == 4 and not with_junk[0].at_least
    await threads.get_thread("c-rel")
    del fake.messages["m2"]
    await threads.get_thread("c-rel")  # m2 is now retained as deleted on the server
    assert (await threads.sizes(["c-rel"]))[0].messages == 3


async def test_single_oversized_message_is_truncated_not_skipped(mailbox: Mailbox, fake: FakeGraph) -> None:
    fake.add(
        FakeMessage("big", "Huge", "f-inbox", "2026-09-01T00:00:00Z", conversation="c-big", text="x" * 500)
    )
    thread = await Threads(mailbox).get_thread("c-big", max_chars=100)
    assert thread.messages[0].truncated and len(thread.messages[0].text or "") == 100


async def test_thread_includes_server_deleted_retained_messages(mailbox: Mailbox, fake: FakeGraph) -> None:
    threads = Threads(mailbox)
    await threads.get_thread("c-rel")
    del fake.messages["m2"]
    thread = await threads.get_thread("c-rel")
    m2 = next(t for t in thread.messages if t.message.id == "m2")
    assert m2.message.is_deleted and m2.text == "Thanks!"


async def test_unknown_conversation(mailbox: Mailbox) -> None:
    with pytest.raises(NotFound):
        await Threads(mailbox).get_thread("nope")


def test_base_subject_strips_reply_and_forward_prefixes() -> None:
    assert base_subject("RE: FW: Enc: Relatório") == "Relatório"


async def test_search_reports_an_approximate_total_only_when_incomplete(mailbox: Mailbox) -> None:
    result = await mailbox.search("relatório", limit=1)
    assert result.cursor and result.coverage.server_total == 3
    assert any("approximate" in n for n in result.coverage.notes)


async def test_stale_folder_cache_is_served_immediately_and_refreshed_in_background(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    import asyncio

    from outlook_connector.service import mailbox as mailbox_module

    await mailbox.folders()  # fills the cache
    fake.add_folder("f-new", "New folder")
    mailbox_module.FOLDER_TTL_SECONDS, saved = -1, mailbox_module.FOLDER_TTL_SECONDS  # everything is stale
    try:
        served = await mailbox.folders()
        assert "f-new" not in {f.id for f in served}  # answered from the cache, without waiting
        assert mailbox._folder_refresh is not None
        await asyncio.wait_for(asyncio.shield(mailbox._folder_refresh), timeout=5)
    finally:
        mailbox_module.FOLDER_TTL_SECONDS = saved
    assert "f-new" in {f.id for f in (await mailbox.folders())}

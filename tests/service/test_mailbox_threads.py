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
    assert [m.id for m in first.items] == ["m5"]  # m4 is in Junk Email, left out by default
    assert first.coverage.excluded == {"deleted_or_junk": 1}
    assert first.cursor and not first.coverage.complete
    assert first.items[0].folder == "Inbox"
    second = await mailbox.list_messages(limit=2, cursor=first.cursor)
    assert [m.id for m in second.items] == ["m3", "m2"]


async def test_list_messages_folder_and_window(mailbox: Mailbox) -> None:
    page = await mailbox.list_messages(folder="inbox", since=datetime(2026, 9, 29, tzinfo=UTC))
    assert [m.id for m in page.items] == ["m5"] and page.coverage.complete


async def test_messages_deleted_on_the_server_are_gone_from_the_next_listing(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    await mailbox.list_messages(folder="inbox")
    del fake.messages["m1"]
    page = await mailbox.list_messages(folder="inbox")
    assert "m1" not in [m.id for m in page.items]


async def test_received_only_leaves_out_sent_deleted_and_junk(mailbox: Mailbox) -> None:
    page = await mailbox.list_messages(received_only=True)
    assert [m.id for m in page.items] == ["m5", "m3", "m1"]
    assert page.coverage.excluded == {"deleted_or_junk": 1, "outgoing": 1}


async def test_deleted_items_and_junk_are_left_out_unless_asked_or_named(mailbox: Mailbox) -> None:
    assert "m4" not in [m.id for m in (await mailbox.list_messages()).items]
    assert "m4" in [m.id for m in (await mailbox.list_messages(include_deleted_items=True)).items]
    named = await mailbox.list_messages(folder="junkemail")  # a folder asked for by name is listed
    assert [m.id for m in named.items] == ["m4"] and not named.coverage.excluded


async def test_list_total_counts_the_server_scope(mailbox: Mailbox) -> None:
    page = await mailbox.list_messages(include_total=True)
    assert page.coverage.server_total == 4  # five messages minus the one in Junk Email


async def test_copies_of_one_message_are_shown_once(mailbox: Mailbox, fake: FakeGraph) -> None:
    for mid, folder in (("self-sent", "f-sent"), ("self-recv", "f-inbox")):
        fake.add(
            FakeMessage(mid, "Note to self", folder, "2026-10-01T09:00:00Z", conversation="c-self",
                        internet_id="<self@example.com>")
        )  # fmt: skip
    page = await mailbox.list_messages()
    copy = next(m for m in page.items if m.conversation_id == "c-self")
    assert [m.id for m in page.items].count(copy.id) == 1 and copy.id == "self-recv"
    assert copy.also_in == ["Sent Items"]
    thread = await Threads(mailbox).get_thread("c-self")
    assert [t.message.id for t in thread.messages] == ["self-recv"]
    assert (await mailbox.conversation_sizes(["c-self"]))[0].messages == 1


async def test_received_only_scans_past_filtered_pages_and_keeps_it_in_the_cursor(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    for i in range(3):
        fake.add(FakeMessage(f"j{i}", "spam", "f-junk", f"2026-10-01T0{i}:00:00Z", conversation=f"cj{i}"))
    first = await mailbox.list_messages(received_only=True, limit=2)
    assert [m.id for m in first.items] == ["m5"] and first.cursor  # pages of junk skipped
    second = await mailbox.list_messages(limit=2, cursor=first.cursor)
    assert [m.id for m in second.items] == ["m3"]  # m4 (junk) and m2 (sent) left out


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


async def test_message_deleted_on_the_server_is_not_found(mailbox: Mailbox, fake: FakeGraph) -> None:
    await mailbox.get_message("m2")
    del fake.messages["m2"]
    with pytest.raises(NotFound, match="deleted on the server"):
        await mailbox.get_message("m2")
    with pytest.raises(NotFound):
        await mailbox.attachments("m2")


async def test_search_groups_by_conversation_with_coverage(mailbox: Mailbox) -> None:
    result = await mailbox.search("relatório")
    assert [h.conversation_id for h in result.conversations] == ["c-rel"]
    assert {m.id for m in result.conversations[0].matching_messages} == {"m1", "m2", "m3"}
    assert result.coverage.server_total is None and result.coverage.complete


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
    assert any("left out: in Deleted Items or Junk Email" in n for n in thread.coverage.notes)
    assert thread.coverage.excluded == {"deleted_or_junk": 1}


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
    sizes = {s.conversation_id: s.messages for s in await mailbox.conversation_sizes(["c-rel", "c-lunch"])}
    assert sizes == {"c-rel": 3, "c-lunch": 1}  # junk m4 left out, as in get_thread
    with_junk = await mailbox.conversation_sizes(["c-rel"], include_deleted_items=True)
    assert with_junk[0].messages == 4 and not with_junk[0].at_least
    del fake.messages["m2"]
    assert (await mailbox.conversation_sizes(["c-rel"]))[0].messages == 2
    assert len((await threads.get_thread("c-rel")).messages) == 2


async def test_single_oversized_message_is_truncated_not_skipped(mailbox: Mailbox, fake: FakeGraph) -> None:
    fake.add(
        FakeMessage("big", "Huge", "f-inbox", "2026-09-01T00:00:00Z", conversation="c-big", text="x" * 500)
    )
    thread = await Threads(mailbox).get_thread("c-big", max_chars=100)
    assert thread.messages[0].truncated and len(thread.messages[0].text or "") == 100


async def test_unknown_conversation(mailbox: Mailbox) -> None:
    with pytest.raises(NotFound):
        await Threads(mailbox).get_thread("nope")


def test_base_subject_strips_reply_and_forward_prefixes() -> None:
    assert base_subject("RE: FW: Enc: Relatório") == "Relatório"


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


async def test_thread_coverage_ignores_a_body_cut_to_fit(mailbox: Mailbox, fake: FakeGraph) -> None:
    fake.add(
        FakeMessage("big", "Huge", "f-inbox", "2026-09-01T00:00:00Z", conversation="c-big", text="x" * 5000)
    )
    thread = await Threads(mailbox).get_thread("c-big", max_chars=1000)
    assert thread.messages[0].truncated and thread.cursor is None and thread.coverage.complete


async def test_truncated_listing_stays_incomplete_when_bodies_fit(
    mailbox: Mailbox, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("outlook_connector.remote.graph_mail.MAX_CONVERSATION", 2)
    thread = await Threads(mailbox).get_thread("c-rel", max_chars=100_000)
    assert not thread.coverage.complete


async def test_search_dates_are_exact_whatever_the_time_zone(mailbox: Mailbox) -> None:
    # m1 Sep 28 09:00, m2 Sep 28 10:00, m3 Sep 29 08:00 (UTC); KQL only knows dates
    result = await mailbox.search(
        "relatório",
        since=datetime(2026, 9, 28, 9, 30, tzinfo=UTC),
        until=datetime(2026, 9, 29, 7, tzinfo=UTC),
    )
    assert [m.id for hit in result.conversations for m in hit.matching_messages] == ["m2"]


async def test_search_hits_carry_the_conversation_size(mailbox: Mailbox) -> None:
    result = await mailbox.search("relatório")
    assert result.conversations[0].message_count == 3  # m4 in Junk is not counted by default


async def test_unknown_folder_name_refreshes_the_folder_list(mailbox: Mailbox, fake: FakeGraph) -> None:
    await mailbox.folders()
    fake.add_folder("f-new", "Brand new", parent="f-inbox")
    assert (await mailbox.resolve_folder("Inbox/Brand new")).id == "f-new"


async def test_copies_split_across_pages_are_returned_once(mailbox: Mailbox, fake: FakeGraph) -> None:
    fake.add(
        FakeMessage(
            "cp-in", "Copy", "f-inbox", "2026-10-01T09:00:01Z", conversation="c-cp", internet_id="<cp@x>"
        )
    )
    fake.add(
        FakeMessage(
            "cp-out", "Copy", "f-sent", "2026-10-01T09:00:00Z", conversation="c-cp", internet_id="<cp@x>"
        )
    )
    seen, cursor = [], None
    while True:
        page = await mailbox.list_messages(limit=1, cursor=cursor)
        seen += [m.id for m in page.items]
        if not (cursor := page.cursor):
            break
    assert seen.count("cp-in") + seen.count("cp-out") == 1


# ---------------------------------------------------------------- reach: hidden folders, Sync Issues


def add_out_of_reach_mail(fake: FakeGraph) -> None:
    """A hidden folder, an item outside the mail folders, Outlook's Sync Issues and a folder that was
    deleted in Outlook (it moved into Deleted Items), each holding one message about 'budget'."""
    fake.add_folder("f-hidden", "TeamsMeetings", hidden=True)
    fake.add_folder("f-sync", "Sync Issues", alias="syncissues", hidden=True)
    fake.add_folder("f-conflicts", "Conflicts", parent="f-sync", alias="conflicts")
    fake.add_folder("f-old", "Old project", parent="f-deleted")
    for mid, folder in (("h1", "f-hidden"), ("o1", "f-outside"), ("c1", "f-conflicts"), ("d1", "f-old")):
        fake.add(FakeMessage(mid, f"budget {mid}", folder, "2026-10-01T09:00:00Z", conversation=f"c-{mid}"))


async def test_hidden_folders_and_items_outside_the_mail_folders_are_out_of_reach(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    add_out_of_reach_mail(fake)
    for include in (False, True):  # include_deleted_items never brings them back
        page = await mailbox.list_messages(include_deleted_items=include)
        assert not {"h1", "o1"} & {m.id for m in page.items}
        assert page.coverage.excluded["hidden"] == 2
        found = await mailbox.search("budget", include_deleted_items=include)
        assert not {"h1", "o1"} & {m.id for hit in found.conversations for m in hit.matching_messages}
    assert "f-hidden" not in {f.id for f in await mailbox.folders()}
    with pytest.raises(InvalidRequest, match="hidden folder"):
        await mailbox.resolve_folder("TeamsMeetings")
    assert (await mailbox.conversation_sizes(["c-h1", "c-o1"]))[0].messages == 0


async def test_sync_issues_are_left_out_like_deleted_items_unless_asked(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    add_out_of_reach_mail(fake)
    page = await mailbox.list_messages()
    assert "c1" not in {m.id for m in page.items} and page.coverage.excluded["sync_issues"] == 1
    assert "c1" in {m.id for m in (await mailbox.list_messages(include_deleted_items=True)).items}
    # listed and reachable by name, even though Graph marks the Sync Issues folder hidden
    assert {"f-sync", "f-conflicts"} <= {f.id for f in await mailbox.folders()}
    named = await mailbox.list_messages(folder="Sync Issues/Conflicts")
    assert [m.id for m in named.items] == ["c1"]


async def test_a_folder_deleted_in_outlook_counts_as_deleted_items(mailbox: Mailbox, fake: FakeGraph) -> None:
    add_out_of_reach_mail(fake)
    page = await mailbox.list_messages()
    assert "d1" not in {m.id for m in page.items}
    assert "d1" in {m.id for m in (await mailbox.list_messages(include_deleted_items=True)).items}


async def test_total_counts_only_reachable_folders_in_scope(mailbox: Mailbox, fake: FakeGraph) -> None:
    add_out_of_reach_mail(fake)
    page = await mailbox.list_messages(include_total=True)
    assert page.coverage.server_total == 4  # m1, m2, m3, m5: not Junk, Sync Issues, hidden or deleted
    with_deleted = await mailbox.list_messages(include_total=True, include_deleted_items=True)
    assert with_deleted.coverage.server_total == 7  # + m4 (Junk), c1 (Sync Issues), d1 (deleted folder)


async def test_mail_in_a_folder_created_meanwhile_is_found(mailbox: Mailbox, fake: FakeGraph) -> None:
    await mailbox.folders()  # the folder list is now cached
    fake.add_folder("f-late", "Created later", parent="f-inbox")
    fake.add(FakeMessage("late1", "New folder mail", "f-late", "2026-10-01T10:00:00Z", conversation="c-late"))
    page = await mailbox.list_messages()
    assert "late1" in {m.id for m in page.items} and "hidden" not in page.coverage.excluded


async def test_items_outside_the_mail_folders_refresh_the_folder_list_once(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    fake.add(FakeMessage("o1", "budget", "f-outside", "2026-10-01T09:00:00Z", conversation="c-o1"))
    await mailbox.list_messages()
    walks = fake.calls.count("GET /v1.0/me/mailFolders")
    await mailbox.list_messages()
    await mailbox.search("budget")
    assert fake.calls.count("GET /v1.0/me/mailFolders") == walks  # remembered as outside: no new refresh

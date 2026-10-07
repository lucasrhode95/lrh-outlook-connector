from __future__ import annotations

import sqlite3
import time
from datetime import UTC, datetime, timedelta, timezone
from pathlib import Path

import httpx
import pytest

from outlook_connector.domain.errors import InvalidRequest, NotFound
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.service import cursors
from outlook_connector.service.conversations import Conversations, base_subject
from outlook_connector.service.mailbox import Mailbox
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


async def test_meeting_mail_is_marked(mailbox: Mailbox, fake: FakeGraph) -> None:
    when = {"dateTime": "2026-10-07T18:00:00.0000000", "timeZone": "UTC"}
    fake.add(
        FakeMessage("inv", "CCB", "f-inbox", "2026-10-02T09:00:00Z", conversation="c-ccb",
                    meeting={"meetingMessageType": "meetingRequest", "meetingRequestType": "fullUpdate",
                             "startDateTime": when, "endDateTime": when, "location": {"displayName": "Teams"},
                             "isAllDay": False, "isOutOfDate": True})
    )  # fmt: skip
    fake.add(
        FakeMessage("acc", "Accepted: CCB", "f-inbox", "2026-10-02T10:00:00Z", conversation="c-ccb",
                    meeting={"meetingMessageType": "meetingTenativelyAccepted"})
    )  # fmt: skip
    items = {m.id: m for m in (await mailbox.list_messages(folder="inbox")).items}
    invite = items["inv"].meeting
    assert invite and invite.kind == "update" and invite.location == "Teams" and invite.out_of_date
    assert invite.start == datetime(2026, 10, 7, 18, tzinfo=UTC)
    assert items["acc"].meeting and items["acc"].meeting.kind == "tentative"
    assert items["m5"].meeting is None  # ordinary mail


def add_meeting_conversations(fake: FakeGraph) -> None:
    """c-only: an invitation, an RSVP and a cancellation. c-talk: an invitation with a real reply."""
    for mid, kind, conv in (
        ("inv1", "meetingRequest", "c-only"),
        ("rsvp1", "meetingAccepted", "c-only"),
        ("cxl1", "meetingCancelled", "c-only"),
        ("inv2", "meetingRequest", "c-talk"),
    ):
        fake.add(FakeMessage(mid, "Sync", "f-inbox", "2026-10-02T09:00:00Z", conversation=conv,
                             meeting={"meetingMessageType": kind}))  # fmt: skip
    fake.add(FakeMessage("reply2", "RE: Sync", "f-inbox", "2026-10-02T10:00:00Z", conversation="c-talk"))


async def test_meeting_mail_can_be_left_out(mailbox: Mailbox, fake: FakeGraph) -> None:
    add_meeting_conversations(fake)
    everything = await mailbox.list_messages(folder="inbox")
    assert {"inv1", "rsvp1", "cxl1", "inv2", "reply2"} <= {m.id for m in everything.items}  # default: shown
    page = await mailbox.list_messages(folder="inbox", include_meeting_mail=False)
    ids = {m.id for m in page.items}
    assert "reply2" in ids  # the conversation with a real reply still shows, through the reply
    assert not ids & {"inv1", "rsvp1", "cxl1", "inv2"}  # the meeting-only conversation is gone
    assert page.coverage.excluded == {"meeting_mail": 4}
    conversation = await Conversations(mailbox).get_conversation("c-talk", include_bodies=False)
    assert [t.message.id for t in conversation.messages] == ["inv2", "reply2"]  # conversations stay whole


async def test_meeting_filter_survives_paging_and_search(mailbox: Mailbox, fake: FakeGraph) -> None:
    add_meeting_conversations(fake)
    first = await mailbox.list_messages(include_meeting_mail=False, limit=2)
    rest, cursor = list(first.items), first.cursor
    while cursor:
        page = await mailbox.list_messages(limit=2, cursor=cursor)
        rest += page.items
        cursor = page.cursor
    assert not any(m.meeting for m in rest) and "reply2" in {m.id for m in rest}
    result = await mailbox.search("Sync", include_meeting_mail=False)
    assert [m.id for h in result.conversations for m in h.matching_messages] == ["reply2"]


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


async def test_without_sent_items_leaves_out_sent_deleted_and_junk(mailbox: Mailbox) -> None:
    page = await mailbox.list_messages(include_sent_items=False)
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
    conversation = await Conversations(mailbox).get_conversation("c-self")
    assert [t.message.id for t in conversation.messages] == ["self-recv"]
    assert (await mailbox.conversation_sizes(["c-self"]))[0].messages == 1


async def test_without_sent_items_scans_past_filtered_pages_and_keeps_it_in_the_cursor(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    for i in range(3):
        fake.add(FakeMessage(f"j{i}", "spam", "f-junk", f"2026-10-01T0{i}:00:00Z", conversation=f"cj{i}"))
    first = await mailbox.list_messages(include_sent_items=False, limit=2)
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


# ---------------------------------------------------------------- conversations


async def test_conversation_spans_folders_sorted_and_excludes_junk_by_default(mailbox: Mailbox) -> None:
    conversation = await Conversations(mailbox).get_conversation("c-rel")
    assert [t.message.id for t in conversation.messages] == ["m1", "m2", "m3"]
    assert [t.message.folder for t in conversation.messages] == ["Inbox", "Sent Items", "Inbox/Projects/RIE"]
    assert [t.text for t in conversation.messages] == ["First report", "Thanks!", "Follow-up with numbers"]
    assert conversation.subject == "Relatório BE semanal"
    assert any("left out: in Deleted Items or Junk Email" in n for n in conversation.coverage.notes)
    assert conversation.coverage.excluded == {"deleted_or_junk": 1}


async def test_conversation_can_include_deleted_items_and_junk(mailbox: Mailbox) -> None:
    conversation = await Conversations(mailbox).get_conversation(
        "c-rel", include_deleted_items=True, include_bodies=False
    )
    assert [t.message.id for t in conversation.messages] == ["m1", "m2", "m3", "m4"]


async def test_conversation_bodies_are_bounded_with_cursor(mailbox: Mailbox) -> None:
    conversations = Conversations(mailbox)
    first = await conversations.get_conversation(
        "c-rel", max_chars=15
    )  # "First report" fits, "Thanks!" does not
    assert [t.message.id for t in first.messages] == ["m1"] and first.cursor
    # the cursor restores the original options: max_chars=1000 here is ignored
    second = await conversations.get_conversation("c-rel", max_chars=1000, cursor=first.cursor)
    assert [t.message.id for t in second.messages] == ["m2"] and second.cursor
    third = await conversations.get_conversation("c-rel", cursor=second.cursor)
    assert [t.message.id for t in third.messages] == ["m3"] and third.cursor is None


async def test_conversation_cursor_keeps_include_deleted_items(mailbox: Mailbox) -> None:
    conversations = Conversations(mailbox)
    first = await conversations.get_conversation("c-rel", include_deleted_items=True, max_chars=15)
    rest = []
    cursor = first.cursor
    while cursor:  # continue without repeating include_deleted_items
        page = await conversations.get_conversation("c-rel", cursor=cursor)
        rest += [t.message.id for t in page.messages]
        cursor = page.cursor
    assert [t.message.id for t in first.messages] + rest == ["m1", "m2", "m3", "m4"]


async def test_conversation_cursor_belongs_to_its_conversation(mailbox: Mailbox) -> None:
    first = await Conversations(mailbox).get_conversation("c-rel", max_chars=15)
    assert first.cursor
    with pytest.raises(InvalidRequest, match="different conversation"):
        await Conversations(mailbox).get_conversation("c-lunch", cursor=first.cursor)


async def test_truncated_conversation_is_not_reported_complete(
    mailbox: Mailbox, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("outlook_connector.remote.graph_mail.MAX_CONVERSATION", 2)
    conversation = await Conversations(mailbox).get_conversation("c-rel", include_bodies=False)
    assert not conversation.coverage.complete and any(
        "listing limit" in n for n in conversation.coverage.notes
    )


async def test_conversation_sizes_count_like_get_conversation(mailbox: Mailbox, fake: FakeGraph) -> None:
    conversations = Conversations(mailbox)
    sizes = {s.conversation_id: s.messages for s in await mailbox.conversation_sizes(["c-rel", "c-lunch"])}
    assert sizes == {"c-rel": 3, "c-lunch": 1}  # junk m4 left out, as in get_conversation
    with_junk = await mailbox.conversation_sizes(["c-rel"], include_deleted_items=True)
    assert with_junk[0].messages == 4 and not with_junk[0].at_least
    del fake.messages["m2"]
    assert (await mailbox.conversation_sizes(["c-rel"]))[0].messages == 2
    assert len((await conversations.get_conversation("c-rel")).messages) == 2


async def test_single_oversized_message_is_truncated_not_skipped(mailbox: Mailbox, fake: FakeGraph) -> None:
    fake.add(
        FakeMessage("big", "Huge", "f-inbox", "2026-09-01T00:00:00Z", conversation="c-big", text="x" * 500)
    )
    conversation = await Conversations(mailbox).get_conversation("c-big", max_chars=100)
    assert conversation.messages[0].truncated and len(conversation.messages[0].text or "") == 100


async def test_unknown_conversation(mailbox: Mailbox) -> None:
    with pytest.raises(NotFound):
        await Conversations(mailbox).get_conversation("nope")


def test_base_subject_strips_reply_and_forward_prefixes() -> None:
    assert base_subject("RE: FW: Enc: Relatório") == "Relatório"


def _new_process_after_days(fake: FakeGraph, tmp_path: Path) -> Mailbox:
    """A new connector process on the same local store, whose folder cache is three days old."""
    with sqlite3.connect(tmp_path / "m.sqlite3") as db:
        db.execute("UPDATE meta SET value = ? WHERE key = 'folders_at'", (str(time.time() - 3 * 86400),))
    transport = Transport(StaticTokens(), client=httpx.AsyncClient(transport=fake.transport()))
    return Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "m.sqlite3", "fp"))


async def test_a_fresh_folder_cache_is_used_without_asking_the_server(
    mailbox: Mailbox, fake: FakeGraph, tmp_path: Path
) -> None:
    await mailbox.folders()  # fills the cache
    transport = Transport(StaticTokens(), client=httpx.AsyncClient(transport=fake.transport()))
    other = Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "m.sqlite3", "fp"))  # a new process
    fake.calls.clear()
    assert {f.id for f in await other.folders()} >= {"f-inbox", "f-rie"} and fake.calls == []


async def test_a_stale_folder_cache_is_never_used(mailbox: Mailbox, fake: FakeGraph, tmp_path: Path) -> None:
    await mailbox.folders()  # fills the cache
    fake.add_folder("f-new", "New folder")
    later = _new_process_after_days(fake, tmp_path)
    assert "f-new" in {f.id for f in await later.folders()}  # waited for the server


async def test_mail_in_a_folder_created_while_the_cache_was_stale_is_listed_folder_by_folder(
    mailbox: Mailbox, fake: FakeGraph, tmp_path: Path
) -> None:
    _junk_heavy(fake)
    await mailbox.folders()  # cached days ago, before the folder below existed
    fake.add_folder("f-invoices", "Invoices", parent="f-inbox")  # e.g. a new rule files mail there
    fake.add(FakeMessage("inv1", "Invoice 42", "f-invoices", "2026-10-09T09:00:00Z", conversation="c-inv"))
    page = await _new_process_after_days(fake, tmp_path).list_messages(limit=5)
    assert any(n.startswith("Listed folder by folder") for n in page.coverage.notes)
    assert page.items[0].id == "inv1" and page.items[0].folder == "Inbox/Invoices"


async def test_a_folder_deleted_during_a_per_folder_listing_does_not_fail_it(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    _junk_heavy(fake)
    first = await mailbox.list_messages(limit=3)
    assert [m.id for m in first.items] == ["n8", "n7", "n6"]
    fake.folders = [f for f in fake.folders if f["id"] != "f-archive"]  # deleted and emptied in Outlook
    for mid in [m for m, msg in fake.messages.items() if msg.folder == "f-archive"]:
        del fake.messages[mid]
    rest, cursor = [], first.cursor
    while cursor:
        page = await mailbox.list_messages(cursor=cursor)
        rest += [m.id for m in page.items]
        cursor = page.cursor
    assert rest == ["n4", "n3", "n2", "m5", "m3", "m2", "m1"]  # n5 and n1 were in Archive


async def test_a_folder_that_answers_not_found_mid_listing_is_dropped(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    _junk_heavy(fake)
    await mailbox.folders()  # the folder list still has Archive below
    fake.folders = [f for f in fake.folders if f["id"] != "f-archive"]
    fake.calls.clear()
    page = await mailbox.list_messages(limit=20)
    assert "n5" not in {m.id for m in page.items} and page.coverage.complete
    assert fake.calls.count("GET /v1.0/me/mailFolders") == 1  # the folder list was refreshed once


async def test_conversation_marks_missing_bodies_with_the_export_error_block(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    await mailbox.folders()
    fake.throttle_items = 10_000
    conversation = await Conversations(mailbox).get_conversation("c-rel")
    text = conversation.messages[0].text or ""
    assert text.startswith("[EXPORT ERROR] The body of this message could not be fetched.\n")
    assert "  Error:  HTTP 429 ApplicationThrottled" in text and "  Fix:    export it again" in text
    assert not conversation.coverage.complete  # throttling is retryable
    assert conversation.body_errors == 3 and all(
        t.export_error and t.export_error.retry for t in conversation.messages
    )


async def test_conversation_body_denied_does_not_make_coverage_incomplete(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    fake.fail[r"/me/messages/m2"] = 403
    conversation = await Conversations(mailbox).get_conversation("c-rel")
    texts = {t.message.id: t.text or "" for t in conversation.messages}
    assert "  Likely: access denied for this item" in texts["m2"] and texts["m1"] == "First report"
    assert conversation.coverage.complete  # retrying will not help
    # ...but the caller still learns about it without reading the text
    assert conversation.body_errors == 1 and any(
        "1 message body(ies) could not be fetched" in n for n in conversation.coverage.notes
    )
    errors = {t.message.id: t.export_error for t in conversation.messages}
    assert errors["m1"] is None and errors["m2"] is not None and errors["m2"].status == 403


async def test_conversation_coverage_ignores_a_body_cut_to_fit(mailbox: Mailbox, fake: FakeGraph) -> None:
    fake.add(
        FakeMessage("big", "Huge", "f-inbox", "2026-09-01T00:00:00Z", conversation="c-big", text="x" * 5000)
    )
    conversation = await Conversations(mailbox).get_conversation("c-big", max_chars=1000)
    assert (
        conversation.messages[0].truncated and conversation.cursor is None and conversation.coverage.complete
    )


async def test_truncated_listing_stays_incomplete_when_bodies_fit(
    mailbox: Mailbox, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("outlook_connector.remote.graph_mail.MAX_CONVERSATION", 2)
    conversation = await Conversations(mailbox).get_conversation("c-rel", max_chars=100_000)
    assert not conversation.coverage.complete


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


# ---------------------------------------------------------------- reach: hidden folders and Sync Issues


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
        assert page.coverage.excluded["hidden"] == 3  # h1, o1 and c1 (Sync Issues)
        found = await mailbox.search("budget", include_deleted_items=include)
        assert not {"h1", "o1"} & {m.id for hit in found.conversations for m in hit.matching_messages}
    assert "f-hidden" not in {f.id for f in await mailbox.folders()}
    with pytest.raises(InvalidRequest, match="hidden folder"):
        await mailbox.resolve_folder("TeamsMeetings")
    assert (await mailbox.conversation_sizes(["c-h1", "c-o1"]))[0].messages == 0


async def test_sync_issues_are_out_of_reach(mailbox: Mailbox, fake: FakeGraph) -> None:
    add_out_of_reach_mail(fake)
    fake.add_folder("f-sync2", "Sync Issues 2", alias="syncissues")  # Graph does not mark it hidden
    fake.add_folder("f-local", "Local Failures", parent="f-sync2", alias="localfailures")
    fake.add(FakeMessage("l1", "budget l1", "f-local", "2026-10-01T09:00:00Z", conversation="c-l1"))
    for include in (False, True):  # include_deleted_items never brings them back
        page = await mailbox.list_messages(include_deleted_items=include)
        assert not {"c1", "l1"} & {m.id for m in page.items}
        found = await mailbox.search("budget", include_deleted_items=include)
        assert not {"c1", "l1"} & {m.id for hit in found.conversations for m in hit.matching_messages}
    assert not {"f-sync", "f-conflicts", "f-sync2", "f-local"} & {f.id for f in await mailbox.folders()}
    with pytest.raises(InvalidRequest, match="hidden folder"):
        await mailbox.resolve_folder("Sync Issues 2/Local Failures")


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
    assert with_deleted.coverage.server_total == 6  # + m4 (Junk), d1 (deleted folder); never Sync Issues


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


# ---------------------------------------------------------------- per-folder listing (H7)


def _junk_heavy(fake: FakeGraph) -> None:
    """Interleaved mail in four folders, and Junk Email holding most of the mailbox."""
    for day in range(1, 9):
        folder = ("f-inbox", "f-archive", "f-sent", "f-proj")[day % 4]
        fake.add(
            FakeMessage(
                f"n{day}", f"Note {day}", folder, f"2026-10-0{day}T09:00:00Z", conversation=f"c-n{day}"
            )
        )
    for index in range(40):
        fake.add(
            FakeMessage(f"j{index}", "Buy now", "f-junk", f"2026-10-0{index % 9 + 1}T10:{index:02d}:00Z")
        )


async def _all_pages(mailbox: Mailbox, **options: object) -> tuple[list[str], list[str]]:
    ids: list[str] = []
    notes: list[str] = []
    cursor = None
    while True:
        page = await mailbox.list_messages(limit=3, cursor=cursor, **options)  # type: ignore[arg-type]
        ids += [m.id for m in page.items]
        notes += page.coverage.notes
        cursor = page.cursor
        if cursor is None:
            assert page.coverage.complete
            return ids, notes


async def test_a_junk_heavy_mailbox_is_listed_folder_by_folder(
    mailbox: Mailbox, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    _junk_heavy(fake)
    fake.calls.clear()
    ids, notes = await _all_pages(mailbox)
    assert not any(c == "GET /v1.0/me/messages" or "f-junk" in c for c in fake.calls)  # never read
    assert any(
        n.startswith("Listed folder by folder: folders this listing leaves out hold 77%") for n in notes
    )
    assert "Junk Email: 41" in notes[0]
    # exactly what the whole-mailbox listing returns, in the same order
    monkeypatch.setattr("outlook_connector.service.mailbox.PER_FOLDER_SHARE", 2.0)
    whole, whole_notes = await _all_pages(mailbox)
    assert ids == whole and len(ids) == 12 and ids[:3] == ["n8", "n7", "n6"]
    assert not any(n.startswith("Listed folder by folder") for n in whole_notes)


async def test_a_per_folder_cursor_keeps_each_folders_position(mailbox: Mailbox, fake: FakeGraph) -> None:
    _junk_heavy(fake)
    page = await mailbox.list_messages(limit=3)
    assert [m.id for m in page.items] == ["n8", "n7", "n6"] and page.cursor
    state = cursors.decode(page.cursor, "list_messages")
    assert state["link"] is None
    assert state["offsets"] == {"f-inbox": 1, "f-sent": 1, "f-archive": 0, "f-proj": 1, "f-rie": 0}


async def test_a_per_folder_listing_reads_only_folders_with_mail_in_the_window(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    _junk_heavy(fake)
    fake.calls.clear()
    page = await mailbox.list_messages(since=datetime(2026, 10, 7, tzinfo=UTC), limit=10)
    assert [m.id for m in page.items] == ["n8", "n7"] and page.coverage.complete and page.cursor is None
    read = {
        c.split("/")[4] for c in fake.calls if c.startswith("GET /v1.0/me/mailFolders/") and "messages" in c
    }
    assert read == {"f-inbox", "f-proj"}  # n8 in Inbox, n7 in Projects; the others have none since then


async def test_a_mostly_clean_mailbox_keeps_the_whole_mailbox_listing(
    mailbox: Mailbox, fake: FakeGraph
) -> None:
    fake.calls.clear()
    page = await mailbox.list_messages(limit=10)
    assert "GET /v1.0/me/messages" in fake.calls  # Junk holds 1 of 5 messages
    assert not any(n.startswith("Listed folder by folder") for n in page.coverage.notes)


async def test_a_named_folder_is_never_listed_folder_by_folder(mailbox: Mailbox, fake: FakeGraph) -> None:
    _junk_heavy(fake)
    fake.calls.clear()
    page = await mailbox.list_messages(folder="inbox", limit=10)
    assert [m.id for m in page.items] == ["n8", "n4", "m5", "m1"]
    assert fake.calls.count("GET /v1.0/me/mailFolders/f-inbox/messages") == 1


@pytest.mark.parametrize(
    "since,until",
    [
        (datetime(2026, 9, 28, 9, 30), datetime(2026, 9, 29, 7)),
        (datetime(2026, 9, 28, 9, 30), datetime(2026, 9, 29, 7, tzinfo=UTC)),
        (
            datetime(2026, 9, 28, 6, 30, tzinfo=timezone(timedelta(hours=-3))),
            datetime(2026, 9, 29, 9, tzinfo=timezone(timedelta(hours=2))),
        ),
    ],
)
async def test_search_normalizes_naive_and_aware_at_service_entry(
    mailbox: Mailbox, since: datetime, until: datetime
) -> None:
    result = await mailbox.search("relatório", since=since, until=until)
    assert [m.id for hit in result.conversations for m in hit.matching_messages] == ["m2"]


async def test_search_cursor_carries_normalized_utc_dates(mailbox: Mailbox) -> None:
    from outlook_connector.service.cursors import decode

    result = await mailbox.search(
        "relatório", since=datetime(2026, 9, 28), until=datetime(2026, 9, 30), limit=1
    )
    assert result.cursor
    state = decode(result.cursor, "search")
    assert state["since"] == "2026-09-28T00:00:00+00:00" and state["until"] == "2026-09-30T00:00:00+00:00"
    continuation = await mailbox.search("relatório", cursor=result.cursor, limit=1)
    assert continuation.conversations

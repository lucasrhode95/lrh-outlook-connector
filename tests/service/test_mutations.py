from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any

import httpx
import pytest

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import AccountMismatch, InvalidRequest
from outlook_connector.domain.models import MessageSummary
from outlook_connector.remote.cloud_settings import CloudSettings
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_mail import OwsMailWriter
from outlook_connector.remote.ports import MAX_CONCURRENT_REQUESTS
from outlook_connector.remote.transport import Transport
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.mutations import Mutations
from outlook_connector.service.signatures import Signatures
from outlook_connector.service.writes import Writes
from outlook_connector.store.db import Store
from tests.fakes.graph_fake import FakeGraph, FakeMessage, StaticTokens, sample_mailbox

ME = Account(tenant_id="tenant-x", object_id="user-x", username="me@example.com")


async def _no_sleep(_s: float) -> None:
    return None


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


def make(fake: FakeGraph, tmp_path: Path, account: Account = ME) -> Mutations:
    tokens = StaticTokens()
    transport = Transport(tokens, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep)
    mailbox = Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "m.sqlite3", "fp"))
    writer = OwsMailWriter(Ows(transport, tokens))
    signatures = Signatures(CloudSettings(transport, tokens), account)
    return Mutations(mailbox, writer, Writes(mailbox, writer, account, signatures).check_account)


@pytest.fixture
def mutations(fake: FakeGraph, tmp_path: Path) -> Mutations:
    return make(fake, tmp_path)


def statuses(result: Any) -> dict[str, str]:
    return {r.id: r.status for r in result.results}


async def test_read_state_per_message_with_unchanged_and_not_found(
    mutations: Mutations, fake: FakeGraph
) -> None:
    result = await mutations.set_read(["m5", "m1", "gone"], True)  # m5 unread, m1 already read
    assert statuses(result) == {"m5": "done", "m1": "unchanged", "gone": "not_found"}
    assert result.counts == {"done": 1, "unchanged": 1, "not_found": 1} and fake.messages["m5"].is_read
    (call,) = fake.ows_calls  # only m5 was sent
    assert [c["ItemId"]["Id"] for c in call[1]["ItemChanges"]] == ["m5"]


async def test_conversation_read_state_covers_its_messages_in_scope(
    mutations: Mutations, fake: FakeGraph
) -> None:
    for mid in ("m1", "m2", "m3", "m4"):
        fake.messages[mid].is_read = False
    result = await mutations.set_read([], True, conversation_ids=["c-rel"])
    assert set(statuses(result)) == {"m1", "m2", "m3"}  # m4 is in Junk: out of scope by default
    assert not fake.messages["m4"].is_read


async def test_conversation_read_state_expands_concurrently_with_transport_limit(
    mutations: Mutations, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    conversation_ids = [f"bulk-read-{index}" for index in range(9)]
    message_ids = []
    for index, conversation_id in enumerate(conversation_ids):
        message_id = f"bulk-read-message-{index}"
        message_ids.append(message_id)
        fake.add(
            FakeMessage(
                message_id,
                "Synthetic bulk read",
                "f-inbox",
                f"2026-09-{index + 1:02d}T09:00:00Z",
                conversation=conversation_id,
                is_read=False,
            )
        )
    await mutations.mailbox.folders()

    active = 0
    peak = 0
    original = mutations.mailbox.reader.conversation

    async def capture(conversation_id: str) -> tuple[list[MessageSummary], bool]:
        nonlocal active, peak
        active += 1
        peak = max(peak, active)
        try:
            await asyncio.sleep(0.001)
            return await original(conversation_id)
        finally:
            active -= 1

    monkeypatch.setattr(mutations.mailbox.reader, "conversation", capture)
    result = await mutations.set_read([], True, conversation_ids=conversation_ids)

    assert peak == MAX_CONCURRENT_REQUESTS
    assert result.counts == {"done": len(message_ids)}
    assert all(fake.messages[message_id].is_read for message_id in message_ids)


async def test_flag(mutations: Mutations, fake: FakeGraph) -> None:
    assert statuses(await mutations.set_flag(["m1"], True)) == {"m1": "done"} and fake.messages["m1"].flagged
    assert statuses(await mutations.set_flag(["m1"], True)) == {"m1": "unchanged"}


async def test_move_to_a_folder_by_path_or_alias(mutations: Mutations, fake: FakeGraph) -> None:
    result = await mutations.move(["m1", "m5"], "Inbox/Projects/Project Alpha")
    assert statuses(result) == {"m1": "done", "m5": "done"} and fake.messages["m1"].folder == "f-project"
    assert statuses(await mutations.move(["m1"], "f-project")) == {"m1": "unchanged"}
    await mutations.move(["m1"], "archive")
    assert fake.messages["m1"].folder == "f-archive"


async def test_move_refuses_deleted_items_and_hidden_folders(mutations: Mutations, fake: FakeGraph) -> None:
    with pytest.raises(InvalidRequest, match="delete_messages"):
        await mutations.move(["m1"], "deleteditems")
    fake.add_folder("f-hidden", "Secret", hidden=True)
    with pytest.raises(InvalidRequest):
        await mutations.move(["m1"], "Secret")
    assert not fake.ows_calls


async def test_items_out_of_reach_are_not_touched(mutations: Mutations, fake: FakeGraph) -> None:
    fake.add_folder("f-hidden", "Secret", hidden=True)
    fake.add(FakeMessage("h1", "hidden", "f-hidden", "2026-09-01T00:00:00Z", is_read=False))
    assert statuses(await mutations.set_read(["h1"], True)) == {"h1": "failed"}
    assert not fake.ows_calls


async def test_delete_moves_to_deleted_items_and_never_further(mutations: Mutations, fake: FakeGraph) -> None:
    fake.add_folder("f-old", "Old project", parent="f-deleted")
    fake.add(FakeMessage("d2", "old", "f-old", "2026-09-01T00:00:00Z"))
    result = await mutations.delete(["m5", "d2"])
    assert statuses(result) == {"m5": "done", "d2": "unchanged"}
    assert fake.messages["m5"].folder == "f-deleted" and fake.messages["d2"].folder == "f-old"
    assert statuses(await mutations.delete(["m5"])) == {"m5": "unchanged"}
    assert len(fake.ows_calls) == 1


async def test_unclear_answer_is_checked_by_reading_back(mutations: Mutations, fake: FakeGraph) -> None:
    fake.ows_next = ["done-no-answer"]
    assert statuses(await mutations.set_flag(["m1"], True)) == {"m1": "done"}
    fake.ows_next = ["no-answer"]
    result = await mutations.set_flag(["m5"], True)
    assert statuses(result) == {"m5": "unknown"} and "not retried" in (result.results[0].detail or "")
    assert len(fake.ows_calls) == 2


async def test_failed_items_are_reported_not_hidden(mutations: Mutations, fake: FakeGraph) -> None:
    fake.ows_next = [{"ResponseClass": "Error", "ResponseCode": "ErrorAccessDenied"}]
    result = await mutations.move(["m1"], "archive")
    assert result.results[0].status == "failed" and result.results[0].detail == "ErrorAccessDenied"


async def test_limits_and_account(fake: FakeGraph, tmp_path: Path, mutations: Mutations) -> None:
    with pytest.raises(InvalidRequest, match="At most 100"):
        await mutations.set_read([f"x{i}" for i in range(101)], True)
    with pytest.raises(InvalidRequest, match="at least one"):
        await mutations.set_read([], True)
    other = make(fake, tmp_path / "o", Account(tenant_id="tenant-x", object_id="other", username=None))
    with pytest.raises(AccountMismatch):
        await other.set_read(["m5"], True)
    assert not fake.ows_calls


async def test_many_messages_go_in_chunks(mutations: Mutations, fake: FakeGraph) -> None:
    for i in range(45):
        fake.add(
            FakeMessage(f"u{i}", "x", "f-inbox", "2026-09-01T00:00:00Z", conversation=f"cu{i}", is_read=False)
        )
    result = await mutations.set_read([f"u{i}" for i in range(45)], True)
    assert result.counts == {"done": 45}
    assert [len(body["ItemChanges"]) for _, body in fake.ows_calls] == [20, 20, 5]


async def test_new_deleted_items_subfolder_is_left_alone(mutations: Mutations, fake: FakeGraph) -> None:
    await mutations.mailbox.folders()  # cached before the subfolder exists
    fake.add_folder("f-new", "Fresh", parent="f-deleted")
    fake.add(FakeMessage("d3", "old", "f-new", "2026-09-01T00:00:00Z"))
    assert statuses(await mutations.delete(["d3"])) == {"d3": "unchanged"}
    assert not fake.ows_calls


async def test_mismatched_item_results_are_read_back(mutations: Mutations, fake: FakeGraph) -> None:
    fake.ows_next = [{"ResponseClass": "Success", "ResponseCode": "NoError"}]  # one result for two
    result = await mutations.set_flag(["m1", "m5"], True)
    assert statuses(result) == {"m1": "unknown", "m5": "unknown"} and len(fake.ows_calls) == 1


@pytest.mark.parametrize("read", [True, False])
async def test_long_conversations_use_normal_chunks_and_compact_counts(
    mutations: Mutations, fake: FakeGraph, read: bool
) -> None:
    for i in range(125):
        fake.add(
            FakeMessage(
                f"long-{i}",
                "x",
                "f-inbox",
                "2026-09-01T00:00:00Z",
                conversation="long",
                is_read=not read if i < 120 else read,
            )
        )
    result = await mutations.set_read([], read, conversation_ids=["long"])
    assert result.counts == {"done": 120, "unchanged": 5}
    assert result.results == [] and any("summarized" in n for n in result.notes)
    assert len(fake.ows_calls) == 6 and all(len(b["ItemChanges"]) == 20 for _, b in fake.ows_calls)
    assert all(fake.messages[f"long-{i}"].is_read is read for i in range(125))


async def test_expansion_truncation_is_reported_at_existing_limit(
    mutations: Mutations, fake: FakeGraph
) -> None:
    for i in range(1001):
        fake.add(
            FakeMessage(f"cap-{i}", "x", "f-inbox", "2026-09-01T00:00:00Z", conversation="cap", is_read=False)
        )
    result = await mutations.set_read([], True, conversation_ids=["cap"])
    assert result.counts == {"done": 1000} and not result.results
    assert any("truncated" in n and "1,000" in n for n in result.notes)
    assert (
        len(fake.ows_calls) == 50
        and sum(m.is_read for mid, m in fake.messages.items() if mid.startswith("cap-")) == 1000
    )


async def test_explicit_limit_checked_before_expansion(mutations: Mutations, fake: FakeGraph) -> None:
    with pytest.raises(InvalidRequest, match="At most 100"):
        await mutations.set_read([f"x{i}" for i in range(101)], True, conversation_ids=["long"])
    assert not fake.calls and not fake.ows_calls


async def test_large_expansion_retains_explicit_and_failure_results(
    mutations: Mutations, fake: FakeGraph
) -> None:
    for i in range(120):
        fake.add(
            FakeMessage(
                f"compact-{i}", "x", "f-inbox", "2026-09-01T00:00:00Z", conversation="compact", is_read=False
            )
        )
    original = fake.ows_UpdateItem

    def one_failure(body: dict[str, Any]) -> list[dict[str, Any]]:
        result = original(body)
        if body["ItemChanges"][0]["ItemId"]["Id"] == "compact/0":
            result[1] = {"ResponseClass": "Error", "ResponseCode": "ErrorAccessDenied"}
        return result

    fake.ows_UpdateItem = one_failure  # type: ignore[method-assign]
    result = await mutations.set_read(["compact-0", "gone"], True, conversation_ids=["compact", "compact"])
    assert result.counts == {"done": 119, "failed": 1, "not_found": 1}
    assert statuses(result) == {"compact-0": "done", "gone": "not_found", "compact-1": "failed"}
    assert len(fake.ows_calls) == 6


@pytest.mark.parametrize("action", ["read", "flag", "move", "delete"])
@pytest.mark.parametrize("keep_going", [True, False])
async def test_failed_middle_chunk_preserves_partial_results(
    mutations: Mutations, fake: FakeGraph, action: str, keep_going: bool
) -> None:
    ids = [f"partial-{i}" for i in range(45)]
    for mid in ids:
        fake.add(FakeMessage(mid, "synthetic", "f-inbox", "2026-09-01T00:00:00Z", is_read=False))
    fake.ows_next = [None, 429, None]
    if action == "read":
        result = await mutations.set_read(ids, True, continue_on_error=keep_going)
    elif action == "flag":
        result = await mutations.set_flag(ids, True, continue_on_error=keep_going)
    elif action == "move":
        result = await mutations.move(ids, "archive", continue_on_error=keep_going)
    else:
        result = await mutations.delete(ids, continue_on_error=keep_going)
    assert [r.id for r in result.results] == ids
    assert all(r.status == "done" for r in result.results[:20])
    assert all(r.status == "failed" and r.detail != "not sent" for r in result.results[20:40])
    assert result.counts == ({"done": 25, "failed": 20} if keep_going else {"done": 20, "failed": 25})
    assert len(fake.ows_calls) == (3 if keep_going else 2)
    if not keep_going:
        assert all(r.status == "failed" and r.detail == "not sent" for r in result.results[40:])


async def test_default_continues_after_failed_chunk(mutations: Mutations, fake: FakeGraph) -> None:
    ids = [f"default-{i}" for i in range(21)]
    for mid in ids:
        fake.add(FakeMessage(mid, "x", "f-inbox", "2026-09-01T00:00:00Z"))
    fake.ows_next = [429, None]
    result = await mutations.set_flag(ids, True)
    assert result.counts == {"failed": 20, "done": 1} and len(fake.ows_calls) == 2


async def test_failed_readback_is_unknown_and_stops_only_when_requested(
    mutations: Mutations, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    from outlook_connector.domain.errors import Throttled

    ids = [f"unknown-{i}" for i in range(41)]
    for mid in ids:
        fake.add(FakeMessage(mid, "x", "f-inbox", "2026-09-01T00:00:00Z"))
    original = mutations.mailbox.reader.get_summaries
    calls = 0

    async def unavailable_after_initial(*args: Any, **kwargs: Any) -> Any:
        nonlocal calls
        calls += 1
        if calls > 1:
            raise Throttled("Graph unavailable")
        return await original(*args, **kwargs)

    monkeypatch.setattr(mutations.mailbox.reader, "get_summaries", unavailable_after_initial)
    fake.ows_next = [None, "no-answer"]
    result = await mutations.set_flag(ids, True, continue_on_error=False)
    assert result.counts == {"done": 20, "unknown": 20, "failed": 1}
    assert result.results[-1].detail == "not sent" and len(fake.ows_calls) == 2


async def test_per_item_error_stops_subsequent_chunks(mutations: Mutations, fake: FakeGraph) -> None:
    ids = [f"item-{i}" for i in range(21)]
    for mid in ids:
        fake.add(FakeMessage(mid, "x", "f-inbox", "2026-09-01T00:00:00Z"))
    original = fake.ows_UpdateItem

    def fail_one(body: dict[str, Any]) -> list[dict[str, Any]]:
        results = original(body)
        results[-1] = {"ResponseClass": "Error", "ResponseCode": "ErrorAccessDenied"}
        return results

    fake.ows_UpdateItem = fail_one  # type: ignore[method-assign]
    result = await mutations.set_flag(ids, True, continue_on_error=False)
    assert result.counts == {"done": 19, "failed": 2} and result.results[-1].detail == "not sent"

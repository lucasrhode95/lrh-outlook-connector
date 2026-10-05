from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx
import pytest

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import AccountMismatch, InvalidRequest
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_mail import OwsMailWriter
from outlook_connector.remote.transport import Transport
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.mutations import Mutations
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
    return Mutations(mailbox, writer, Writes(mailbox, writer, account).check_account)


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


async def test_flag(mutations: Mutations, fake: FakeGraph) -> None:
    assert statuses(await mutations.set_flag(["m1"], True)) == {"m1": "done"} and fake.messages["m1"].flagged
    assert statuses(await mutations.set_flag(["m1"], True)) == {"m1": "unchanged"}


async def test_move_to_a_folder_by_path_or_alias(mutations: Mutations, fake: FakeGraph) -> None:
    result = await mutations.move(["m1", "m5"], "Inbox/Projects/RIE")
    assert statuses(result) == {"m1": "done", "m5": "done"} and fake.messages["m1"].folder == "f-rie"
    assert statuses(await mutations.move(["m1"], "f-rie")) == {"m1": "unchanged"}
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

from __future__ import annotations

import httpx
import pytest

from outlook_connector.domain.errors import NotFound, Throttled, Upstream, WriteOutcomeUnknown
from outlook_connector.domain.models import EmailProposal
from outlook_connector.remote import ids
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_mail import OwsMailWriter
from outlook_connector.remote.transport import Transport
from tests.fakes.graph_fake import FakeGraph, StaticTokens, sample_mailbox


async def _no_sleep(_s: float) -> None:
    return None


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


def writer_for(fake: FakeGraph, tokens: StaticTokens | None = None) -> OwsMailWriter:
    tokens = tokens or StaticTokens()
    transport = Transport(tokens, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep)
    return OwsMailWriter(Ows(transport, tokens))


def proposal(**changes: object) -> EmailProposal:
    fields: dict[str, object] = {
        "sender": "me@example.com", "to": ["bob@example.com"], "subject": "Hi", "body": "Hello Bob",
        "confirmation": "SEND-TEST",
    }  # fmt: skip
    return EmailProposal.model_validate(fields | changes)


def test_ids_swap_the_base64_alphabet() -> None:
    assert ids.to_ows("AAMk-a_b") == "AAMk/a+b" and ids.to_graph("AAMk/a+b") == "AAMk-a_b"


async def test_draft_is_saved_in_drafts_and_its_id_mapped_back(fake: FakeGraph) -> None:
    draft_id = await writer_for(fake).create_draft(proposal(cc=["c@example.com"]))
    assert draft_id and "-" in draft_id and "_" in draft_id  # Graph alphabet
    saved = fake.messages[draft_id]
    assert saved.folder == "f-drafts" and saved.is_draft and saved.to == ("bob@example.com",)
    assert saved.cc == ("c@example.com",) and saved.text == "Hello Bob"
    action, body = fake.ows_calls[0]
    assert action == "CreateItem" and body["MessageDisposition"] == "SaveOnly"
    assert body["ComposeOperation"] == "newMail" and body["Items"][0]["Body"]["BodyType"] == "Text"


async def test_send_uses_send_and_save_copy(fake: FakeGraph) -> None:
    await writer_for(fake).send(proposal())
    assert fake.ows_calls[0][1]["MessageDisposition"] == "SendAndSaveCopy"
    sent = [m for m in fake.messages.values() if m.folder == "f-sent" and m.subject == "Hi"]
    assert len(sent) == 1 and not sent[0].is_draft


async def test_small_requests_travel_in_the_header_large_ones_in_the_body(fake: FakeGraph) -> None:
    seen: list[bool] = []
    original = fake.handle_ows

    def spy(request: httpx.Request) -> httpx.Response:
        seen.append("x-owa-urlpostdata" in request.headers and not request.content)
        return original(request)

    fake.handle_ows = spy  # type: ignore[method-assign]
    writer = writer_for(fake)
    await writer.create_draft(proposal(body="short"))
    await writer.create_draft(proposal(body="long " * 1000))
    assert seen == [True, False]
    assert fake.messages[max(fake.messages)].text.startswith("long")


async def test_replies_reference_the_original_with_ows_ids(fake: FakeGraph) -> None:
    fake.messages["m-1_a"] = fake.messages.pop("m1")
    fake.messages["m-1_a"].id = "m-1_a"
    await writer_for(fake).create_draft(
        proposal(reply_to_message_id="m-1_a", reply_all=True, subject="RE: Relatório BE semanal")
    )
    item = fake.ows_calls[0][1]["Items"][0]
    assert item["__type"] == "ReplyAllToItem:#Exchange" and item["ReferenceItemId"]["Id"] == "m/1+a"
    assert fake.ows_calls[0][1]["ComposeOperation"] == "replyAll"


async def test_item_errors_name_the_code(fake: FakeGraph) -> None:
    fake.ows_next = [{"ResponseClass": "Error", "ResponseCode": "ErrorInvalidRecipients", "MessageText": "x"}]
    with pytest.raises(Upstream, match="ErrorInvalidRecipients"):
        await writer_for(fake).send(proposal())
    fake.ows_next = [{"ResponseClass": "Error", "ResponseCode": "ErrorItemNotFound"}]
    with pytest.raises(NotFound):
        await writer_for(fake).create_draft(proposal(reply_to_message_id="gone"))


async def test_no_answer_or_server_error_is_an_unknown_outcome_never_retried(fake: FakeGraph) -> None:
    for script in ("no-answer", 500, 502):
        fake.ows_calls.clear()
        fake.ows_next = [script]
        with pytest.raises(WriteOutcomeUnknown, match="not retried"):
            await writer_for(fake).send(proposal())
        assert len(fake.ows_calls) == 1


async def test_a_success_without_readable_results_is_an_unknown_outcome(fake: FakeGraph) -> None:
    for script in ("no-items", "not-json"):
        fake.ows_next = [script]
        with pytest.raises(WriteOutcomeUnknown):
            await writer_for(fake).send(proposal())


async def test_throttled_write_is_not_retried(fake: FakeGraph) -> None:
    fake.ows_next = [429]
    with pytest.raises(Throttled):
        await writer_for(fake).send(proposal())
    assert len(fake.ows_calls) == 1


async def test_rejected_token_is_renewed_once_before_the_write_is_sent(fake: FakeGraph) -> None:
    tokens = StaticTokens()
    fake.reject_tokens = 1  # refused before OWS processed anything: safe to send again
    await writer_for(fake, tokens).send(proposal())
    assert tokens.renewals == [{"force_refresh": True}] and len(fake.ows_calls) == 1


async def test_anchor_mailbox_is_the_write_account(fake: FakeGraph) -> None:
    headers: list[str] = []
    original = fake.handle_ows

    def spy(request: httpx.Request) -> httpx.Response:
        headers.append(request.headers["x-anchormailbox"])
        return original(request)

    fake.handle_ows = spy  # type: ignore[method-assign]
    await writer_for(fake).send(proposal())
    assert headers == ["AAD-SMTP:me@example.com"]


# ---------------------------------------------------------------- mutations


async def test_update_item_reports_each_message(fake: FakeGraph) -> None:
    writer = writer_for(fake)
    assert await writer.set_read(["m5", "gone"], True) == {"m5": None, "gone": "ErrorItemNotFound"}
    assert fake.messages["m5"].is_read
    assert await writer.set_flag(["m1"], True) == {"m1": None} and fake.messages["m1"].flagged
    (change,) = fake.ows_calls[0][1]["ItemChanges"][:1]
    assert change["ItemId"]["Id"] == "m5" and fake.ows_calls[0][1]["ConflictResolution"] == "AlwaysOverwrite"


async def test_move_to_a_well_known_or_any_folder(fake: FakeGraph) -> None:
    from outlook_connector.remote.ports import FolderTarget

    writer = writer_for(fake)
    assert await writer.move(["m1"], FolderTarget("f-archive", "archive")) == {"m1": None}
    assert fake.messages["m1"].folder == "f-archive"
    assert fake.ows_calls[-1][1]["ToFolderId"]["BaseFolderId"]["__type"] == "DistinguishedFolderId:#Exchange"
    assert await writer.move(["m1"], FolderTarget("f-rie")) == {"m1": None}
    assert fake.messages["m1"].folder == "f-rie"
    assert fake.ows_calls[-1][1]["ToFolderId"]["BaseFolderId"] == {
        "__type": "FolderId:#Exchange",
        "Id": "f/rie",
    }


async def test_delete_only_moves_to_deleted_items(fake: FakeGraph) -> None:
    assert await writer_for(fake).delete(["m5", "gone"]) == {"m5": None, "gone": "ErrorItemNotFound"}
    assert fake.messages["m5"].folder == "f-deleted"
    assert fake.ows_calls[0][1]["DeleteType"] == "MoveToDeletedItems"


async def test_mismatched_item_results_are_an_error(fake: FakeGraph) -> None:
    fake.ows_next = [{"ResponseClass": "Success", "ResponseCode": "NoError"}]  # one result for two
    with pytest.raises(WriteOutcomeUnknown, match="1 item results for 2 messages"):
        await writer_for(fake).set_read(["m1", "m5"], True)


def ows_for(fake: FakeGraph) -> Ows:
    tokens = StaticTokens()
    return Ows(
        Transport(tokens, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep), tokens
    )


async def test_bare_request_posts_the_request_object_itself(fake: FakeGraph) -> None:
    fake.ows_next = [{"WasSuccessful": True, "ErrorCode": 0, "InboxRuleCollection": {"InboxRules": []}}]
    answer = await ows_for(fake).call_request("GetInboxRule", {"UseServerRulesLoader": True}, time_zone="UTC")
    assert answer["InboxRuleCollection"] == {"InboxRules": []}
    action, sent = fake.ows_calls[0]
    assert action == "GetInboxRule" and sent["__type"] == "GetInboxRuleRequest:#Exchange"
    assert "Body" not in sent and sent["UseServerRulesLoader"] is True
    assert sent["Header"]["TimeZoneContext"]["TimeZoneDefinition"]["Id"] == "UTC"
    await ows_for(fake).call_request("EnableInboxRule", {"Identity": {"RawIdentity": "r"}})
    assert "TimeZoneContext" not in fake.ows_calls[1][1]["Header"]


async def test_bare_request_failures(fake: FakeGraph) -> None:
    ows = ows_for(fake)
    fake.ows_next = [{"WasSuccessful": False, "ErrorCode": 5, "ErrorMessage": "no such rule"}]
    with pytest.raises(Upstream, match="no such rule"):
        await ows.call_request("RemoveInboxRule", {"Identity": {"RawIdentity": "r"}})
    fake.ows_next = [{"Body": {}}]
    with pytest.raises(WriteOutcomeUnknown):
        await ows.call_request("RemoveInboxRule", {"Identity": {"RawIdentity": "r"}})
    fake.ows_next = ["no-answer"]
    with pytest.raises(WriteOutcomeUnknown):
        await ows.call_request("RemoveInboxRule", {"Identity": {"RawIdentity": "r"}})
    assert len(fake.ows_calls) == 3  # each sent once, never retried

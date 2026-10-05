from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import httpx
import pytest

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import (
    AccountMismatch,
    ConnectorError,
    InvalidRequest,
    NotFound,
    Throttled,
)
from outlook_connector.domain.models import OutgoingMessage
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_mail import OwsMailWriter
from outlook_connector.remote.transport import Transport
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.writes import Writes
from outlook_connector.store.db import Store
from tests.fakes.graph_fake import FakeAttachment, FakeGraph, FakeMessage, StaticTokens, sample_mailbox

ME = Account(tenant_id="tenant-x", object_id="user-x", username="me@example.com")


async def _no_sleep(_s: float) -> None:
    return None


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


def make_writes(fake: FakeGraph, tmp_path: Path, tokens: Any = None, account: Account = ME) -> Writes:
    tokens = tokens or StaticTokens()
    transport = Transport(tokens, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep)
    mailbox = Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "m.sqlite3", "fp"))
    return Writes(mailbox, OwsMailWriter(Ows(transport, tokens)), account)


@pytest.fixture
def writes(fake: FakeGraph, tmp_path: Path) -> Writes:
    return make_writes(fake, tmp_path)


def message(**changes: Any) -> OutgoingMessage:
    return OutgoingMessage.model_validate(
        {"to": ["bob@example.com"], "subject": "Hi", "body": "Hello"} | changes
    )


# ---------------------------------------------------------------- proposal


async def test_proposal_is_resolved_and_changes_nothing(writes: Writes, fake: FakeGraph) -> None:
    proposal = await writes.propose(message(cc=[" Carol@Example.com ", "carol@example.com"]))
    assert proposal.sender == "me@example.com" and proposal.cc == ["Carol@Example.com"]
    assert proposal.confirmation.startswith("SEND-") and not fake.ows_calls


async def test_any_change_changes_the_confirmation(writes: Writes) -> None:
    base = (await writes.propose(message())).confirmation
    for change in ({"to": ["eve@example.com"]}, {"subject": "Hi!"}, {"body": "Hello."}, {"bcc": ["x@y.com"]}):
        assert (await writes.propose(message(**change))).confirmation != base
    assert (await writes.propose(message())).confirmation == base  # deterministic


async def test_confirmation_is_bound_to_the_account(fake: FakeGraph, tmp_path: Path) -> None:
    other = Account(tenant_id="tenant-x", object_id="someone-else", username="me@example.com")
    mine = await make_writes(fake, tmp_path).propose(message())
    theirs = await make_writes(fake, tmp_path / "o", account=other).propose(message())
    assert mine.confirmation != theirs.confirmation


@pytest.mark.parametrize(
    ("changes", "error"),
    [
        ({"to": []}, "at least one recipient"),
        ({"to": ["not-an-address"]}, "Not an email address"),
        ({"to": ["Bob <bob@example.com>"]}, "Not an email address"),
        ({"cc": ["bob@example.com"]}, "listed twice"),
        ({"subject": "  "}, "needs a subject"),
        ({"body": " "}, "body is empty"),
        ({"reply_all": True}, "reply_all needs reply_to_message_id"),
        ({"to": [f"u{i}@example.com" for i in range(101)]}, "At most 100"),
    ],
)
async def test_invalid_messages_are_refused(writes: Writes, changes: dict[str, Any], error: str) -> None:
    with pytest.raises(InvalidRequest, match=error):
        await writes.propose(message(**changes))


async def test_reply_defaults_follow_outlook(writes: Writes, fake: FakeGraph) -> None:
    fake.messages["m1"].to = ("me@example.com", "dan@example.com")
    fake.messages["m1"].cc = ("erin@example.com", "me@example.com")
    reply = await writes.propose(OutgoingMessage(reply_to_message_id="m1", body="Thanks"))
    assert reply.to == ["alice@example.com"] and reply.subject == "RE: Relatório BE semanal"
    assert reply.quotes_original and not reply.cc
    everyone = await writes.propose(OutgoingMessage(reply_to_message_id="m1", reply_all=True, body="Thanks"))
    assert everyone.to == ["alice@example.com", "dan@example.com"] and everyone.cc == ["erin@example.com"]


async def test_reply_to_your_own_message_goes_to_its_recipients(writes: Writes) -> None:
    reply = await writes.propose(OutgoingMessage(reply_to_message_id="m2", body="Ping"))  # m2: sent by me
    assert reply.to == ["alice@example.com"] and reply.subject == "RE: Relatório BE semanal"


async def test_reply_to_mail_you_sent_yourself_goes_back_to_you(writes: Writes, fake: FakeGraph) -> None:
    fake.messages["m5"].sender = "Me@Example.com"  # live: from you, to you (V2 step 2)
    fake.messages["m5"].to = ("me@example.com",)
    for reply_all in (False, True):
        reply = await writes.propose(OutgoingMessage(reply_to_message_id="m5", reply_all=reply_all, body="x"))
        assert reply.to == ["Me@Example.com"] and not reply.cc and reply.subject == "RE: Lunch?"


async def test_reply_to_a_missing_message_is_not_found(writes: Writes) -> None:
    with pytest.raises(NotFound):
        await writes.propose(OutgoingMessage(reply_to_message_id="gone", body="x"))


# ---------------------------------------------------------------- drafts


async def test_draft_is_saved_read_back_and_never_sent(writes: Writes, fake: FakeGraph) -> None:
    draft = await writes.create_draft(message(bcc=["boss@example.com"]))
    assert draft.verified and draft.folder == "Drafts" and draft.proposal.bcc == ["boss@example.com"]
    saved = fake.messages[draft.id]
    assert saved.folder == "f-drafts" and saved.is_draft
    assert [body["MessageDisposition"] for _, body in fake.ows_calls] == ["SaveOnly"]
    content = await writes.mailbox.get_message(draft.id)
    assert content.text == "Hello"


async def test_reply_draft_joins_the_conversation(writes: Writes, fake: FakeGraph) -> None:
    draft = await writes.create_draft(OutgoingMessage(reply_to_message_id="m1", body="On it"))
    assert fake.messages[draft.id].conversation == "c-rel"
    assert fake.ows_calls[0][1]["Items"][0]["__type"] == "ReplyToItem:#Exchange"


async def test_reply_body_is_html_so_the_history_keeps_its_formatting(
    writes: Writes, fake: FakeGraph
) -> None:
    await writes.create_draft(OutgoingMessage(reply_to_message_id="m1", body="a < b & c\nsecond line"))
    content = fake.ows_calls[0][1]["Items"][0]["NewBodyContent"]
    assert content == {
        "__type": "BodyContentType:#Exchange",
        "BodyType": "HTML",
        "Value": "<div>a &lt; b &amp; c<br>second line</div>",
    }


# ---------------------------------------------------------------- send


async def test_send_needs_the_matching_confirmation(writes: Writes, fake: FakeGraph) -> None:
    proposal = await writes.propose(message())
    with pytest.raises(InvalidRequest, match="does not match"):
        await writes.send(message(), "SEND-00000000")
    with pytest.raises(InvalidRequest, match="does not match"):  # the message changed after confirming
        await writes.send(message(to=["eve@example.com"]), proposal.confirmation)
    assert not fake.ows_calls
    result = await writes.send(message(), proposal.confirmation.lower())
    assert result.status == "sent"
    assert [body["MessageDisposition"] for _, body in fake.ows_calls] == ["SendAndSaveCopy"]


async def test_send_refuses_a_write_sign_in_of_another_account(fake: FakeGraph, tmp_path: Path) -> None:
    class OtherWriteAccount(StaticTokens):
        class _Other(StaticTokens._T):
            def claims(self) -> dict[str, Any]:
                return {"tid": "tenant-x", "oid": "intruder", "upn": "x@example.com"}

        def get_token(self, profile: str, **renewal: Any) -> Any:
            return self._Other() if profile == "write" else super().get_token(profile, **renewal)

    writes = make_writes(fake, tmp_path, tokens=OtherWriteAccount())
    proposal = await writes.propose(message())
    with pytest.raises(AccountMismatch):
        await writes.send(message(), proposal.confirmation)
    assert not fake.ows_calls


async def test_unclear_send_that_went_through_is_found_in_sent_items(writes: Writes, fake: FakeGraph) -> None:
    proposal = await writes.propose(message())
    fake.ows_next = ["done-no-answer"]
    result = await writes.send(message(), proposal.confirmation)
    assert result.status == "sent" and result.sent_item_id and len(fake.ows_calls) == 1


async def test_unclear_send_without_a_sent_copy_is_unknown_and_not_retried(
    writes: Writes, fake: FakeGraph
) -> None:
    proposal = await writes.propose(message())
    fake.ows_next = ["no-answer"]
    result = await writes.send(message(), proposal.confirmation)
    assert result.status == "unknown" and "Do not send again" in result.detail
    assert len(fake.ows_calls) == 1


async def test_failed_sent_items_check_keeps_the_unknown_status(
    writes: Writes, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    proposal = await writes.propose(message())
    fake.ows_next = ["done-no-answer"]

    async def unavailable(**_: Any) -> Any:
        raise Throttled("Graph is throttling")

    monkeypatch.setattr(writes.mailbox.reader, "list_messages", unavailable)
    result = await writes.send(message(), proposal.confirmation)
    assert result.status == "unknown" and "Do not send again" in result.detail


async def test_sent_items_check_matches_every_recipient_including_bcc(
    writes: Writes, fake: FakeGraph
) -> None:
    bcc_only = message(to=[], bcc=["boss@example.com"])
    proposal = await writes.propose(bcc_only)
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    fake.add(FakeMessage("other", "Hi", "f-sent", now, sender="me@example.com", to=("x@example.com",)))
    fake.ows_next = ["no-answer"]  # not sent; an unrelated mail with the same subject is in Sent Items
    assert (await writes.send(bcc_only, proposal.confirmation)).status == "unknown"
    fake.ows_next = ["done-no-answer"]  # sent this time
    result = await writes.send(bcc_only, proposal.confirmation)
    assert result.status == "sent" and fake.messages[result.sent_item_id or ""].bcc == ("boss@example.com",)


# ---------------------------------------------------------------- send a reply (W6)


def reply(**changes: Any) -> OutgoingMessage:
    return OutgoingMessage.model_validate({"reply_to_message_id": "m3", "body": "Thanks"} | changes)


async def test_reply_is_sent_as_the_checked_draft(writes: Writes, fake: FakeGraph) -> None:
    proposal = await writes.propose(reply())
    result = await writes.send(reply(), proposal.confirmation)
    assert result.status == "sent"
    actions = [(action, body["MessageDisposition"]) for action, body in fake.ows_calls]
    assert actions == [("CreateItem", "SaveOnly"), ("UpdateItem", "SendAndSaveCopy")]
    (sent_id,) = fake.sent_drafts
    sent = fake.messages[sent_id]
    assert sent.folder == "f-sent" and not sent.is_draft and sent.conversation == "c-rel"
    assert "Follow-up with numbers" in sent.text  # the original is quoted
    assert [a.name for a in sent.attachments] == ["image001.png", "logo.png"]  # inline images kept


async def test_reply_draft_without_the_original_is_not_sent(writes: Writes, fake: FakeGraph) -> None:
    fake.reply_drops_history = True
    proposal = await writes.propose(reply())
    with pytest.raises(ConnectorError, match="Not sent: the reply draft did not keep the original") as caught:
        await writes.send(reply(), proposal.confirmation)
    assert not fake.sent_drafts and [action for action, _ in fake.ows_calls] == ["CreateItem"]
    drafts = [m for m in fake.messages.values() if m.is_draft]
    assert len(drafts) == 1 and drafts[0].folder == "f-drafts" and drafts[0].id in str(caught.value)


async def test_reply_missing_an_inline_image_is_not_sent(
    writes: Writes, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    proposal = await writes.propose(reply())
    created = fake.ows_CreateItem

    def drop_images(body: dict[str, Any]) -> list[dict[str, Any]]:
        items = created(body)
        for m in fake.messages.values():
            if m.is_draft:
                m.attachments = []
        return items

    monkeypatch.setattr(fake, "ows_CreateItem", drop_images)
    with pytest.raises(ConnectorError, match="2 inline image\\(s\\) of the original are missing"):
        await writes.send(reply(), proposal.confirmation)
    assert not fake.sent_drafts


async def test_unclear_reply_send_that_went_through_is_found_in_sent_items(
    writes: Writes, fake: FakeGraph
) -> None:
    proposal = await writes.propose(reply())
    fake.ows_next = [None, "done-no-answer"]  # the draft is saved; the send goes through, unanswered
    result = await writes.send(reply(), proposal.confirmation)
    assert result.status == "sent" and result.sent_item_id == fake.sent_drafts[0]
    assert [action for action, _ in fake.ows_calls] == ["CreateItem", "UpdateItem"]  # not retried


async def test_changed_inline_image_bytes_are_not_sent(writes: Writes, fake: FakeGraph) -> None:
    proposal = await writes.propose(reply())
    created = fake.ows_CreateItem

    def change_image(body: dict[str, Any]) -> list[dict[str, Any]]:
        items = created(body)
        draft = next(m for m in fake.messages.values() if m.is_draft)
        first = draft.attachments[0]
        draft.attachments[0] = FakeAttachment(first.id, first.name, b"other bytes", first.content_type,
                                              inline=True, content_id=first.content_id)  # fmt: skip
        return items

    fake.ows_CreateItem = change_image  # type: ignore[method-assign]
    with pytest.raises(ConnectorError, match="missing or changed"):
        await writes.send(reply(), proposal.confirmation)
    assert not fake.sent_drafts


async def test_reply_draft_reports_the_history_check(writes: Writes, fake: FakeGraph) -> None:
    intact = await writes.create_draft(reply())
    assert intact.history_intact is True and intact.history_problem is None
    fake.reply_drops_history = True
    broken = await writes.create_draft(reply())
    assert broken.history_intact is False and "quoted original" in (broken.history_problem or "")
    plain = await writes.create_draft(message())
    assert plain.history_intact is None  # not a reply
    assert not fake.sent_drafts


async def test_reply_draft_that_lost_formatting_is_not_sent(writes: Writes, fake: FakeGraph) -> None:
    fake.messages["m3"].html = '<p>Follow-up <b>with</b></p><ul><li>numbers</li></ul><img src="cid:img1">'
    fake.reply_flattens_html = True  # same words and images; the list, bold and image reference are gone
    proposal = await writes.propose(reply())
    with pytest.raises(ConnectorError, match=r"lost formatting \(image reference, img, li, ul\)"):
        await writes.send(reply(), proposal.confirmation)
    assert not fake.sent_drafts


async def test_reply_html_body_shows_the_confirmed_text_exactly(writes: Writes, fake: FakeGraph) -> None:
    await writes.create_draft(OutgoingMessage(reply_to_message_id="m1", body="  indented\na  b\tc"))
    value = fake.ows_calls[0][1]["Items"][0]["NewBodyContent"]["Value"]
    assert value == "<div>&nbsp; indented<br>a&nbsp; b&nbsp;&nbsp;&nbsp;&nbsp;c</div>"

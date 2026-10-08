from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx
import pytest

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import (
    AccountMismatch,
    InvalidRequest,
    NotFound,
    Throttled,
    Upstream,
)
from outlook_connector.domain.models import OutgoingMessage
from outlook_connector.remote.cloud_settings import CloudSettings
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.ows import Ows
from outlook_connector.remote.ows_mail import OwsMailWriter
from outlook_connector.remote.transport import Transport
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.signatures import Signatures
from outlook_connector.service.writes import Writes
from outlook_connector.store.db import Store
from tests.fakes.graph_fake import FakeAttachment, FakeGraph, StaticTokens, sample_mailbox

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
    signatures = Signatures(CloudSettings(transport, tokens), account)
    return Writes(mailbox, OwsMailWriter(Ows(transport, tokens)), account, signatures)


@pytest.fixture
def writes(fake: FakeGraph, tmp_path: Path) -> Writes:
    return make_writes(fake, tmp_path)


def message(**changes: Any) -> OutgoingMessage:
    return OutgoingMessage.model_validate(
        {"to": ["bob@example.com"], "subject": "Hi", "text_body": "Hello"} | changes
    )


def reply(**changes: Any) -> OutgoingMessage:
    return OutgoingMessage.model_validate({"reply_to_message_id": "m3", "text_body": "Thanks"} | changes)


@pytest.mark.parametrize(
    ("changes", "error"),
    [
        ({"to": []}, "at least one recipient"),
        ({"to": ["bad"]}, "Not an email address"),
        ({"to": ["Bob <bob@example.com>"]}, "Not an email address"),
        ({"cc": ["bob@example.com"]}, "listed twice"),
        ({"subject": " "}, "needs a subject"),
        ({"text_body": " "}, "body is empty"),
        ({"text_body": None}, "exactly one"),
        ({"html_body": "<b>x</b>"}, "exactly one"),
        ({"reply_all": True}, "reply_all needs"),
        ({"to": [f"u{i}@example.com" for i in range(101)]}, "At most 100"),
        ({"subject": "x" * 256}, "255"),
        ({"text_body": "x" * 100001}, "100000"),
    ],
)
async def test_invalid_messages_are_refused(
    writes: Writes, fake: FakeGraph, changes: dict[str, Any], error: str
) -> None:
    with pytest.raises(InvalidRequest, match=error):
        await writes.create_draft(message(**changes))
    assert not fake.ows_calls


async def test_draft_readback_and_recipient_dedup(writes: Writes, fake: FakeGraph) -> None:
    draft = await writes.create_draft(
        message(cc=[" Carol@Example.com ", "carol@example.com"], bcc=["boss@example.com"])
    )
    assert draft.status == "saved" and draft.verified and draft.text_body == "Hello"
    assert draft.html_body == "<div>Hello</div>" and draft.message
    assert draft.message.cc[0].address == "Carol@Example.com"
    assert draft.message.bcc[0].address == "boss@example.com"
    assert fake.messages[draft.id].is_draft and not fake.sent_drafts


async def test_reply_defaults_and_self_replies(writes: Writes, fake: FakeGraph) -> None:
    fake.messages["m1"].to = ("me@example.com", "dan@example.com")
    fake.messages["m1"].cc = ("erin@example.com", "me@example.com")
    draft = await writes.create_draft(
        OutgoingMessage(reply_to_message_id="m1", text_body="Thanks", reply_all=True)
    )
    saved = fake.messages[draft.id]
    assert saved.to == ("alice@example.com", "dan@example.com") and saved.cc == ("erin@example.com",)
    assert saved.subject == "RE: Relatório de exemplo semanal" and saved.conversation == "c-rel"
    own = await writes.create_draft(OutgoingMessage(reply_to_message_id="m2", text_body="Ping"))
    assert fake.messages[own.id].to == ("alice@example.com",)
    fake.messages["m5"].sender = "Me@Example.com"
    fake.messages["m5"].to = ("me@example.com",)
    for reply_all in (False, True):
        own = await writes.create_draft(
            OutgoingMessage(reply_to_message_id="m5", text_body="Ping", reply_all=reply_all)
        )
        assert fake.messages[own.id].to == ("Me@Example.com",)


async def test_missing_reply(writes: Writes) -> None:
    with pytest.raises(NotFound):
        await writes.create_draft(OutgoingMessage(reply_to_message_id="gone", text_body="x"))


async def test_plain_text_preserves_spacing_and_nbsp(writes: Writes, fake: FakeGraph) -> None:
    draft = await writes.create_draft(message(text_body="  indented\na  b\tc\n<>&\u00a0 end"))
    assert (
        draft.html_body
        == "<div>&nbsp; indented<br>a&nbsp; b&nbsp;&nbsp;&nbsp;&nbsp;c<br>&lt;&gt;&amp;&nbsp; end</div>"
    )
    assert fake.ows_calls[0][1]["Items"][0]["Body"]["BodyType"] == "HTML"


@pytest.mark.parametrize(
    "page",
    [
        "<b>Hello</b>",
        "<html><head><style>p {color:red}</style></head><body><p>x</p></body></html>",
        '<img src="https://example.com/p.png"><p style="margin:0">x',
        '<a href="https://example.com">x</a>',
    ],
)
async def test_intentional_html_passes_through(writes: Writes, fake: FakeGraph, page: str) -> None:
    draft = await writes.create_draft(message(text_body=None, html_body=page))
    assert draft.html_body == page
    assert fake.ows_calls[0][1]["Items"][0]["Body"]["Value"] == page


@pytest.mark.parametrize(
    "page",
    [
        "<script>x</script>",
        '<iframe src="x">',
        "<object>",
        "<embed/>",
        "<form>",
        '<img onerror="alert(1)">',
        '<a href="java&#x73;cript:alert(1)">x</a>',
        '<a href="java\nscript:alert(1)">x</a>',
    ],
)
async def test_active_html_rejected(writes: Writes, fake: FakeGraph, page: str) -> None:
    with pytest.raises(InvalidRequest, match="Active HTML"):
        await writes.create_draft(message(text_body=None, html_body=page))
    assert not fake.ows_calls


async def test_send_uses_only_the_stored_draft(writes: Writes, fake: FakeGraph) -> None:
    draft = await writes.create_draft(reply())
    saved = fake.messages[draft.id]
    saved.subject = "Edited in Outlook"
    saved.html = "<b>Edited on server</b>"
    saved.text = "Edited on server"
    result = await writes.send_draft(draft.id)
    assert result.status == "sent" and fake.sent_drafts == [draft.id]
    assert saved.subject == "Edited in Outlook" and saved.html == "<b>Edited on server</b>"
    assert fake.ows_calls[-1][1]["ItemChanges"][0]["Updates"] == []
    with pytest.raises(InvalidRequest, match="not an existing"):
        await writes.send_draft(draft.id)


async def test_a_draft_edited_after_the_send_read_it_is_not_sent(
    writes: Writes, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    draft = await writes.create_draft(reply())
    read = writes.mailbox.message

    async def read_then_edited_in_outlook(*args: Any, **kwargs: Any) -> Any:
        message = await read(*args, **kwargs)
        fake.messages[draft.id].subject = "Edited in Outlook meanwhile"
        return message

    monkeypatch.setattr(writes.mailbox, "message", read_then_edited_in_outlook)
    with pytest.raises(Upstream, match="nothing was changed or sent"):
        await writes.send_draft(draft.id)
    assert fake.sent_drafts == [] and fake.messages[draft.id].subject == "Edited in Outlook meanwhile"


async def test_non_drafts_cannot_send(writes: Writes, fake: FakeGraph) -> None:
    with pytest.raises(InvalidRequest):
        await writes.send_draft("m1")
    assert not fake.ows_calls


@pytest.mark.parametrize("script, status", [("done-no-answer", "sent"), ("no-answer", "unknown")])
async def test_unclear_send_never_retried(writes: Writes, fake: FakeGraph, script: str, status: str) -> None:
    draft = await writes.create_draft(message())
    fake.ows_next = [script]
    result = await writes.send_draft(draft.id)
    assert result.status == status and len(fake.ows_calls) == 2
    if status == "unknown":
        assert "Do not send again" in result.detail


async def test_failed_readback_returns_saved_id_and_failure(
    writes: Writes, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    async def unavailable(*_: Any, **__: Any) -> Any:
        raise Throttled("Graph unavailable")

    monkeypatch.setattr(writes.mailbox.reader, "get_message", unavailable)
    draft = await writes.create_draft(message())
    assert draft.status == "failed" and not draft.verified and draft.id in fake.messages
    assert len(fake.ows_calls) == 1


async def test_readback_retries_reads_only(
    writes: Writes, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    original = writes.mailbox.reader.get_message
    calls = 0

    async def delayed(*args: Any, **kwargs: Any) -> Any:
        nonlocal calls
        calls += 1
        if calls == 1:
            raise NotFound("Propagation delay")
        return await original(*args, **kwargs)

    monkeypatch.setattr(writes.mailbox.reader, "get_message", delayed)
    assert (await writes.create_draft(message())).verified
    assert calls == 3 and len(fake.ows_calls) == 1


async def test_account_mismatch_blocks_all_writes(fake: FakeGraph, tmp_path: Path) -> None:
    other = Account(tenant_id="tenant-x", object_id="other", username="me@example.com")
    writes = make_writes(fake, tmp_path, account=other)
    with pytest.raises(AccountMismatch):
        await writes.create_draft(message())
    fake.messages["m1"].is_draft = True
    with pytest.raises(AccountMismatch):
        await writes.send_draft("m1")
    assert not fake.ows_calls


async def test_reply_history_findings(
    writes: Writes, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    intact = await writes.create_draft(reply())
    assert intact.history_intact is True
    fake.reply_drops_history = True
    broken = await writes.create_draft(reply())
    assert broken.history_intact is False and "quoted original" in (broken.history_problem or "")
    fake.reply_drops_history = False
    fake.reply_flattens_html = True
    fake.messages["m3"].html = '<p>Follow-up <b>with</b></p><ul><li>numbers</li></ul><img src="cid:img1">'
    broken = await writes.create_draft(reply())
    assert broken.history_intact is False and "lost formatting" in (broken.history_problem or "")
    fake.reply_flattens_html = False
    created = fake.ows_CreateItem

    def change_image(body: dict[str, Any]) -> list[dict[str, Any]]:
        items = created(body)
        saved = list(fake.messages.values())[-1]
        first = saved.attachments[0]
        saved.attachments[0] = FakeAttachment(
            first.id, first.name, b"changed", first.content_type, inline=True, content_id=first.content_id
        )
        return items

    monkeypatch.setattr(fake, "ows_CreateItem", change_image)
    broken = await writes.create_draft(reply())
    assert broken.history_intact is False and "missing or changed" in (broken.history_problem or "")


async def test_empty_body_and_missing_cid_findings(writes: Writes, fake: FakeGraph) -> None:
    empty = await writes.create_draft(message(text_body=None, html_body="<div></div>"))
    assert empty.verified and any("empty" in f for f in empty.findings)
    missing = await writes.create_draft(message(text_body=None, html_body='<p>x</p><img src="cid:missing">'))
    assert missing.verified and any("missing inline" in f for f in missing.findings)


async def test_failed_check_after_ambiguous_send_stays_unknown(
    writes: Writes, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    draft = await writes.create_draft(message())
    original = writes.mailbox.message
    calls = 0

    async def unavailable_after_validation(*args: Any, **kwargs: Any) -> Any:
        nonlocal calls
        calls += 1
        if calls > 1:
            raise Throttled("Graph unavailable")
        return await original(*args, **kwargs)

    monkeypatch.setattr(writes.mailbox, "message", unavailable_after_validation)
    fake.ows_next = ["done-no-answer"]
    result = await writes.send_draft(draft.id)
    assert result.status == "unknown" and len(fake.ows_calls) == 2


async def test_missing_saved_id_never_creates_again(writes: Writes, fake: FakeGraph) -> None:
    from outlook_connector.domain.errors import WriteOutcomeUnknown

    fake.ows_next = [{"ResponseClass": "Success", "ResponseCode": "NoError", "Items": []}]
    with pytest.raises(WriteOutcomeUnknown, match="draft id"):
        await writes.create_draft(message())
    assert len(fake.ows_calls) == 1

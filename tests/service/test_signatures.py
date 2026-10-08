from __future__ import annotations

import base64
from pathlib import Path

import httpx
import pytest

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import AccountMismatch, InvalidRequest, NotFound, Upstream
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
from tests.fakes.graph_fake import FakeGraph, StaticTokens, sample_mailbox

ME = Account(tenant_id="tenant-x", object_id="user-x", username="me@example.com")


async def _no_sleep(_seconds: float) -> None:
    return None


def make_services(fake: FakeGraph, tmp_path: Path) -> tuple[Signatures, Writes]:
    tokens = StaticTokens()
    transport = Transport(tokens, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep)
    mailbox = Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "signature.sqlite3", "fp"))
    signatures = Signatures(CloudSettings(transport, tokens), ME)
    writer = OwsMailWriter(Ows(transport, tokens))
    return signatures, Writes(mailbox, writer, ME, signatures)


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


@pytest.fixture
def services(fake: FakeGraph, tmp_path: Path) -> tuple[Signatures, Writes]:
    return make_services(fake, tmp_path)


async def test_native_signature_crud_defaults_and_encoded_exact_names(
    services: tuple[Signatures, Writes], fake: FakeGraph
) -> None:
    signatures, _ = services
    name = "Team café + &"
    assert (await signatures.create_signature(name, "<p>Hello <b>there</b></p>")).status == "created"
    assert fake.signature_contents[name] == {"htm": "<p>Hello <b>there</b></p>", "txt": "Hello there"}

    assert (await signatures.set_default_signature(name, "both")).for_type == "both"
    listed = await signatures.list_signatures()
    assert listed.new_default == name and listed.reply_default == name
    assert [(entry.name, entry.readable) for entry in listed.signatures] == [(name, True)]
    contents = await signatures.get_signature(name)
    assert contents.html == "<p>Hello <b>there</b></p>" and contents.text == "Hello there"

    assert (await signatures.update_signature(name, "<p>Updated</p>")).status == "updated"
    assert fake.signature_contents[name] == {"htm": "<p>Updated</p>", "txt": "Updated"}
    assert (await signatures.delete_signature(name)).status == "deleted"
    assert fake.signature_contents == {}
    assert fake.signature_new_default == name and fake.signature_reply_default == name


async def test_names_are_exact_case_sensitive_and_commas_are_rejected(
    services: tuple[Signatures, Writes], fake: FakeGraph
) -> None:
    signatures, _ = services
    await signatures.create_signature("Case", "<p>upper</p>")
    await signatures.create_signature("case", "<p>lower</p>")
    listed = await signatures.list_signatures()
    assert [entry.name for entry in listed.signatures] == ["Case", "case"]
    assert (await signatures.get_signature("Case")).text == "upper"
    assert (await signatures.get_signature("case")).text == "lower"

    before = len(fake.calls)
    with pytest.raises(InvalidRequest, match="commas"):
        await signatures.create_signature("invalid,name", "<p>no</p>")
    assert len(fake.calls) == before


async def test_empty_default_and_include_false_never_select_a_signature(
    services: tuple[Signatures, Writes], fake: FakeGraph
) -> None:
    _, writes = services
    draft = await writes.create_draft(
        OutgoingMessage(to=["bob@example.com"], subject="No signature", text_body="Body")
    )
    assert draft.html_body == "<div>Body</div>"
    assert 'id="Signature"' not in (draft.html_body or "")

    signature_reads_before = sum(path == "/ows/v1/OutlookCloudSettings/settings/" for path in fake.calls)
    no_signature = await writes.create_draft(
        OutgoingMessage(
            to=["bob@example.com"],
            subject="Explicitly suppressed",
            text_body="Body",
            signature="invalid,name",
            include_signature=False,
        )
    )
    assert no_signature.html_body == "<div>Body</div>"
    signature_reads_after = sum(path == "/ows/v1/OutlookCloudSettings/settings/" for path in fake.calls)
    assert signature_reads_after == signature_reads_before
    assert len(fake.ows_calls) == 2


async def test_named_signature_missing_or_unreadable_fails_before_draft_write(
    services: tuple[Signatures, Writes], fake: FakeGraph
) -> None:
    signatures, writes = services
    with pytest.raises(NotFound):
        await writes.create_draft(
            OutgoingMessage(to=["bob@example.com"], subject="Missing", text_body="Body", signature="missing")
        )
    assert not fake.ows_calls

    fake.signature_contents["Unreadable"] = {}
    unreadable_list = await signatures.list_signatures()
    assert [(entry.name, entry.readable) for entry in unreadable_list.signatures] == [("Unreadable", False)]
    fake.signature_new_default = "Unreadable"
    with pytest.raises(Upstream, match="no readable"):
        await writes.create_draft(
            OutgoingMessage(to=["bob@example.com"], subject="Unreadable", text_body="Body")
        )
    assert not fake.ows_calls

    fake.signature_new_default = "Deleted"
    with pytest.raises(Upstream, match="configured native signature default"):
        await writes.create_draft(
            OutgoingMessage(to=["bob@example.com"], subject="Dangling", text_body="Body")
        )
    assert not fake.ows_calls


async def test_native_default_is_resolved_again_for_replies_and_changes_during_read_fail(
    services: tuple[Signatures, Writes], fake: FakeGraph
) -> None:
    signatures, writes = services
    fake.signature_contents.update(
        {
            "New": {"htm": "<p>New default</p>", "txt": "New default"},
            "Reply": {"htm": "<p>Reply default</p>", "txt": "Reply default"},
        }
    )
    fake.signature_new_default = "New"
    fake.signature_reply_default = "Reply"

    new_draft = await writes.create_draft(
        OutgoingMessage(to=["bob@example.com"], subject="New", text_body="Body")
    )
    assert (new_draft.html_body or "").index('data-signature-name="New"') > (new_draft.html_body or "").index(
        "Body"
    )

    fake.signature_new_default = "Reply"
    changed_default = await writes.create_draft(
        OutgoingMessage(to=["bob@example.com"], subject="Changed default", text_body="Body")
    )
    assert 'data-signature-name="Reply"' in (changed_default.html_body or "")

    reply = await writes.create_draft(OutgoingMessage(reply_to_message_id="m3", text_body="Reply body"))
    page = reply.html_body or ""
    assert page.index('data-signature-name="Reply"') < page.index("<div><b>From:</b>")
    assert "quoted" not in page

    fake.signature_change_after_contents = True
    before = len(fake.ows_calls)
    with pytest.raises(Upstream, match="changed while"):
        await signatures.get_signature("New")
    assert len(fake.ows_calls) == before


async def test_data_image_becomes_inline_attachment_and_is_read_back_byte_exactly(
    services: tuple[Signatures, Writes], fake: FakeGraph
) -> None:
    signatures, writes = services
    image = b"\x89PNG\r\nsynthetic image bytes"
    data_uri = base64.b64encode(image).decode("ascii")
    html = f'<p>With logo</p><img alt="logo" src="data:image/png;base64,{data_uri}">'
    await signatures.create_signature("Logo", html)
    await signatures.set_default_signature("Logo", "new")

    draft = await writes.create_draft(
        OutgoingMessage(to=["bob@example.com"], subject="Logo", text_body="Body")
    )
    saved = fake.messages[draft.id]
    assert "data:image" not in (draft.html_body or "")
    assert "cid:signature-" in (draft.html_body or "")
    (attachment,) = saved.attachments
    assert attachment.inline and attachment.content_type == "image/png"
    assert attachment.data == image and attachment.content_id
    assert not any("signature image" in finding for finding in draft.findings)
    payload = fake.ows_calls[-1][1]["Items"][0]["Attachments"][0]
    assert base64.b64decode(payload["Content"]) == image

    fake.signature_reply_default = "Logo"
    reply = await writes.create_draft(OutgoingMessage(reply_to_message_id="m3", text_body="Reply with logo"))
    reply_page = reply.html_body or ""
    assert reply_page.index('data-signature-name="Logo"') < reply_page.index("<div><b>From:</b>")
    assert reply.history_intact is True
    assert not any("signature image" in finding for finding in reply.findings)
    assert any(item.inline and item.data == image for item in fake.messages[reply.id].attachments)


async def test_text_only_native_signature_is_escaped_into_html(
    services: tuple[Signatures, Writes], fake: FakeGraph
) -> None:
    _, writes = services
    fake.signature_contents["Plain"] = {"txt": "A < B\nSecond  line\tend"}
    fake.signature_new_default = "Plain"
    draft = await writes.create_draft(
        OutgoingMessage(to=["bob@example.com"], subject="Text", text_body="Body")
    )
    assert "A &lt; B<br>Second&nbsp; line&nbsp;&nbsp;&nbsp;&nbsp;end" in (draft.html_body or "")


async def test_write_account_must_match_the_bound_read_account(fake: FakeGraph) -> None:
    tokens = StaticTokens()
    transport = Transport(tokens, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep)
    wrong = Account(tenant_id="tenant-x", object_id="different-user", username="me@example.com")
    signatures = Signatures(CloudSettings(transport, tokens), wrong)
    with pytest.raises(AccountMismatch):
        await signatures.list_signatures()
    assert not fake.calls

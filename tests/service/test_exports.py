from __future__ import annotations

import zipfile
from datetime import UTC, datetime
from pathlib import Path

import httpx
import pytest
from pydantic import ValidationError

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import ExportRequest, Recipient
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.service.export.attachments import dedupe, safe_name
from outlook_connector.service.export.formatter import people
from outlook_connector.service.export.orchestrator import DELETED_OR_JUNK, NOT_RECEIVED, Exports
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.threads import Threads
from outlook_connector.store.db import Store
from tests.fakes.graph_fake import FakeAttachment, FakeGraph, FakeMessage, StaticTokens, sample_mailbox


async def _no_sleep(_s: float) -> None:
    return None


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


@pytest.fixture
def exports(fake: FakeGraph, tmp_path: Path) -> Exports:
    transport = Transport(
        StaticTokens(), client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep
    )
    return Exports(Threads(Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "m.sqlite3", "fp"))))


def zip_names(path: str) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        return sorted(archive.namelist())


def zip_text(path: str, name: str) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


async def test_thread_with_attachments_is_one_zip_with_sibling_folder(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"], include_attachments=True))
    assert artifact.filename.endswith(".zip") and artifact.message_count == 3
    stem = "2026-09-28 Relatório BE semanal"
    # numbers.xlsx (regular) and image001.png (inline, referenced by the unique body) are in;
    # logo.png (inline signature, not referenced) is out; junk m4 is excluded by default.
    assert zip_names(artifact.path) == [f"{stem}.txt", f"{stem}/image001.png", f"{stem}/numbers.xlsx"]
    text = zip_text(artifact.path, f"{stem}.txt")
    assert "Attachment: 2026-09-28 Relatório BE semanal/numbers.xlsx" in text
    assert text.index("First report") < text.index("Thanks!") < text.index("Follow-up with numbers")
    assert "> First report" not in text  # unique body by default
    assert "buy now" not in text


async def test_without_attachments_a_single_thread_is_a_flat_txt(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"]))
    assert artifact.filename == "2026-09-28 Relatório BE semanal.txt" and artifact.content_type.startswith(
        "text/plain"
    )
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert "Attachment: numbers.xlsx" in text and "image001.png" not in text  # inline images are not listed


async def test_full_body_keeps_quoted_history(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(message_ids=["m2"], body="full"))
    assert "> First report" in Path(artifact.path).read_text(encoding="utf-8")


async def test_combine_all_is_one_txt_with_sections(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel", "c-lunch"], combine="all"))
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert artifact.filename.endswith(".txt") and text.count("### Conversation:") == 2


async def test_combine_none_is_one_file_per_message_in_a_zip(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(message_ids=["m1", "m5"], combine="none"))
    assert zip_names(artifact.path) == ["2026-09-28 Relatório BE semanal.txt", "2026-09-30 Lunch_.txt"]


async def test_selected_messages_join_their_thread_file(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-lunch"], message_ids=["m1", "m3"]))
    assert artifact.text_files == 2 and artifact.message_count == 3


async def test_failed_download_becomes_a_marker(exports: Exports, fake: FakeGraph) -> None:
    fake.messages["m3"].attachments[0].broken = True
    artifact = await exports.export(ExportRequest(message_ids=["m3"], include_attachments=True))
    assert artifact.attachments_unavailable == 1
    text = zip_text(artifact.path, "2026-09-29 Relatório BE semanal.txt")
    assert "[Attachment unavailable: numbers.xlsx]" in text


async def test_forwarded_mail_attachment_is_saved_as_eml(exports: Exports, fake: FakeGraph) -> None:
    fake.add(
        FakeMessage(
            "fw",
            "FW: thing",
            "f-inbox",
            "2026-09-30T08:00:00Z",
            conversation="c-fw",
            attachments=[
                FakeAttachment("i1", "Original message", b"MIME", "message/rfc822", kind="itemAttachment")
            ],
        )
    )
    artifact = await exports.export(ExportRequest(message_ids=["fw"], include_attachments=True))
    assert "2026-09-30 thing/Original message.eml" in zip_names(artifact.path)


async def test_server_deleted_message_is_exported_from_retention_and_labelled(
    exports: Exports, fake: FakeGraph
) -> None:
    await exports.export(ExportRequest(conversation_ids=["c-rel"]))
    del fake.messages["m2"]
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"]))
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert "!! DELETED on the server" in text and "Thanks!" in text


async def test_inline_only_attachments_are_found_when_exporting_files(
    exports: Exports, fake: FakeGraph
) -> None:
    fake.add(
        FakeMessage(
            "pic",
            "Chart",
            "f-inbox",
            "2026-09-30T08:00:00Z",
            conversation="c-pic",
            attachments=[
                FakeAttachment("p1", "chart.png", b"png", "image/png", inline=True, content_id="c1")
            ],
            html='<img src="cid:c1">',
        )
    )
    artifact = await exports.export(ExportRequest(message_ids=["pic"], include_attachments=True))
    assert "2026-09-30 Chart/chart.png" in zip_names(artifact.path)


async def test_cached_message_deleted_on_the_server_is_labelled(exports: Exports, fake: FakeGraph) -> None:
    await exports.export(ExportRequest(message_ids=["m3"]))  # caches m3's summary and body
    del fake.messages["m3"]
    artifact = await exports.export(ExportRequest(message_ids=["m3"], include_attachments=True))
    text = Path(artifact.path).read_text(encoding="utf-8")  # nothing downloadable: a flat .txt
    assert "!! DELETED on the server" in text and "Follow-up with numbers" in text
    assert "[Attachment unavailable: numbers.xlsx]" in text


async def test_truncated_conversation_is_not_exported_silently(
    exports: Exports, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("outlook_connector.remote.graph_mail.MAX_CONVERSATION", 2)
    with pytest.raises(InvalidRequest, match="listing limit"):
        await exports.export(ExportRequest(conversation_ids=["c-rel"]))


async def test_identical_exports_in_the_same_second_get_distinct_files(exports: Exports) -> None:
    first = await exports.export(ExportRequest(message_ids=["m1"]))
    second = await exports.export(ExportRequest(message_ids=["m1"]))
    assert first.path != second.path and Path(first.path).exists() and Path(second.path).exists()


async def test_each_message_carries_its_source_ids(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(message_ids=["m1"]))
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert "Message id:      m1" in text and "Conversation id: c-rel" in text
    assert "Internet id:     <m1@example.com>" in text


async def test_range_export_selects_the_window_and_counts_what_it_leaves_out(exports: Exports) -> None:
    artifact = await exports.export(
        ExportRequest(
            since=datetime(2026, 9, 28, tzinfo=UTC), until=datetime(2026, 9, 29, 23, 59, tzinfo=UTC)
        )
    )
    assert artifact.message_count == 3  # m1, m2, m3; junk m4 left out, m5 outside the window
    assert artifact.messages_excluded == {DELETED_OR_JUNK: 1} and artifact.messages_unavailable == 0
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert f"Left out: 1 message(s) {DELETED_OR_JUNK}." in text and "buy now" not in text


async def test_range_export_received_only_and_deleted_items(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(received_only=True, include_deleted_items=True))
    assert artifact.message_count == 4 and artifact.messages_excluded == {NOT_RECEIVED: 1}  # m2 is sent


async def test_range_export_combines_with_explicit_ids(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(folder="inbox/projects/rie", message_ids=["m5"]))
    assert artifact.message_count == 2


async def test_export_limit_is_enforced_with_counts(exports: Exports) -> None:
    with pytest.raises(InvalidRequest, match="holds 3 messages, above the limit of 2"):
        await exports.export(ExportRequest(received_only=True, limit=2))
    with pytest.raises(InvalidRequest, match="above the limit of 1"):
        await exports.export(ExportRequest(message_ids=["m1", "m5"], limit=1))


def test_export_limit_cannot_exceed_the_hard_cap() -> None:
    with pytest.raises(ValidationError):
        ExportRequest(message_ids=["m1"], limit=2001)


async def test_unfetchable_bodies_are_marked_and_counted_not_fatal(exports: Exports, fake: FakeGraph) -> None:
    await exports.mailbox.list_messages()  # ids come from a listing, so their summaries are known
    fake.throttle_items = 10_000  # every body sub-request stays throttled
    artifact = await exports.export(ExportRequest(message_ids=["m5"], combine="all"))
    assert artifact.messages_unavailable == 1 and artifact.unavailable_message_ids == ["m5"]
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert "(Content unavailable: While fetching message bodies" in text
    assert "Unavailable: 1 message body(ies); they are marked below." in text


async def test_deleted_before_ever_read_counts_as_unavailable(exports: Exports, fake: FakeGraph) -> None:
    await exports.mailbox.list_messages()  # only summaries are stored, no bodies
    del fake.messages["m5"]
    artifact = await exports.export(ExportRequest(message_ids=["m5"]))
    assert artifact.messages_unavailable == 1 and artifact.unavailable_message_ids == ["m5"]
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert "(Content unavailable: deleted on the server and never retained by this app)" in text


async def test_bulk_export_stays_within_batch_limits_under_throttling(
    exports: Exports, fake: FakeGraph
) -> None:
    for i in range(150):
        fake.add(
            FakeMessage(
                f"b{i:03d}",
                f"Bulk {i}",
                "f-inbox",
                f"2026-08-{1 + i % 28:02d}T10:{i % 60:02d}:00Z",
                conversation=f"cb{i}",
            )
        )
    ids = [f"b{i:03d}" for i in range(150)]
    await exports.export(ExportRequest(message_ids=ids[:1]))  # warm the folder cache
    fake.throttle_items = 60  # several batches throttled at once
    artifact = await exports.export(ExportRequest(message_ids=ids, combine="all"))
    assert artifact.message_count == 150 and artifact.messages_unavailable == 0
    assert max(fake.batch_sizes) <= 20


async def test_empty_request_is_rejected(exports: Exports) -> None:
    with pytest.raises(InvalidRequest):
        await exports.export(ExportRequest())


def test_recipients_are_separated_by_semicolons() -> None:
    # display names are often "Last, First", so commas cannot separate people
    names = [Recipient(name="Rhode, Lucas", address="lr@example.com"), Recipient(name="Garcia, Felipe")]
    assert people(names) == "Rhode, Lucas <lr@example.com>; Garcia, Felipe"


def test_safe_names() -> None:
    assert safe_name("a/b:c*?.pdf", fallback="x") == "a_b_c__.pdf"
    assert safe_name("CON.txt", fallback="x") == "_CON.txt"
    assert safe_name("   ", fallback="file") == "file"
    taken: set[str] = set()
    assert [dedupe(n, taken) for n in ("r.pdf", "R.pdf", "r.pdf")] == ["r.pdf", "R (2).pdf", "r (3).pdf"]

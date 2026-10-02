from __future__ import annotations

import zipfile
from pathlib import Path

import httpx
import pytest

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.domain.models import ExportRequest
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.service.export.attachments import dedupe, safe_name
from outlook_connector.service.export.orchestrator import Exports
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


async def test_empty_request_is_rejected(exports: Exports) -> None:
    with pytest.raises(InvalidRequest):
        await exports.export(ExportRequest())


def test_safe_names() -> None:
    assert safe_name("a/b:c*?.pdf", fallback="x") == "a_b_c__.pdf"
    assert safe_name("CON.txt", fallback="x") == "_CON.txt"
    assert safe_name("   ", fallback="file") == "file"
    taken: set[str] = set()
    assert [dedupe(n, taken) for n in ("r.pdf", "R.pdf", "r.pdf")] == ["r.pdf", "R (2).pdf", "r (3).pdf"]

from __future__ import annotations

import asyncio
import json
import zipfile
from datetime import UTC, datetime
from pathlib import Path

import httpx
import pytest
from pydantic import ValidationError

from outlook_connector.domain.errors import InvalidRequest, NotFound, Throttled
from outlook_connector.domain.models import EXCLUSION_TEXT, ExportRequest, Recipient
from outlook_connector.remote.graph import Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.service.conversations import Conversations
from outlook_connector.service.export.attachments import dedupe, safe_name
from outlook_connector.service.export.formatter import people
from outlook_connector.service.export.orchestrator import Exports
from outlook_connector.service.files import Files
from outlook_connector.service.localfiles import kept_dir
from outlook_connector.service.mailbox import Mailbox
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
    return Exports(
        Conversations(Mailbox(GraphMailReader(Graph(transport)), Store(tmp_path / "m.sqlite3", "fp")))
    )


def zip_names(path: str) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        return sorted(archive.namelist())


def zip_text(path: str, name: str) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read(name).decode("utf-8")


async def test_conversation_with_attachments_is_one_zip_with_sibling_folder(exports: Exports) -> None:
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


async def test_without_attachments_a_single_conversation_is_a_flat_txt(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"]))
    assert artifact.filename == "2026-09-28 Relatório BE semanal.txt" and artifact.content_type.startswith(
        "text/plain"
    )
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert "Attachment: numbers.xlsx" in text and "image001.png" not in text  # inline images are not listed


async def test_full_body_keeps_quoted_history(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(message_ids=["m2"], body="full"))
    assert "> First report" in Path(artifact.path).read_text(encoding="utf-8")


async def test_a_download_that_times_out_is_marked_not_fatal(exports: Exports, fake: FakeGraph) -> None:
    fake.drop_downloads = 100  # every download attempt loses its connection (read timeout)
    artifact = await exports.export(ExportRequest(message_ids=["m3"], include_attachments=True))
    assert artifact.export_errors == {"downloading an attachment": 2}  # numbers.xlsx, image001.png
    text = Path(artifact.path).read_text(encoding="utf-8")  # no file was downloaded: a flat TXT
    assert "[EXPORT ERROR] The attachment numbers.xlsx could not be downloaded.\n" in text
    assert "  Likely: Microsoft service or network problem\n" in text


async def test_combine_all_is_one_txt_with_sections(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel", "c-lunch"], combine="all"))
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert artifact.filename.endswith(".txt") and text.count("### Conversation:") == 2


async def test_combine_none_is_one_file_per_message_in_a_zip(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(message_ids=["m1", "m5"], combine="none"))
    assert zip_names(artifact.path) == ["2026-09-28 Relatório BE semanal.txt", "2026-09-30 Lunch_.txt"]


async def test_selected_messages_join_their_conversation_file(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-lunch"], message_ids=["m1", "m3"]))
    assert artifact.text_files == 2 and artifact.message_count == 3


async def test_failed_download_becomes_an_export_error_block(exports: Exports, fake: FakeGraph) -> None:
    fake.messages["m3"].attachments[0].broken = True  # its download answers 503
    artifact = await exports.export(ExportRequest(message_ids=["m3"], include_attachments=True))
    assert artifact.export_errors == {"downloading an attachment": 1}
    text = zip_text(artifact.path, "2026-09-29 Relatório BE semanal.txt")
    assert "[EXPORT ERROR] The attachment numbers.xlsx could not be downloaded.\n" in text
    assert "  Step:   downloading an attachment\n" in text
    assert "  Error:  HTTP 503 ServiceUnavailable, request-id req-" in text
    assert "  Likely: Microsoft throttled the mailbox" in text
    assert "  Fix:    export it again in a few minutes" in text
    summary = (
        "Export errors: 1 attachment (1 throttled) could not be exported; "
        "they are marked [EXPORT ERROR] below."
    )
    assert summary in text and artifact.error_summary == summary


async def test_failed_attachment_listing_becomes_an_export_error_block(
    exports: Exports, fake: FakeGraph
) -> None:
    fake.fail[r"/me/messages/m3/attachments"] = 500
    artifact = await exports.export(ExportRequest(message_ids=["m3"], format="jsonl"))
    record = json.loads(Path(artifact.path).read_text(encoding="utf-8"))
    error = record["attachments_export_error"]
    assert error["step"] == "listing attachments" and error["status"] == 500 and error["retry"] is True
    assert error["likely_cause"] == "Microsoft service or network problem" and record["body"]
    assert "export_error" not in record  # the body is there
    txt = await exports.export(ExportRequest(message_ids=["m3"]))
    text = Path(txt.path).read_text(encoding="utf-8")
    assert "[EXPORT ERROR] The attachments of this message could not be listed." in text
    assert "the attachments of 1 message (1 service or network problem)" in text


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


async def test_server_deleted_message_is_gone_from_the_export(exports: Exports, fake: FakeGraph) -> None:
    await exports.export(ExportRequest(conversation_ids=["c-rel"]))
    del fake.messages["m2"]
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"]))
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert artifact.message_count == 2 and "Thanks!" not in text and "DELETED" not in text


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


async def test_inline_image_ids_of_many_messages_are_read_in_shared_batches(
    exports: Exports, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    for n in range(12):  # a signature image on every message, as in real mail
        fake.add(
            FakeMessage(
                f"sig{n}",
                f"Note {n}",
                "f-inbox",
                f"2026-09-30T08:{n:02d}:00Z",
                conversation=f"c-sig{n}",
                attachments=[
                    FakeAttachment(f"s{n}a", "logo.png", b"png", "image/png", inline=True, content_id="L"),
                    FakeAttachment(f"s{n}b", "pic.png", b"pic", "image/png", inline=True, content_id="P"),
                ],
                html='<img src="cid:P">',
            )
        )
    lookups: list[int] = []
    original = fake.batch

    def recording(body):
        urls = [r["url"] for r in body["requests"]]
        if any("contentId" in u for u in urls):
            lookups.append(len(urls))
        return original(body)

    monkeypatch.setattr(fake, "batch", recording)
    ids = [f"sig{n}" for n in range(12)]
    artifact = await exports.export(ExportRequest(message_ids=ids, include_attachments=True, combine="all"))
    assert sorted(lookups) == [4, 20]  # 24 attachments of 12 messages: two batches, not twelve
    assert sum(name.endswith("pic.png") for name in zip_names(artifact.path)) == 1  # identical bytes, once


HINT = "Refresh the list (or list/search again) and retry the export. Nothing was exported."


def exported_files() -> list[Path]:
    folder = kept_dir("exports")
    return sorted(folder.iterdir()) if folder.exists() else []


async def test_message_deleted_before_export_fails_the_export_and_writes_nothing(
    exports: Exports, fake: FakeGraph
) -> None:
    await exports.mailbox.list_messages()
    del fake.messages["m3"]
    with pytest.raises(NotFound) as caught:
        await exports.export(ExportRequest(message_ids=["m1", "m3"], include_attachments=True))
    assert "1 of the 2 message(s) selected by id could not be read: m3 (not found" in str(caught.value)
    assert "deleted or moved in Outlook" in str(caught.value) and HINT in str(caught.value)
    assert exported_files() == []


async def test_message_still_throttled_fails_the_export_as_throttled(
    exports: Exports, fake: FakeGraph
) -> None:
    await exports.mailbox.folders()
    fake.throttle_items = 10_000  # every body sub-request stays throttled
    with pytest.raises(Throttled) as caught:
        await exports.export(ExportRequest(message_ids=["m5"]))
    assert "m5 (failed: HTTP 429 ApplicationThrottled" in str(caught.value) and HINT in str(caught.value)
    assert exported_files() == []


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
    assert artifact.messages_excluded == {"deleted_or_junk": 1}
    text = Path(artifact.path).read_text(encoding="utf-8")
    assert f"Left out: 1 message(s) {EXCLUSION_TEXT['deleted_or_junk']}." in text and "buy now" not in text


async def test_range_export_without_sent_items_but_with_deleted_items(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(include_sent_items=False, include_deleted_items=True))
    assert artifact.message_count == 4 and artifact.messages_excluded == {"outgoing": 1}  # m2 is sent


async def test_range_export_combines_with_explicit_ids(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(folder="inbox/projects/rie", message_ids=["m5"]))
    assert artifact.message_count == 2


async def test_export_limit_is_enforced_with_counts(exports: Exports) -> None:
    with pytest.raises(InvalidRequest, match="holds 3 messages, above the limit of 2"):
        await exports.export(ExportRequest(include_sent_items=False, limit=2))
    with pytest.raises(InvalidRequest, match="above the limit of 1"):
        await exports.export(ExportRequest(message_ids=["m1", "m5"], limit=1))


def test_export_limit_cannot_exceed_the_hard_cap() -> None:
    with pytest.raises(ValidationError):
        ExportRequest(message_ids=["m1"], limit=2001)


async def test_throttled_bodies_are_export_error_blocks_not_fatal(exports: Exports, fake: FakeGraph) -> None:
    await exports.mailbox.folders()
    fake.throttle_items = 10_000  # every body sub-request stays throttled
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"], combine="all"))
    assert len(artifact.unavailable_message_ids) == 3
    assert artifact.export_errors == {"fetching message bodies": 3}
    text = Path(artifact.path).read_text(encoding="utf-8")
    block = (
        "[EXPORT ERROR] The body of this message could not be fetched.\n"
        "  Step:   fetching message bodies\n"
        "  Error:  HTTP 429 ApplicationThrottled: Too many requests.\n"
        "  Likely: Microsoft throttled the mailbox (about 4 parallel requests or 10,000 per 10 minutes); "
        "the message itself is fine\n"
        "  Fix:    export it again in a few minutes"
    )
    assert text.count(block) == 3
    summary = (
        "Export errors: 3 message bodies (3 throttled) could not be exported; "
        "they are marked [EXPORT ERROR] below."
    )
    assert summary in text and artifact.error_summary == summary


async def test_denied_body_is_not_retryable_and_jsonl_carries_the_error(
    exports: Exports, fake: FakeGraph
) -> None:
    fake.fail[r"/me/messages/m2"] = 403
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"], format="jsonl"))
    records = {
        r["id"]: r for r in map(json.loads, Path(artifact.path).read_text(encoding="utf-8").splitlines())
    }
    error = records["m2"]["export_error"]
    assert records["m2"]["body"] is None
    assert error == {
        "step": "fetching message bodies",
        "status": 403,
        "code": "ErrorAccessDenied",
        "message": "Injected failure.",
        "likely_cause": "access denied for this item (for example an encrypted or protected message)",
        "retry": False,
        "fix": "retrying will not help",
    }
    assert "export_error" not in records["m1"] and records["m1"]["body"] == "First report"


async def test_jsonl_failed_attachment_record_carries_its_export_error(
    exports: Exports, fake: FakeGraph
) -> None:
    fake.messages["m3"].attachments[0].broken = True
    artifact = await exports.export(
        ExportRequest(message_ids=["m3"], format="jsonl", include_attachments=True)
    )
    jsonl = next(n for n in zip_names(artifact.path) if n.endswith(".jsonl"))
    record = json.loads(zip_text(artifact.path, jsonl))
    failed = next(a for a in record["attachments"] if a["name"] == "numbers.xlsx")
    assert failed["export_error"]["step"] == "downloading an attachment" and "file" not in failed
    assert failed["export_error"]["status"] == 503 and failed["export_error"]["request_id"].startswith("req-")
    assert all("export_error" not in a for a in record["attachments"] if a["name"] != "numbers.xlsx")


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
    assert artifact.message_count == 150
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


async def test_copies_are_exported_once(exports: Exports, fake: FakeGraph) -> None:
    for mid, folder in (("s1", "f-sent"), ("r1", "f-inbox")):
        fake.add(
            FakeMessage(
                mid, "To myself", folder, "2026-10-01T09:00:00Z", conversation="c-s", internet_id="<s@x>"
            )
        )
    artifact = await exports.export(ExportRequest(message_ids=["s1", "r1"]))
    assert artifact.message_count == 1
    assert "Also in: Sent Items" in Path(artifact.path).read_text(encoding="utf-8")


async def test_jsonl_export_has_one_record_per_message(exports: Exports) -> None:
    artifact = await exports.export(ExportRequest(conversation_ids=["c-rel"], format="jsonl"))
    assert artifact.filename.endswith(".jsonl") and artifact.content_type.startswith("application/x-ndjson")
    records = [json.loads(line) for line in Path(artifact.path).read_text(encoding="utf-8").splitlines()]
    assert [r["id"] for r in records] == ["m1", "m2", "m3"]
    assert records[0]["conversation_id"] == "c-rel" and records[0]["body"] == "First report"
    assert records[2]["attachments"][0]["name"] == "numbers.xlsx" and records[0]["from"]["address"]


async def test_jsonl_export_with_attachment_files_is_a_zip(exports: Exports) -> None:
    artifact = await exports.export(
        ExportRequest(conversation_ids=["c-rel"], format="jsonl", include_attachments=True)
    )
    names = zip_names(artifact.path)
    jsonl = next(n for n in names if n.endswith(".jsonl"))
    records = [json.loads(line) for line in zip_text(artifact.path, jsonl).splitlines()]
    files = [a["file"] for r in records for a in r["attachments"] if "file" in a]
    assert files and all(f in names for f in files)


async def test_downloads_never_share_a_file(exports: Exports) -> None:
    files = Files(exports.mailbox)
    first, second = await asyncio.gather(
        files.download_attachment("m3", "a1"), files.download_attachment("m3", "a1")
    )
    assert first.path != second.path and Path(first.path).read_bytes() == Path(second.path).read_bytes()


async def test_save_mime_names_the_file_after_the_subject_from_one_light_read(
    exports: Exports, fake: FakeGraph
) -> None:
    files = Files(exports.mailbox)
    fake.calls.clear()
    saved = await files.save_mime("m3")
    assert saved.name == "RE_ Relatório BE semanal.eml"
    assert Path(saved.path).read_bytes().startswith(b"Subject:")
    # one summary read (in a $batch) and the download: no body, no attachment listing
    assert fake.calls == ["POST /v1.0/$batch", "GET /v1.0/me/messages/m3/$value"]
    with pytest.raises(NotFound):
        await files.save_mime("gone")


async def test_identical_attachment_files_are_stored_once(exports: Exports, fake: FakeGraph) -> None:
    for mid, day in (("sig1", "01"), ("sig2", "02")):
        fake.add(
            FakeMessage(
                mid, "Status", "f-inbox", f"2026-10-{day}T09:00:00Z", conversation="c-sig",
                attachments=[FakeAttachment(f"{mid}-logo", "logo.png", b"same-logo", "image/png")],
            )
        )  # fmt: skip
    artifact = await exports.export(ExportRequest(conversation_ids=["c-sig"], include_attachments=True))
    stem = "2026-10-01 Status"
    assert zip_names(artifact.path) == [f"{stem}.txt", f"{stem}/logo.png"] and artifact.attachment_files == 1
    assert zip_text(artifact.path, f"{stem}.txt").count(f"Attachment: {stem}/logo.png") == 2


async def test_copies_on_different_pages_are_exported_once_naming_both_folders(
    exports: Exports, fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    for mid, folder, at in (
        ("s9", "f-sent", "2026-10-01T09:00:00Z"),
        ("r9", "f-inbox", "2026-10-01T09:00:01Z"),
    ):
        fake.add(FakeMessage(mid, "Note to self", folder, at, conversation="c-s9", internet_id="<s9@x>"))
    monkeypatch.setattr("outlook_connector.service.export.orchestrator.RANGE_PAGE", 1)  # one message a page
    artifact = await exports.export(ExportRequest(since=datetime(2026, 10, 1, tzinfo=UTC), format="jsonl"))
    records = [json.loads(line) for line in Path(artifact.path).read_text(encoding="utf-8").splitlines()]
    assert [(r["id"], r["also_in"]) for r in records] == [("r9", ["Sent Items"])]


async def test_a_range_export_of_a_junk_heavy_mailbox_never_reads_junk(
    exports: Exports, fake: FakeGraph
) -> None:
    from tests.service.test_mailbox_conversations import _junk_heavy

    _junk_heavy(fake)
    fake.calls.clear()
    artifact = await exports.export(ExportRequest(since=datetime(2026, 9, 1, tzinfo=UTC), format="jsonl"))
    assert artifact.message_count == 12  # m1-m3, m5 and n1-n8; m4 and the 40 junk messages left out
    assert not any(c == "GET /v1.0/me/messages" or "f-junk" in c for c in fake.calls)


async def test_mime_disappearing_after_summary_has_neutral_not_found_text(
    exports: Exports, monkeypatch: pytest.MonkeyPatch
) -> None:
    async def missing(*args, **kwargs):  # type: ignore[no-untyped-def]
        raise NotFound("Synthetic not found")

    monkeypatch.setattr(exports.mailbox.reader, "download_mime", missing)
    with pytest.raises(NotFound, match="MIME source was not found;.*id may be wrong"):
        await Files(exports.mailbox).save_mime("m1")
    assert exported_files() == []

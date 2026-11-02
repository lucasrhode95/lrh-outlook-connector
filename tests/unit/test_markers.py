from __future__ import annotations

from outlook_connector.domain.models import Attachment
from outlook_connector.service.markers import file_names, named, with_markers


def _file(aid: str, name: str | None, *, inline: bool = False, kind: str = "file") -> Attachment:
    return Attachment(
        id=aid,
        message_id="m",
        name=name,
        is_inline=inline,
        kind=kind,  # type: ignore[arg-type]
    )


CHART = _file("a1", "image001.png", inline=True)
REPORT = _file("a2", "budget-report.xlsx")


def test_exchange_text_is_kept_and_attachments_follow_in_listing_order() -> None:
    forwarded = _file("a3", "Weekly update", kind="item")
    link = _file("a4", "plan.docx", kind="reference")
    unknown = _file("a5", "odd", kind="unknown")
    text = "See [cid:image001] below. Exchange wrote [same label]."
    assert with_markers(text, [CHART, REPORT, forwarded, link, unknown]) == (
        "See [cid:image001] below. Exchange wrote [same label].\n\n"
        "[Image: image001.png]\n[File: budget-report.xlsx]\n"
        "[Message: Weekly update.eml]\n[Link: plan.docx]"
    )


def test_placeholders_and_literal_image_markers_are_not_rewritten() -> None:
    text = "[cid:encoded+value] [same label] [Image: prose.png]"
    assert with_markers(text, [CHART]) == f"{text}\n\n[Image: image001.png]"


def test_empty_text_lists_attachments_without_an_extra_separator() -> None:
    assert with_markers("", [CHART, REPORT]) == "[Image: image001.png]\n[File: budget-report.xlsx]"


def test_no_known_attachments_preserves_text() -> None:
    unknown = _file("a6", "odd", kind="unknown")
    text = "Text\n"
    assert with_markers(text, [unknown]) == text


def test_names_are_safe_and_unique_within_the_message() -> None:
    attachments = [
        _file("a1", "report.pdf"),
        _file("a2", "Report.pdf"),
        _file("a3", "a/b:c.txt"),
        _file("a4", None),
        _file("a5", "CON.backup.txt"),
    ]
    assert file_names(attachments) == {
        "a1": "report.pdf",
        "a2": "Report (2).pdf",
        "a3": "a_b_c.txt",
        "a4": "attachment",
        "a5": "_CON.backup.txt",
    }
    assert [a.file_name for a in named(attachments)] == list(file_names(attachments).values())


def test_names_given_by_an_export_replace_the_message_names() -> None:
    assert with_markers("Text", [REPORT], names={"a2": "budget-report (2).xlsx"}) == (
        "Text\n\n[File: budget-report (2).xlsx]"
    )

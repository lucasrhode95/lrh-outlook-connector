from __future__ import annotations

from outlook_connector.domain.models import Attachment
from outlook_connector.service.markers import file_names, named, with_markers


def _file(
    aid: str, name: str | None, *, inline: bool = False, cid: str | None = None, kind: str = "file"
) -> Attachment:
    return Attachment(
        id=aid,
        message_id="m",
        name=name,
        is_inline=inline,
        content_id=cid,
        kind=kind,  # type: ignore[arg-type]
    )


CHART = _file("a1", "image001.png", inline=True, cid="image001.png@synthetic")
REPORT = _file("a2", "budget-report.xlsx")


def test_an_inline_image_takes_the_place_of_its_alt_text_and_files_follow_the_text() -> None:
    text = "See below [chart] and reply.\n\nregards"
    html = '<p>See below <img alt="chart" src="cid:image001.png@synthetic"> and reply.</p><p>regards</p>'
    assert with_markers(text, [CHART, REPORT], html=html) == (
        "See below [Image: image001.png] and reply.\n\nregards\n\n[File: budget-report.xlsx]"
    )


def test_placeholders_are_matched_in_document_order() -> None:
    first = _file("a1", "one.png", inline=True, cid="c1")
    second = _file("a2", "two.png", inline=True, cid="c2")
    text = "[pic] then [pic]"
    html = '<img alt="pic" src="cid:c1"> then <img alt=" pic " src="CID:c2">'
    assert with_markers(text, [first, second], html=html) == "[Image: one.png] then [Image: two.png]"


def test_an_image_without_its_placeholder_is_listed_after_the_text_never_on_another_image() -> None:
    text = "Hello [other words]"
    html = (
        '<p>Hello</p><img src="cid:image001.png@synthetic">'
        '<img alt="missing" src="cid:image001.png@synthetic">'
    )
    assert with_markers(text, [CHART], html=html) == "Hello [other words]\n\n[Image: image001.png]"


def test_without_html_images_keep_their_placeholder_and_unreferenced_images_get_no_marker() -> None:
    logo = _file("a3", "logo.png", inline=True, cid="sig")
    assert (
        with_markers("Hi [chart]", [CHART, logo, REPORT], html=None)
        == "Hi [chart]\n\n[File: budget-report.xlsx]"
    )
    html = '<img alt="chart" src="cid:image001.png@synthetic">'
    assert with_markers("Hi [chart]", [CHART, logo], html=html) == "Hi [Image: image001.png]"


def test_forwarded_messages_and_cloud_links_have_their_own_markers() -> None:
    forwarded = _file("a4", "Weekly update", kind="item")
    link = _file("a5", "plan.docx", kind="reference")
    unknown = _file("a6", "odd", kind="unknown")
    assert with_markers("", [forwarded, link, unknown], html=None) == (
        "[Message: Weekly update.eml]\n[Link: plan.docx]"
    )


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
    assert with_markers("Text", [REPORT], html=None, names={"a2": "budget-report (2).xlsx"}) == (
        "Text\n\n[File: budget-report (2).xlsx]"
    )

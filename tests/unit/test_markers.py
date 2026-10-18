from __future__ import annotations

from outlook_connector.domain.body import ImagePart, TextPart
from outlook_connector.domain.models import Attachment
from outlook_connector.service.export.attachments import wanted
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


def test_inline_image_keeps_its_source_position_and_other_attachment_kinds_follow() -> None:
    forwarded = _file("a3", "Weekly update", kind="item")
    link = _file("a4", "plan.docx", kind="reference")
    unknown = _file("a5", "odd", kind="unknown")
    parts = [TextPart("See "), ImagePart(CHART.content_id), TextPart(" below.")]
    assert with_markers("See [chart] below.", [CHART, REPORT, forwarded, link, unknown], parts=parts) == (
        "See [Image: image001.png] below.\n\n[File: budget-report.xlsx]\n"
        "[Message: Weekly update.eml]\n[Link: plan.docx]"
    )


def test_marker_selection_and_placement_do_not_depend_on_body_text() -> None:
    text = 'cid:image001.png@synthetic [Image: image001.png] <img src="cid:elsewhere">'
    parts = [TextPart("Before "), ImagePart(CHART.content_id), TextPart(" after.")]
    assert with_markers(text, [CHART], parts=parts) == "Before [Image: image001.png] after."


def test_empty_text_lists_attachments_without_an_extra_separator() -> None:
    assert with_markers("", [CHART, REPORT], parts=[ImagePart(CHART.content_id)]) == (
        "[Image: image001.png]\n\n[File: budget-report.xlsx]"
    )


def test_no_known_attachments_preserves_text() -> None:
    unknown = _file("a6", "odd", kind="unknown")
    assert with_markers("Text\n", [unknown]) == "Text\n"


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


def test_inline_export_names_are_resolved_by_attachment_id_at_render_time() -> None:
    assert (
        with_markers(
            "unused text", [CHART], parts=[ImagePart(CHART.content_id)], names={"a1": "stored-image.png"}
        )
        == "[Image: stored-image.png]"
    )


def test_body_without_source_parts_keeps_literal_markers_and_labels_unknown_inline_positions() -> None:
    text = "Someone wrote [Image: image001.png] and cid:image001.png@synthetic."
    assert with_markers(text, [CHART]) == (
        f"{text}\n\nInline attachments without a body position:\n[Image: image001.png]"
    )


def test_repeated_source_nodes_repeat_the_image_without_listing_it_again() -> None:
    parts = [ImagePart(CHART.content_id), TextPart(" then "), ImagePart(CHART.content_id)]
    assert with_markers("", [CHART], parts=parts) == ("[Image: image001.png] then [Image: image001.png]")


def test_ambiguous_or_missing_cids_do_not_select_an_arbitrary_attachment() -> None:
    duplicate = _file("a3", "another.png", inline=True, cid=CHART.content_id)
    parts = [ImagePart(CHART.content_id, "chart"), ImagePart("not-listed", "missing")]
    assert with_markers("", [CHART, duplicate], parts=parts) == (
        "[Inline image unavailable: chart] [Inline image unavailable: missing]\n\n"
        "Inline attachments without a body position:\n[Image: image001.png]\n[Image: another.png]"
    )


def test_external_images_do_not_claim_to_be_attachment_files() -> None:
    assert with_markers("", [], parts=[ImagePart(None, "a chart")]) == "[External image: a chart]"


def test_inline_selection_uses_typed_source_nodes_and_never_cid_strings_in_prose() -> None:
    parts = [TextPart(f"cid:{CHART.content_id} [Image: image001.png]")]
    assert wanted(CHART, parts=parts) is False
    assert wanted(CHART, parts=[ImagePart(CHART.content_id)]) is True
    assert wanted(CHART, parts=None) is True
    unknown = _file("unknown", "unknown.png", inline=True)
    assert wanted(unknown, parts=[]) is True
    assert with_markers("", [CHART], parts=parts) == parts[0].text

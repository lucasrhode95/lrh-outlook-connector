from __future__ import annotations

import pytest

from outlook_connector.domain.body import ImagePart, TextPart
from outlook_connector.remote.graph_mapping import message
from outlook_connector.remote.html_body import parse_html
from outlook_connector.service.markers import with_markers


def test_source_image_nodes_remain_distinct_from_escaped_markup_and_literal_markers() -> None:
    parts = parse_html(
        'Before [Image: invented.png] &lt;img src="cid:fake"&gt;'
        '<img alt="same label" src="CID:first%40example.invalid">'
        '<img alt="same label" src="cid:second@example.invalid"> after'
    )
    assert parts == [
        TextPart('Before [Image: invented.png] <img src="cid:fake">'),
        ImagePart("first@example.invalid", "same label"),
        ImagePart("second@example.invalid", "same label"),
        TextPart(" after"),
    ]


def test_source_blocks_lists_links_tables_quotes_and_preformatted_text_are_readable() -> None:
    parts = parse_html(
        "<head><style>ignored CSS</style></head><p>Hello <b>reader</b>.</p>"
        '<ul><li>One</li><li>Two</li></ul><p><a href="https://example.invalid/plan">Plan</a></p>'
        "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table>"
        "<blockquote><p>Quoted</p><p>History</p></blockquote><pre>  code\n    indented</pre>"
        '<div hidden>secret</div><div style="display: none !important">hidden<img src="cid:hidden"></div>'
        "<script>ignored()</script><p>End</p>"
    )
    assert parts is not None
    rendered = with_markers("", [], parts=parts)
    assert rendered == (
        "Hello reader.\n\n- One\n- Two\n\nPlan <https://example.invalid/plan>\n\n"
        "A\tB\nC\tD\n\n> Quoted\n\n> History\n\n  code\n    indented\n\nEnd"
    )
    assert not any(isinstance(part, ImagePart) for part in parts)


def test_declared_plain_text_is_never_interpreted_as_source_html() -> None:
    source = '<img src="cid:fake"> [Image: invented.png]'
    page = message({"id": "synthetic", "body": {"contentType": "text", "content": source}}, html=True)
    assert page.body_parts == [TextPart(source)]
    assert with_markers("", [], parts=page.body_parts) == source


def test_full_and_unique_source_parts_are_separate_and_internal() -> None:
    page = message(
        {
            "id": "synthetic",
            "body": {"contentType": "html", "content": 'Full<img src="cid:full">'},
            "uniqueBody": {"contentType": "html", "content": 'Unique<img src="cid:unique">'},
        },
        html=True,
    )
    assert page.parts("full") == [TextPart("Full"), ImagePart("full")]
    assert page.parts("unique") == [TextPart("Unique"), ImagePart("unique")]
    assert "body_parts" not in page.model_dump() and "unique_body_parts" not in page.model_dump()


def test_missing_unique_source_structure_does_not_substitute_full_history() -> None:
    page = message(
        {"id": "synthetic", "body": {"contentType": "html", "content": 'Full<img src="cid:full">'}},
        html=True,
    )
    assert page.parts("unique") is None


def test_unclosed_source_markup_and_malformed_link_urls_do_not_lose_visible_content() -> None:
    parts = parse_html('<div>Before <a href="https://[">link<img src="cid:chart"> after')
    assert parts == [TextPart("Before link"), ImagePart("chart"), TextPart(" after")]


def test_consecutive_source_line_breaks_preserve_blank_lines() -> None:
    assert with_markers("", [], parts=parse_html("<p>Before<br><br>after</p>")) == "Before\n\nafter"


def test_preformatted_body_keeps_indentation_on_its_first_line() -> None:
    assert with_markers("", [], parts=parse_html("<pre>  first\n    second</pre>")) == "  first\n    second"


@pytest.mark.parametrize(
    "source",
    [
        '<textarea><img src="cid:literal-text"></textarea>',
        '<div style="background-image: url(cid:background)">Visible text</div>',
        '<!--[if gte mso 9]><v:imagedata src="cid:conditional"><![endif]-->',
        '<picture><source srcset="cid:alternate"><img src="cid:fallback"></picture>',
        '<v:imagedata src="cid:vml">',
        '<![CDATA[<img src="cid:literal">]]>',
    ],
)
def test_unsupported_or_malformed_structure_does_not_claim_complete_image_positions(source: str) -> None:
    assert parse_html(source) is None


@pytest.mark.parametrize(
    "style",
    [
        "display:none; display:block",
        "visibility:hidden; visibility:visible",
        "display:none; display:block !important",
    ],
)
def test_later_inline_declarations_override_earlier_hiding(style: str) -> None:
    assert parse_html(f'<span style="{style}">Shown<img src="cid:chart"></span>') == [
        TextPart("Shown"),
        ImagePart("chart"),
    ]


def test_important_hiding_survives_a_later_ordinary_declaration() -> None:
    parts = parse_html('<div style="display:none !important; display:block">Hidden<img src="cid:x"></div>End')
    assert parts == [TextPart("End")]


def test_author_display_overrides_the_hidden_attribute() -> None:
    assert parse_html('<span hidden style="display:inline">Shown</span>') == [TextPart("Shown")]


def test_descendant_restoring_inherited_visibility_is_not_claimed_complete() -> None:
    source = (
        '<div style="visibility:hidden">Hidden<span style="visibility:visible"><img src="cid:x"></span></div>'
    )
    assert parse_html(source) is None


def test_inherited_invisibility_without_override_hides_the_subtree() -> None:
    parts = parse_html(
        '<div style="visibility:hidden">Hidden<span style="visibility:inherit">'
        '<img src="cid:x"></span></div>End'
    )
    assert parts == [TextPart("End")]


def test_comment_markup_cannot_create_an_image_reference() -> None:
    assert parse_html('<!-- <img src="cid:fake"> --><img src="cid:real">') == [ImagePart("real")]


def test_bogus_declarations_treated_as_comments_do_not_create_image_metadata() -> None:
    parts = parse_html('<![unexpected[<img src="cid:unknown">]]>')
    assert parts is not None and not any(isinstance(part, ImagePart) for part in parts)

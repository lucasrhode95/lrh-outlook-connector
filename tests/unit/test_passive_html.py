from __future__ import annotations

import pytest

from outlook_connector.domain.errors import InvalidRequest
from outlook_connector.service.passive_html import passive_start_tags


@pytest.mark.parametrize(
    "page",
    [
        "<p>Tom&Jerry</p>",
        "<p>Tom&Jerry",
        "<p>a < b</p>",
        "<!-- note --><p>x</p>",
        "<style>p{color:red}</style><p>x</p>",
        '<img src="data:image/png;base64,AAAA" alt="logo">',
    ],
)
def test_complete_markup_passes(page: str) -> None:
    passive_start_tags(page)


@pytest.mark.parametrize(
    "page",
    [
        "<p>hi</p><img src=x onerror=alert(1)",
        '<img alt="abc',
        "<p>x</p><!-- open",
        "<style>p{color:red}",
        "<img src=x onerror=alert(1)//",
    ],
)
def test_unfinished_markup_is_refused(page: str) -> None:
    with pytest.raises(InvalidRequest, match="ends inside"):
        passive_start_tags(page)


@pytest.mark.parametrize(
    ("page", "match"),
    [
        ("<script>x</script>", "Active HTML content is not allowed: script"),
        ('<img src=x onerror="alert(1)">', "Active HTML URLs and event handlers"),
        ('<a href="java&#x73;cript:alert(1)">x</a>', "Active HTML URLs"),
    ],
)
def test_active_markup_is_refused(page: str, match: str) -> None:
    with pytest.raises(InvalidRequest, match=match):
        passive_start_tags(page)


def test_start_tags_carry_their_source_for_later_rules() -> None:
    (tag,) = passive_start_tags('<img src="data:image/png;base64,AAAA">')
    assert tag.name == "img"
    assert 'src="data:image/png;base64,AAAA"' in tag.source

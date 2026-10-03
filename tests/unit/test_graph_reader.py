from __future__ import annotations

import asyncio
import base64
from datetime import UTC, datetime
from pathlib import Path

import httpx
import pytest

from outlook_connector.domain.errors import (
    AuthenticationRequired,
    Failure,
    InvalidRequest,
    NotFound,
    Throttled,
    Upstream,
)
from outlook_connector.remote.graph import BATCH_CONCURRENCY, Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.remote.transport import _path as transport_path
from tests.fakes.graph_fake import FakeGraph, FakeMessage, StaticTokens, sample_mailbox


async def _no_sleep(_s: float) -> None:
    return None


def reader_for(fake: FakeGraph) -> GraphMailReader:
    transport = Transport(
        StaticTokens(), client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep
    )
    return GraphMailReader(Graph(transport))


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


async def test_folders_are_walked_recursively_with_aliases(fake: FakeGraph) -> None:
    folders = {f.id: f for f in await reader_for(fake).list_folders()}
    assert set(folders) == {
        "f-inbox",
        "f-sent",
        "f-drafts",
        "f-deleted",
        "f-junk",
        "f-archive",
        "f-proj",
        "f-rie",
    }
    assert folders["f-inbox"].well_known == "inbox"
    assert folders["f-rie"].parent_id == "f-proj" and folders["f-rie"].well_known is None


async def test_list_messages_newest_first_with_date_window_and_paging(fake: FakeGraph) -> None:
    reader = reader_for(fake)
    page, link = await reader.list_messages(
        folder_id=None, since=datetime(2026, 9, 28, 9, 30, tzinfo=UTC), until=None, page_size=2, page=None
    )
    assert [m.id for m in page] == ["m5", "m4"]
    assert link
    more, link = await reader.list_messages(folder_id=None, since=None, until=None, page_size=2, page=link)
    assert [m.id for m in more] == ["m3", "m2"] and link is None


async def test_list_messages_in_one_folder(fake: FakeGraph) -> None:
    page, _ = await reader_for(fake).list_messages(
        folder_id="f-inbox", since=None, until=None, page_size=50, page=None
    )
    assert [m.id for m in page] == ["m5", "m1"]
    assert page[0].is_read is False and page[0].sender and page[0].sender.address == "alice@example.com"


async def test_conversation_spans_folders_and_never_uses_orderby(fake: FakeGraph) -> None:
    messages, truncated = await reader_for(fake).conversation("c-rel")
    assert {m.folder_id for m in messages} == {"f-inbox", "f-sent", "f-rie", "f-junk"} and not truncated


async def test_large_conversation_reports_truncation(
    fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("outlook_connector.remote.graph_mail.MAX_CONVERSATION", 2)
    messages, truncated = await reader_for(fake).conversation("c-rel")
    assert len(messages) == 2 and truncated


async def test_conversation_folders_in_one_batch(fake: FakeGraph) -> None:
    result = await reader_for(fake).conversation_folders(["c-rel", "c-lunch", "c-none"])
    assert sorted(folder for folder, _ in result["c-rel"][0]) == ["f-inbox", "f-junk", "f-rie", "f-sent"]
    assert result["c-lunch"] == ([("f-inbox", "<m5@example.com>")], False) and result["c-none"] == ([], False)
    assert fake.calls == ["POST /v1.0/$batch"]


async def test_list_attachments_many_in_one_batch(fake: FakeGraph) -> None:
    result, failed = await reader_for(fake).list_attachments_many(["m3", "m5"])
    assert [a.name for a in result["m3"]] == ["numbers.xlsx", "image001.png", "logo.png"]
    assert result["m5"] == [] and not failed and fake.calls == ["POST /v1.0/$batch"]


async def test_get_message_text_and_html_with_attachments(fake: FakeGraph) -> None:
    reader = reader_for(fake)
    text = await reader.get_message("m2")
    assert text.unique_body_text == "Thanks!" and "> First report" in (text.body_text or "")
    html = await reader.get_message("m3", body_format="html")
    assert "cid:img1" in (html.unique_body_html or "")
    assert [(a.name, a.kind, a.is_inline) for a in html.attachments][:2] == [
        ("numbers.xlsx", "file", False),
        ("image001.png", "file", True),
    ]


async def test_get_message_missing_raises_not_found(fake: FakeGraph) -> None:
    with pytest.raises(NotFound):
        await reader_for(fake).get_message("gone")


async def test_batch_get_returns_none_for_missing(fake: FakeGraph) -> None:
    ids = [f"m{i}" for i in range(1, 6)] + [f"gone{i}" for i in range(20)]  # forces two batch chunks
    result = await reader_for(fake).get_messages(ids)
    assert result.messages["m1"] and result.messages["m1"].unique_body_text == "First report"
    assert all(result.messages[f"gone{i}"] is None for i in range(20)) and not result.failed
    assert fake.calls.count("POST /v1.0/$batch") == 2


async def test_batch_request_ids_never_collide_ignoring_case(fake: FakeGraph) -> None:
    # Immutable ids can differ only by case; Graph rejects a batch whose request ids do.
    fake.add(FakeMessage("AAkx", "upper", "f-inbox", "2026-09-01T00:00:00Z", conversation="c-a"))
    fake.add(FakeMessage("AAkX", "lower", "f-inbox", "2026-09-01T00:01:00Z", conversation="c-b"))
    result = await reader_for(fake).get_messages(["AAkx", "AAkX"])
    assert {mid: m.subject for mid, m in result.messages.items() if m} == {"AAkx": "upper", "AAkX": "lower"}


async def test_throttled_batch_items_are_retried_in_batches_of_at_most_20(fake: FakeGraph) -> None:
    ids = [f"gone{i}" for i in range(60)]
    fake.throttle_items = 45  # spread over three batches; one retry batch of 45 would be rejected
    result = await reader_for(fake).get_messages(ids)
    assert not result.failed and all(result.messages[i] is None for i in ids)
    assert max(fake.batch_sizes) <= 20 and sum(fake.batch_sizes) == 60 + 45


async def test_batch_concurrency_is_shared_across_calls(
    fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    reader = reader_for(fake)
    graph = reader._graph
    original = graph._batch_once
    in_flight = peak = 0

    async def counting(requests, prefer, headers):  # type: ignore[no-untyped-def]
        nonlocal in_flight, peak
        in_flight += 1
        peak = max(peak, in_flight)
        await asyncio.sleep(0.01)
        try:
            return await original(requests, prefer, headers)
        finally:
            in_flight -= 1

    monkeypatch.setattr(graph, "_batch_once", counting)
    # many small batch() calls at once
    await asyncio.gather(*(reader.attachment_content_ids({"m3": ["a2", "a3"]}) for _ in range(8)))
    assert peak == BATCH_CONCURRENCY


async def test_persistently_throttled_items_are_reported_not_raised(fake: FakeGraph) -> None:
    fake.throttle_items = 10_000
    result = await reader_for(fake).get_messages(["m1", "m2"])
    assert set(result.failed) == {"m1", "m2"} and not result.messages
    assert result.failed["m1"] == Failure(
        status=429, code="ApplicationThrottled", message="Too many requests."
    )


async def test_errors_name_the_operation_code_message_and_request_id(fake: FakeGraph) -> None:
    with pytest.raises(InvalidRequest) as caught:
        await reader_for(fake).list_messages(
            folder_id="no-such-route/x", since=None, until=None, page_size=5, page=None
        )
    text = str(caught.value)
    assert text.startswith("While listing messages: graph.microsoft.com rejected the request (HTTP 400")
    assert "UnsupportedByFake" in text and "request-id req-" in text


async def test_search(fake: FakeGraph) -> None:
    reader = reader_for(fake)
    hits, link = await reader.search(query="relatório", folder_id=None, page_size=25, page=None)
    assert {m.id for m in hits} == {"m1", "m2", "m3"} and link is None


async def test_attachment_content_ids_and_downloads(fake: FakeGraph, tmp_path: Path) -> None:
    reader = reader_for(fake)
    found = await reader.attachment_content_ids({"m3": ["a2", "a3"], "m1": []})
    assert found == {"m3": {"a2": "img1", "a3": "sig"}, "m1": {}}
    size = await reader.download_attachment("m3", "a1", tmp_path / "numbers.xlsx")
    assert size == len(b"xlsx-bytes") and (tmp_path / "numbers.xlsx").read_bytes() == b"xlsx-bytes"
    assert await reader.download_mime("m1", tmp_path / "m1.eml") > 0


async def test_throttling_is_retried(fake: FakeGraph) -> None:
    fake.throttle_next = 2
    folders = await reader_for(fake).list_folders()
    assert folders


async def test_persistent_throttling_raises(fake: FakeGraph) -> None:
    fake.throttle_next = 100
    with pytest.raises(Throttled):
        await reader_for(fake).list_folders()


async def test_continuation_links_must_stay_on_graph(fake: FakeGraph) -> None:
    with pytest.raises(InvalidRequest):
        await reader_for(fake).list_messages(
            folder_id=None,
            since=None,
            until=None,
            page_size=5,
            page="https://evil.example.com/v1.0/me/messages",
        )


async def test_odata_quotes_are_escaped_in_conversation_ids() -> None:
    fake = FakeGraph()
    fake.add(FakeMessage("q1", "quote", "f", "2026-09-01T00:00:00Z", conversation="it's"))
    messages, _ = await reader_for(fake).conversation("it's")
    assert [m.id for m in messages] == ["q1"]


async def test_rejected_token_is_renewed_once(fake: FakeGraph) -> None:
    tokens_ = StaticTokens()
    fake.reject_tokens = 1
    reader = GraphMailReader(Graph(Transport(tokens_, client=httpx.AsyncClient(transport=fake.transport()))))
    assert await reader.list_folders()
    assert tokens_.renewals == [{"force_refresh": True}]


async def test_claims_challenge_is_passed_on(fake: FakeGraph) -> None:
    tokens_ = StaticTokens()
    fake.reject_tokens, fake.claims_challenge = 1, base64.b64encode(b'{"access_token":{}}').decode()
    reader = GraphMailReader(Graph(Transport(tokens_, client=httpx.AsyncClient(transport=fake.transport()))))
    await reader.list_folders()
    assert tokens_.renewals == [{"claims_challenge": '{"access_token":{}}'}]


async def test_token_still_rejected_after_renewal_asks_to_sign_in(fake: FakeGraph) -> None:
    fake.reject_tokens = 2
    with pytest.raises(AuthenticationRequired, match="outlook-connector auth read"):
        await reader_for(fake).list_folders()


async def test_access_denied_is_not_a_sign_in_problem(
    fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    original = fake.route

    def deny(method, path, params, prefer, request):  # type: ignore[no-untyped-def]
        if path == "/me/messages/m1":
            return 403, {"error": {"code": "ErrorAccessDenied"}}, None
        return original(method, path, params, prefer, request)

    monkeypatch.setattr(fake, "route", deny)
    with pytest.raises(Upstream, match="denied access"):
        await reader_for(fake).get_message("m1")


def test_logged_paths_never_carry_ids() -> None:
    url = "https://graph.microsoft.com/v1.0/me/messages/" + "A" * 120 + "/attachments"
    assert transport_path(url) == "/v1.0/me/messages/{id}/attachments"


async def test_renewal_applies_to_one_request_only(fake: FakeGraph) -> None:
    tokens_ = StaticTokens()
    fake.reject_tokens, fake.throttle_next = 1, 1  # 401, then (renewed) 429, then success
    transport = Transport(tokens_, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep)
    assert await GraphMailReader(Graph(transport)).list_folders()
    assert tokens_.renewals == [{"force_refresh": True}]  # the retry after 429 did not refresh again


async def test_search_returns_the_same_permanent_ids_as_listings(fake: FakeGraph) -> None:
    # Graph's $search ignores the immutable-id preference; the reader reads the ids back.
    reader = reader_for(fake)
    hits, _ = await reader.search(query="relatório", folder_id=None, page_size=25, page=None)
    listed, _ = await reader.list_messages(folder_id=None, since=None, until=None, page_size=25, page=None)
    assert {m.id for m in hits} <= {m.id for m in listed} and not any(m.id.startswith("rest.") for m in hits)

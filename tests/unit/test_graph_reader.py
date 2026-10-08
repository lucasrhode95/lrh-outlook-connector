from __future__ import annotations

import asyncio
import base64
from datetime import UTC, datetime
from pathlib import Path

import httpx
import pytest

from outlook_connector.domain.errors import (
    AuthenticationRequired,
    DownloadLimitExceeded,
    Failure,
    InvalidRequest,
    NotFound,
    Throttled,
    Upstream,
)
from outlook_connector.remote import graph_mapping as mapping
from outlook_connector.remote.graph import BATCH_CONCURRENCY, Graph
from outlook_connector.remote.graph_mail import GraphMailReader
from outlook_connector.remote.transport import Transport
from outlook_connector.remote.transport import _path as transport_path
from tests.fakes.graph_fake import FakeGraph, FakeMessage, StaticTokens, sample_mailbox


async def _no_sleep(_s: float) -> None:
    return None


def graph_for(fake: FakeGraph) -> Graph:
    transport = Transport(
        StaticTokens(), client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep
    )
    return Graph(transport)


def reader_for(fake: FakeGraph) -> GraphMailReader:
    return GraphMailReader(graph_for(fake))


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


def test_summary_uses_one_received_first_timestamp() -> None:
    data = FakeMessage("time", "Date", "f-inbox", "2026-09-30T12:00:00Z").json(text_body=True)
    data["receivedDateTime"] = "2026-09-30T12:00:00Z"
    data["sentDateTime"] = "2026-09-30T11:00:00Z"
    assert mapping.summary(data).received_at == datetime(2026, 9, 30, 12, tzinfo=UTC)

    data["receivedDateTime"] = None
    assert mapping.summary(data).received_at == datetime(2026, 9, 30, 11, tzinfo=UTC)


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
        "f-project",
    }
    assert folders["f-inbox"].well_known == "inbox"
    assert folders["f-project"].parent_id == "f-proj" and folders["f-project"].well_known is None


async def test_count_messages_returns_count_and_newest_date(
    fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    queries: list[dict[str, str]] = []
    original_route = fake.route

    def capture(method, path, params, prefer, request):  # type: ignore[no-untyped-def]
        if path.endswith("/mailFolders/f-inbox/messages"):
            queries.append(params)
        return original_route(method, path, params, prefer, request)

    monkeypatch.setattr(fake, "route", capture)
    counts = await reader_for(fake).count_messages(folder_ids=["f-inbox"], since=None, until=None)

    assert counts["f-inbox"].count == 2
    assert counts["f-inbox"].newest_received_at == datetime(2026, 9, 30, 12, tzinfo=UTC)
    assert queries == [
        {
            "$count": "true",
            "$top": "1",
            "$select": "receivedDateTime",
            "$orderby": "receivedDateTime desc",
        }
    ]


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
    assert {m.folder_id for m in messages} == {"f-inbox", "f-sent", "f-project", "f-junk"} and not truncated


async def test_summary_lookup_failures_keep_structured_details(fake: FakeGraph) -> None:
    fake.fail[r"/me/messages/m1"] = 403
    result = await reader_for(fake).get_summaries(["m1"])
    assert isinstance(result.failed["m1"], Failure)
    assert result.failed["m1"].status == 403


async def test_large_conversation_reports_truncation(
    fake: FakeGraph, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("outlook_connector.remote.graph_mail.MAX_CONVERSATION", 2)
    messages, truncated = await reader_for(fake).conversation("c-rel")
    assert len(messages) == 2 and truncated


async def test_conversation_folders_in_one_batch(fake: FakeGraph) -> None:
    result = await reader_for(fake).conversation_folders(["c-rel", "c-lunch", "c-none"])
    assert sorted(folder for folder, _ in result["c-rel"][0]) == ["f-inbox", "f-junk", "f-project", "f-sent"]
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
    await asyncio.gather(*(reader.list_attachments_many(["m3"]) for _ in range(8)))
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


async def test_attachment_listings_include_content_ids_without_per_item_reads(
    fake: FakeGraph, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    reader = reader_for(fake)
    selectors: list[str] = []
    per_item_reads: list[str] = []
    original_route = fake.route

    def record_route(method, path, params, prefer, request):
        if path.endswith("/attachments"):
            selectors.append(params.get("$select", ""))
        elif "/attachments/" in path and not path.endswith("/$value"):
            per_item_reads.append(path)
        return original_route(method, path, params, prefer, request)

    monkeypatch.setattr(fake, "route", record_route)
    single = await reader.list_attachments("m3")
    many, failed = await reader.list_attachments_many(["m3", "m1"])

    assert not failed
    assert {item.id: item.content_id for item in single} == {"a1": None, "a2": "img1", "a3": "sig"}
    assert {item.id: item.content_id for item in many["m3"]} == {"a1": None, "a2": "img1", "a3": "sig"}
    assert len(selectors) == 3 and all(
        "microsoft.graph.fileAttachment/contentId" in value for value in selectors
    )
    assert not per_item_reads

    size = await reader.download_attachment("m3", "a1", tmp_path / "numbers.xlsx")
    assert size == len(b"xlsx-bytes") and (tmp_path / "numbers.xlsx").read_bytes() == b"xlsx-bytes"
    assert await reader.download_mime("m1", tmp_path / "m1.eml") > 0


async def test_download_limit_is_a_typed_connector_error(fake: FakeGraph, tmp_path: Path) -> None:
    with pytest.raises(DownloadLimitExceeded) as caught:
        await graph_for(fake).download(
            "/me/messages/m3/attachments/a1/$value", tmp_path / "oversized.bin", max_bytes=3
        )
    assert caught.value.limit_bytes == 3
    assert str(caught.value) == ("Download exceeds the connector's local limit of 3 bytes.")


async def test_downloads_are_retried_after_throttling_and_dropped_connections(
    fake: FakeGraph, tmp_path: Path
) -> None:
    reader = reader_for(fake)
    fake.throttle_next = 2
    assert await reader.download_attachment("m3", "a1", tmp_path / "a.xlsx") == len(b"xlsx-bytes")
    fake.drop_downloads = 2
    assert await reader.download_attachment("m3", "a1", tmp_path / "b.xlsx") == len(b"xlsx-bytes")
    assert (tmp_path / "b.xlsx").read_bytes() == b"xlsx-bytes"


async def test_a_download_that_keeps_timing_out_is_a_domain_error(fake: FakeGraph, tmp_path: Path) -> None:
    fake.drop_downloads = 100
    with pytest.raises(Upstream, match="No complete response") as raised:
        await reader_for(fake).download_attachment("m3", "a1", tmp_path / "a.xlsx")
    assert raised.value.failure == Failure(status=None, message=str(raised.value))


async def test_download_renews_a_rejected_token_once(fake: FakeGraph, tmp_path: Path) -> None:
    tokens_ = StaticTokens()
    fake.reject_tokens = 1
    transport = Transport(tokens_, client=httpx.AsyncClient(transport=fake.transport()), sleep=_no_sleep)
    reader = GraphMailReader(Graph(transport))
    size = await reader.download_attachment("m3", "a1", tmp_path / "renewed.xlsx")
    assert size == len(b"xlsx-bytes")
    assert tokens_.renewals == [{"force_refresh": True}]


async def test_throttling_is_retried(fake: FakeGraph) -> None:
    fake.throttle_next = 2
    folders = await reader_for(fake).list_folders()
    assert folders


async def test_persistent_throttling_raises(fake: FakeGraph) -> None:
    fake.throttle_next = 100
    with pytest.raises(Throttled):
        await reader_for(fake).list_folders()


@pytest.mark.parametrize(
    "page",
    [
        "https://evil.example.com/v1.0/me/messages",
        "https://outlook.office.com/v1.0/me/messages",
    ],
)
async def test_continuation_links_must_stay_on_graph(fake: FakeGraph, page: str) -> None:
    with pytest.raises(InvalidRequest):
        await reader_for(fake).list_messages(folder_id=None, since=None, until=None, page_size=5, page=page)


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
    with pytest.raises(AuthenticationRequired, match="outlook-connector auth graph"):
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
    # Graph's $search ignores the immutable-id preference; one translateExchangeIds call fixes a page.
    reader = reader_for(fake)
    hits, _ = await reader.search(query="relatório", folder_id=None, page_size=25, page=None)
    listed, _ = await reader.list_messages(folder_id=None, since=None, until=None, page_size=25, page=None)
    assert hits and {m.id for m in hits} <= {m.id for m in listed}
    assert not any(m.id.startswith("rest.") for m in hits)
    assert fake.calls.count("POST /v1.0/me/translateExchangeIds") == 1


async def test_search_keeps_its_ids_when_translation_fails(fake: FakeGraph) -> None:
    fake.fail[r"/me/translateExchangeIds"] = 500
    hits, _ = await reader_for(fake).search(query="relatório", folder_id=None, page_size=25, page=None)
    assert hits and all(m.id.startswith("rest.") for m in hits)  # the search itself still works


@pytest.mark.parametrize("status", [429, 502, 503, 504])
async def test_batch_retries_only_transient_items_preserving_successes(
    fake: FakeGraph, monkeypatch: pytest.MonkeyPatch, status: int
) -> None:
    reader = reader_for(fake)
    original = fake.route
    seen = {"/me/messages/m1": 0, "/me/messages/m5": 0}

    def transient(method, path, params, prefer, request):  # type: ignore[no-untyped-def]
        if path in seen:
            seen[path] += 1
            if path == "/me/messages/m1" and seen[path] == 1:
                return status, {"error": {"code": "SyntheticTransient"}}, None
        return original(method, path, params, prefer, request)

    monkeypatch.setattr(fake, "route", transient)
    result = await reader.get_messages(["m1", "m5"])
    assert not result.failed and result.messages["m1"] and result.messages["m5"]
    assert seen == {"/me/messages/m1": 2, "/me/messages/m5": 1}
    assert fake.batch_sizes == [2, 1]


@pytest.mark.parametrize("status", [400, 403, 404, 500])
async def test_batch_does_not_retry_nontransient_items(fake: FakeGraph, status: int) -> None:
    fake.fail[r"/me/messages/m1"] = status
    result = await reader_for(fake).get_messages(["m1"])
    assert fake.batch_sizes == [1]
    if status == 404:
        assert result.messages["m1"] is None
    else:
        assert result.failed["m1"].status == status


@pytest.mark.parametrize("status", [429, 502, 503, 504])
async def test_batch_exhaustion_retains_final_transient_status(fake: FakeGraph, status: int) -> None:
    fake.fail[r"/me/messages/m1"] = status
    result = await reader_for(fake).get_messages(["m1"])
    assert result.failed["m1"].status == status and fake.batch_sizes == [1] * 5

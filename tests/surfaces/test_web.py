from __future__ import annotations

import zipfile
from io import BytesIO

import httpx
import pytest
from starlette.testclient import TestClient

from outlook_connector.bootstrap import AppContext
from outlook_connector.surfaces.web.routes import Activity, create_app
from tests.fakes.graph_fake import FakeGraph, sample_mailbox
from tests.surfaces.test_mcp import FakeTokens

TOKEN = "session-token-for-tests"
PORT = 8765


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


@pytest.fixture
def client(fake: FakeGraph) -> TestClient:
    context = AppContext(tokens=FakeTokens(), http_client=httpx.AsyncClient(transport=fake.transport()))  # type: ignore[arg-type]
    app = create_app(context, session_token=TOKEN, port=PORT, activity=Activity())
    return TestClient(app, base_url=f"http://127.0.0.1:{PORT}", headers={"X-Session-Token": TOKEN})


def test_index_embeds_the_session_token(client: TestClient) -> None:
    page = client.get("/")
    assert page.status_code == 200 and f'content="{TOKEN}"' in page.text
    assert client.get("/static/app.js").status_code == 200


def test_api_requires_the_session_token(client: TestClient) -> None:
    assert client.get("/api/folders", headers={"X-Session-Token": "wrong"}).status_code == 403


def test_foreign_host_header_is_refused(client: TestClient) -> None:
    # A DNS-rebinding page would arrive with its own host name.
    assert client.get("/", headers={"Host": f"evil.example.com:{PORT}"}).status_code == 403


def test_folders_messages_search_thread_and_message(client: TestClient) -> None:
    assert any(f["path"] == "Inbox/Projects/RIE" for f in client.get("/api/folders").json())
    page = client.get("/api/messages", params={"folder": "inbox", "limit": 10}).json()
    assert [m["id"] for m in page["items"]] == ["m5", "m1"]
    found = client.get("/api/search", params={"q": "relatório"}).json()
    assert found["conversations"][0]["conversation_id"] == "c-rel"
    thread = client.get("/api/threads/c-rel").json()
    assert [t["message"]["id"] for t in thread["messages"]] == ["m1", "m2", "m3"]
    assert "text" not in thread["messages"][0]  # the UI lists threads without bodies
    message = client.get("/api/messages/m2", params={"body": "full"}).json()
    assert "> First report" in message["text"]


def test_thread_sizes(client: TestClient) -> None:
    sizes = client.post("/api/thread-sizes", json={"conversation_ids": ["c-rel", "c-lunch"]}).json()
    assert [(s["conversation_id"], s["messages"]) for s in sizes] == [("c-rel", 3), ("c-lunch", 1)]
    assert client.post("/api/thread-sizes", json={"conversation_ids": "c-rel"}).status_code == 400


def test_attachment_download_and_junk_scope(client: TestClient) -> None:
    response = client.get("/api/messages/m3/attachments/a1")
    assert response.status_code == 200 and response.content == b"xlsx-bytes"
    ids = [m["id"] for m in client.get("/api/messages").json()["items"]]
    with_junk = client.get("/api/messages", params={"include_deleted_items": "true"}).json()["items"]
    assert "m4" not in ids and "m4" in [m["id"] for m in with_junk]


def test_dates_from_the_browser_are_accepted(client: TestClient) -> None:
    page = client.get("/api/messages", params={"since": "2026-09-30T03:00:00.000Z"}).json()
    assert [m["id"] for m in page["items"]] == ["m5"]


def test_errors_map_to_http_status(client: TestClient) -> None:
    response = client.get("/api/messages", params={"folder": "Nope"})
    assert response.status_code == 400 and response.json()["kind"] == "InvalidRequest"
    assert client.get("/api/messages/does-not-exist").status_code == 404
    assert client.get("/api/messages", params={"since": "yesterday"}).status_code == 400


def test_export_downloads_one_file(client: TestClient) -> None:
    response = client.post("/api/export", json={"conversation_ids": ["c-rel"], "include_attachments": True})
    assert response.status_code == 200 and response.headers["content-type"] == "application/zip"
    assert "attachment" in response.headers["content-disposition"]
    with zipfile.ZipFile(BytesIO(response.content)) as archive:
        assert any(name.endswith("numbers.xlsx") for name in archive.namelist())


def test_export_rejects_an_empty_selection(client: TestClient) -> None:
    assert client.post("/api/export", json={}).status_code == 400


def test_activity_is_tracked(fake: FakeGraph) -> None:
    activity = Activity()
    activity.last -= 1000
    context = AppContext(tokens=FakeTokens(), http_client=httpx.AsyncClient(transport=fake.transport()))  # type: ignore[arg-type]
    app = create_app(context, session_token=TOKEN, port=PORT, activity=activity)
    TestClient(app, base_url=f"http://127.0.0.1:{PORT}").post(
        "/api/heartbeat", headers={"X-Session-Token": TOKEN}
    )
    assert activity.idle_seconds() < 5

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import httpx
import pytest
from mcp.server.fastmcp import FastMCP

from outlook_connector.auth.tokens import AccessToken, CacheStatus, ProfileStatus
from outlook_connector.bootstrap import AppContext
from outlook_connector.surfaces.mcp_main import build_server
from tests.fakes.graph_fake import FakeGraph, sample_mailbox
from tests.fakes.msal_fakes import jwt

READ_TOOLS = {
    "auth_status", "list_folders", "list_messages", "search_messages", "get_thread",
    "get_message", "list_attachments", "download_attachment", "save_message_mime", "export_messages",
}  # fmt: skip


class FakeTokens:
    """Token provider stand-in carrying a synthetic account identity."""

    def get_token(self, profile: str, **_: object) -> AccessToken:
        claims = {
            "tid": "tenant-x",
            "oid": "user-x",
            "upn": "me@example.com",
            "aud": "https://graph.microsoft.com",
        }
        return AccessToken(profile=profile, value=jwt(claims), source="cache", expires_on=None)

    def sign_in_command(self, profile: str) -> str:
        return f"outlook-connector auth {profile}"

    def status(self) -> CacheStatus:
        return CacheStatus(
            "encrypted", Path("x"), True, (), (ProfileStatus("read", "c", (), "reads", True, None),)
        )


@pytest.fixture
def fake() -> FakeGraph:
    return sample_mailbox()


@pytest.fixture
def server(fake: FakeGraph) -> FastMCP:
    context = AppContext(tokens=FakeTokens(), http_client=httpx.AsyncClient(transport=fake.transport()))  # type: ignore[arg-type]
    return build_server(context)


async def call(server: FastMCP, name: str, **arguments: Any) -> Any:
    result = await server.call_tool(name, arguments)
    if isinstance(result, tuple):  # (content blocks, structured output)
        structured = result[1]
        return structured.get("result", structured) if isinstance(structured, dict) else structured
    if isinstance(result, dict):
        return result
    return json.loads(result[0].text)  # type: ignore[union-attr]


async def test_tools_and_annotations(server: FastMCP) -> None:
    tools = {t.name: t for t in await server.list_tools()}
    assert set(tools) == READ_TOOLS  # the MVP surface is read-only
    for tool in tools.values():
        assert tool.annotations and tool.annotations.readOnlyHint and not tool.annotations.destructiveHint
    assert "never attempt to sign in" in (server.instructions or "")


async def test_list_search_thread_message_flow(server: FastMCP) -> None:
    page = await call(server, "list_messages", folder="inbox", limit=10)
    assert [m["id"] for m in page["items"]] == ["m5", "m1"] and page["coverage"]["complete"]
    found = await call(server, "search_messages", query="relatório")
    conversation_id = found["conversations"][0]["conversation_id"]
    thread = await call(server, "get_thread", conversation_id=conversation_id)
    assert [m["message"]["id"] for m in thread["messages"]] == ["m1", "m2", "m3"]
    message = await call(server, "get_message", message_id="m2")
    assert message["text"] == "Thanks!"


async def test_list_messages_received_only(server: FastMCP) -> None:
    page = await call(server, "list_messages", received_only=True)
    assert [m["id"] for m in page["items"]] == ["m5", "m3", "m1"]
    assert "received_only=true" in (server.instructions or "")


async def test_export_by_range_and_throttling_guidance(server: FastMCP) -> None:
    artifact = await call(server, "export_messages", since="2026-09-30T00:00:00", limit=10)
    assert artifact["message_count"] == 1 and artifact["messages_unavailable"] == 0
    assert "4 concurrent requests and 10,000 requests per 10 minutes" in (server.instructions or "")


async def test_list_results_are_compact_by_default(server: FastMCP) -> None:
    compact = await call(server, "list_messages", folder="inbox")
    assert "to" not in compact["items"][0] and "internet_message_id" not in compact["items"][0]
    full = await call(server, "list_messages", folder="inbox", detail="full")
    assert full["items"][0]["to"] and full["items"][0]["internet_message_id"]


async def test_naive_datetimes_are_treated_as_utc(server: FastMCP) -> None:
    page = await call(server, "list_messages", since="2026-09-30T00:00:00")
    assert [m["id"] for m in page["items"]] == ["m5"]


async def test_attachment_download_and_export_return_local_paths(server: FastMCP) -> None:
    attachments = await call(server, "list_attachments", message_id="m3")
    saved = await call(server, "download_attachment", message_id="m3", attachment_id=attachments[0]["id"])
    assert Path(saved["path"]).read_bytes() == b"xlsx-bytes" and saved["name"] == "numbers.xlsx"
    eml = await call(server, "save_message_mime", message_id="m1")
    assert eml["path"].endswith(".eml")
    artifact = await call(server, "export_messages", conversation_ids=["c-rel"])
    assert Path(artifact["path"]).exists() and artifact["message_count"] == 3


async def test_errors_are_reported_as_tool_errors(server: FastMCP) -> None:
    with pytest.raises(Exception, match="Unknown folder"):
        await call(server, "list_messages", folder="No such folder")


async def test_auth_status_is_offline(server: FastMCP, fake: FakeGraph) -> None:
    status = await call(server, "auth_status")
    assert status["profiles"]["read"]["signed_in"] and fake.calls == []


def test_server_name_is_not_mistakable_for_an_official_connector(server: FastMCP) -> None:
    assert server.name == "lrh-outlook"

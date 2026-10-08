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
    "auth_status", "list_folders", "list_messages", "search_messages", "get_conversation",
    "get_message", "list_attachments", "download_attachment", "save_message_mime", "export_messages",
    "list_rules", "list_signatures", "get_signature",
}  # fmt: skip
WRITE_TOOLS = {
    "create_draft",
    "send_draft",
    "create_signature",
    "update_signature",
    "delete_signature",
    "set_default_signature",
}
CHANGE_TOOLS = {"set_read_state", "set_flag"}
RELOCATE_TOOLS = {
    "move_messages",
    "delete_messages",
    "create_rule",
    "update_rule",
    "reorder_rules",
    "delete_rule",
}


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


async def call(server: FastMCP, tool_name: str, **arguments: Any) -> Any:
    result = await server.call_tool(tool_name, arguments)
    if isinstance(result, tuple):  # (content blocks, structured output)
        structured = result[1]
        return structured.get("result", structured) if isinstance(structured, dict) else structured
    if isinstance(result, dict):
        return result
    return json.loads(result[0].text)  # type: ignore[union-attr]


async def test_tools_and_annotations(server: FastMCP) -> None:
    tools = {t.name: t for t in await server.list_tools()}
    assert set(tools) == READ_TOOLS | WRITE_TOOLS | CHANGE_TOOLS | RELOCATE_TOOLS
    for name in CHANGE_TOOLS | RELOCATE_TOOLS:
        annotations = tools[name].annotations
        assert annotations and not annotations.readOnlyHint and not annotations.openWorldHint
        assert annotations.destructiveHint is (name in RELOCATE_TOOLS)
    for name in READ_TOOLS:
        annotations = tools[name].annotations
        assert annotations and annotations.readOnlyHint and not annotations.destructiveHint
    draft, send = tools["create_draft"].annotations, tools["send_draft"].annotations
    assert draft and not draft.readOnlyHint and not draft.destructiveHint and not draft.openWorldHint
    assert send and not send.readOnlyHint and send.destructiveHint and send.openWorldHint
    assert "never attempt to sign in" in (server.instructions or "")
    assert "Never confirm on the user's behalf" in (server.instructions or "")
    assert "Drafts are composed once" in (server.instructions or "")
    assert "Never delete first" in (server.instructions or "")


async def test_signature_tools_are_public_and_default_for_argument_is_required(
    server: FastMCP, fake: FakeGraph
) -> None:
    tools = {tool.name: tool for tool in await server.list_tools()}
    assert {"list_signatures", "get_signature"} <= set(tools)
    assert {"create_signature", "update_signature", "delete_signature", "set_default_signature"} <= set(tools)
    assert tools["list_signatures"].annotations and tools["list_signatures"].annotations.readOnlyHint
    assert tools["create_signature"].annotations and not tools["create_signature"].annotations.readOnlyHint
    assert tools["set_default_signature"].inputSchema["required"] == ["name", "for"]
    assert set(tools["set_default_signature"].inputSchema["properties"]) == {"name", "for"}

    created = await call(server, "create_signature", **{"name": "Synthetic", "html": "<p>Signature</p>"})
    assert created["status"] == "created"
    selected = await call(server, "set_default_signature", **{"name": "Synthetic", "for": "both"})
    assert selected["status"] == "default_set" and fake.signature_new_default == "Synthetic"
    listed = await call(server, "list_signatures")
    assert listed["signatures"][0]["name"] == "Synthetic" and listed["signatures"][0]["readable"]


async def test_draft_replacement_is_created_and_verified_before_deletion(
    server: FastMCP, fake: FakeGraph
) -> None:
    original = await call(
        server, "create_draft", to=["bob@example.com"], subject="First", text_body="First version"
    )
    replacement = await call(
        server, "create_draft", to=["bob@example.com"], subject="Second", text_body="Second version"
    )
    assert original["verified"] and replacement["verified"]
    assert original["id"] != replacement["id"]
    assert replacement["message"]["subject"] == "Second"
    assert replacement["text_body"] == "Second version"

    deleted = await call(server, "delete_messages", message_ids=[original["id"]])
    assert deleted["counts"] == {"done": 1}
    assert fake.messages[original["id"]].folder == fake.aliases["deleteditems"]
    assert fake.messages[replacement["id"]].is_draft
    # Keep the fake mailbox tidy after checking the public replacement workflow.
    await call(server, "delete_messages", message_ids=[replacement["id"]])
    tools = {t.name: t for t in await server.list_tools()}
    assert set(tools["send_draft"].inputSchema["properties"]) == {"draft_id"}


async def test_list_search_conversation_message_flow(server: FastMCP) -> None:
    page = await call(server, "list_messages", folder="inbox", limit=10)
    assert [m["id"] for m in page["items"]] == ["m5", "m1"] and page["coverage"]["complete"]
    found = await call(server, "search_messages", query="relatório")
    conversation_id = found["conversations"][0]["conversation_id"]
    conversation = await call(server, "get_conversation", conversation_id=conversation_id)
    assert [m["message"]["id"] for m in conversation["messages"]] == ["m1", "m2", "m3"]
    message = await call(server, "get_message", message_id="m2")
    assert message["text"] == "Thanks!"


async def test_scope_is_one_tool_argument(server: FastMCP) -> None:
    tools = {tool.name: tool for tool in await server.list_tools()}
    for name in ("list_messages", "search_messages", "get_conversation", "export_messages", "set_read_state"):
        properties = tools[name].inputSchema["properties"]
        assert "scope" in properties


async def test_list_messages_without_sent_items(server: FastMCP) -> None:
    page = await call(server, "list_messages", scope={"sent_items": False})
    assert [m["id"] for m in page["items"]] == ["m5", "m3", "m1"]
    assert "scope.sent_items=false" in (server.instructions or "")
    assert "m2" in [m["id"] for m in (await call(server, "list_messages"))["items"]]  # included by default


async def test_export_by_folder_date_window_and_throttling_guidance(server: FastMCP) -> None:
    artifact = await call(server, "export_messages", since="2026-09-30T00:00:00", limit=10)
    assert artifact["message_count"] == 1
    assert "4 concurrent requests and 10,000 requests per 10 minutes" in (server.instructions or "")
    assert "Scope alone does not select anything." in (server.instructions or "")


async def test_export_scope_alone_returns_selection_guidance(server: FastMCP) -> None:
    with pytest.raises(Exception, match="Select conversations, messages, a folder, or a date window"):
        await call(server, "export_messages", scope={"sent_items": False})


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
    assert Path(artifact["path"]).exists() and artifact["message_count"] == 4


async def test_errors_are_reported_as_tool_errors(server: FastMCP) -> None:
    with pytest.raises(Exception, match="Unknown folder"):
        await call(server, "list_messages", folder="No such folder")


async def test_auth_status_is_offline(server: FastMCP, fake: FakeGraph) -> None:
    status = await call(server, "auth_status")
    assert status["profiles"]["read"]["signed_in"] and fake.calls == []


def test_server_name_is_not_mistakable_for_an_official_connector(server: FastMCP) -> None:
    assert server.name == "lrh-outlook"


async def test_mutation_tools_report_per_message(server: FastMCP, fake: FakeGraph) -> None:
    read = await call(server, "set_read_state", read=True, message_ids=["m5", "gone"])
    assert {r["id"]: r["status"] for r in read["results"]} == {"m5": "done", "gone": "not_found"}
    moved = await call(server, "move_messages", message_ids=["m5"], folder="archive")
    assert moved["counts"] == {"done": 1} and fake.messages["m5"].folder == "f-archive"
    deleted = await call(server, "delete_messages", message_ids=["m5"])
    assert deleted["counts"] == {"done": 1} and fake.messages["m5"].folder == "f-deleted"
    flagged = await call(server, "set_flag", message_ids=["m1"], flagged=True)
    assert flagged["counts"] == {"done": 1}


async def test_mutation_continue_on_error_schema_defaults(server: FastMCP) -> None:
    tools = {t.name: t for t in await server.list_tools()}
    for name in CHANGE_TOOLS | {"move_messages", "delete_messages"}:
        assert tools[name].inputSchema["properties"]["continue_on_error"]["default"] is True


async def test_rule_proposal_and_confirmed_mcp_write(server: FastMCP, fake: FakeGraph) -> None:
    changes = {
        "name": "Rule",
        "subject_contains": ["synthetic"],
        "move_to_folder": "archive",
        "stop_processing": True,
    }
    proposal = await call(server, "create_rule", changes=changes)
    assert proposal["status"] == "proposed" and not fake.inbox_rules
    assert proposal["changes"] == {**changes, "move_to_folder": proposal["changes"]["move_to_folder"]}
    result = await call(server, "create_rule", changes=changes, user_confirmation=proposal["confirmation"])
    assert result["status"] == "done" and result["rule_id"]
    rules = await call(server, "list_rules")
    assert rules[0]["name"] == "Rule"


async def test_search_naive_dates_use_service_normalization(server: FastMCP) -> None:
    result = await call(
        server,
        "search_messages",
        query="relatório",
        since="2026-09-28T09:30:00",
        until="2026-09-29T07:00:00Z",
    )
    assert [m["id"] for hit in result["conversations"] for m in hit["matching_messages"]] == ["m2"]

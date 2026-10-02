"""An in-memory Microsoft Graph mailbox served through httpx.MockTransport.

It understands exactly the request shapes the connector sends (research §3): folder listing,
message listing with receivedDateTime / conversationId filters, $search, $top paging,
$batch, /search/query, attachments and $value downloads. All data is synthetic.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from urllib.parse import parse_qs, urlencode, urlsplit

import httpx

ROOT = "https://graph.microsoft.com/v1.0"


@dataclass
class FakeAttachment:
    id: str
    name: str
    data: bytes = b"bytes"
    content_type: str = "application/pdf"
    inline: bool = False
    kind: str = "fileAttachment"
    content_id: str | None = None
    broken: bool = False  # the $value download fails


@dataclass
class FakeMessage:
    id: str
    subject: str
    folder: str
    received: str  # ISO with Z
    conversation: str = "conv-1"
    sender: str = "alice@example.com"
    to: tuple[str, ...] = ("me@example.com",)
    text: str = "Hello"
    unique_text: str | None = None
    html: str = "<p>Hello</p>"
    unique_html: str | None = None
    is_read: bool = True
    attachments: list[FakeAttachment] = field(default_factory=list)

    def json(self, *, text_body: bool) -> dict[str, Any]:
        body = self.text if text_body else self.html
        unique = (
            (self.unique_text if self.unique_text is not None else self.text)
            if text_body
            else (self.unique_html if self.unique_html is not None else self.html)
        )
        kind = "text" if text_body else "html"
        return {
            "id": self.id,
            "conversationId": self.conversation,
            "parentFolderId": self.folder,
            "subject": self.subject,
            "from": {"emailAddress": {"name": self.sender.split("@")[0].title(), "address": self.sender}},
            "toRecipients": [{"emailAddress": {"name": t, "address": t}} for t in self.to],
            "ccRecipients": [],
            "bccRecipients": [],
            "receivedDateTime": self.received,
            "sentDateTime": self.received,
            "isRead": self.is_read,
            "isDraft": False,
            "hasAttachments": bool(self.attachments),
            "importance": "normal",
            "categories": [],
            "flag": {"flagStatus": "notFlagged"},
            "bodyPreview": self.text[:50],
            "internetMessageId": f"<{self.id}@example.com>",
            "body": {"contentType": kind, "content": body},
            "uniqueBody": {"contentType": kind, "content": unique},
        }


@dataclass
class FakeGraph:
    folders: list[dict[str, Any]] = field(default_factory=list)
    aliases: dict[str, str] = field(default_factory=dict)
    messages: dict[str, FakeMessage] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)
    throttle_next: int = 0  # respond 429 to this many upcoming top-level requests

    # ------------------------------------------------------------------ helpers for tests
    def add_folder(
        self, fid: str, name: str, parent: str = "root", alias: str | None = None, hidden: bool = False
    ):
        self.folders.append(
            {
                "id": fid,
                "displayName": name,
                "parentFolderId": parent,
                "isHidden": hidden,
                "childFolderCount": 0,
                "totalItemCount": 0,
                "unreadItemCount": 0,
            }
        )
        for f in self.folders:
            if f["id"] == parent:
                f["childFolderCount"] += 1
        if alias:
            self.aliases[alias] = fid

    def add(self, message: FakeMessage) -> FakeMessage:
        self.messages[message.id] = message
        return message

    def transport(self) -> httpx.MockTransport:
        return httpx.MockTransport(self.handle)

    # ------------------------------------------------------------------ request handling
    def handle(self, request: httpx.Request) -> httpx.Response:
        self.calls.append(f"{request.method} {request.url.path}")
        if self.throttle_next:
            self.throttle_next -= 1
            return httpx.Response(
                429, headers={"retry-after": "0"}, json={"error": {"code": "TooManyRequests"}}
            )
        assert request.headers.get("authorization", "").startswith("Bearer "), "missing bearer token"
        prefer = request.headers.get("prefer", "")
        assert 'IdType="ImmutableId"' in prefer, "every Graph request must ask for immutable ids"
        path = request.url.path.removeprefix("/v1.0")
        params = {k: v[0] for k, v in parse_qs(urlsplit(str(request.url)).query).items()}
        status, body, content = self.route(request.method, path, params, prefer, request)
        if content is not None:
            return httpx.Response(
                status, content=content, headers={"content-type": "application/octet-stream"}
            )
        return httpx.Response(status, json=body)

    def route(
        self, method: str, path: str, params: dict[str, str], prefer: str, request: httpx.Request | None
    ):
        if method == "POST" and path == "/$batch":
            return 200, self.batch(json.loads(request.content)), None
        if method == "POST" and path == "/search/query":
            q = json.loads(request.content)["requests"][0]["query"]["queryString"]
            total = len(self.search_matches(q, None))
            return 200, {"value": [{"hitsContainers": [{"total": total, "hits": []}]}]}, None
        text_body = 'outlook.body-content-type="text"' in prefer

        if m := re.fullmatch(r"/me/mailFolders", path):
            return (
                200,
                self.paged([f for f in self.folders if f["parentFolderId"] == "root"], path, params),
                None,
            )
        if m := re.fullmatch(r"/me/mailFolders/([^/]+)/childFolders", path):
            kids = [f for f in self.folders if f["parentFolderId"] == m[1]]
            return 200, self.paged(kids, path, params), None
        if m := re.fullmatch(r"/me/mailFolders/([^/]+)", path):
            fid = self.aliases.get(m[1], m[1])
            match = [f for f in self.folders if f["id"] == fid]
            return (200, match[0], None) if match else (404, {"error": {"code": "ErrorItemNotFound"}}, None)
        if m := re.fullmatch(
            r"(?:/me/mailFolders/([^/]+))?/me/messages|/me/mailFolders/([^/]+)/messages", path
        ):
            folder = m[1] or m[2]
            folder = self.aliases.get(folder, folder) if folder else None
            return 200, self.list_messages(folder, params, path, text_body), None
        if m := re.fullmatch(r"/me/messages/([^/]+)", path):
            msg = self.messages.get(m[1])
            if not msg:
                return 404, {"error": {"code": "ErrorItemNotFound"}}, None
            return 200, msg.json(text_body=text_body), None
        if m := re.fullmatch(r"/me/messages/([^/]+)/attachments", path):
            msg = self.messages.get(m[1])
            if not msg:
                return 404, {"error": {"code": "ErrorItemNotFound"}}, None
            return 200, {"value": [self.attachment_json(a) for a in msg.attachments]}, None
        if m := re.fullmatch(r"/me/messages/([^/]+)/attachments/([^/]+)(/\$value)?", path):
            msg = self.messages.get(m[1])
            att = next((a for a in (msg.attachments if msg else []) if a.id == m[2]), None)
            if not att:
                return 404, {"error": {"code": "ErrorItemNotFound"}}, None
            if m[3]:
                if att.broken:
                    return 503, {"error": {"code": "ServiceUnavailable"}}, None
                return 200, None, att.data
            return 200, {"contentId": att.content_id, "id": att.id}, None
        if m := re.fullmatch(r"/me/messages/([^/]+)/\$value", path):
            msg = self.messages.get(m[1])
            if not msg:
                return 404, {"error": {"code": "ErrorItemNotFound"}}, None
            return 200, None, f"Subject: {msg.subject}\r\n\r\n{msg.text}".encode()
        return 400, {"error": {"code": "UnsupportedByFake", "message": path}}, None

    def attachment_json(self, a: FakeAttachment) -> dict[str, Any]:
        return {
            "@odata.type": f"#microsoft.graph.{a.kind}",
            "id": a.id,
            "name": a.name,
            "contentType": a.content_type,
            "size": len(a.data),
            "isInline": a.inline,
        }

    def list_messages(
        self, folder: str | None, params: dict[str, str], path: str, text_body: bool
    ) -> dict[str, Any]:
        rows = [m for m in self.messages.values() if folder is None or m.folder == folder]
        flt = params.get("$filter", "")
        if "$orderby" in params and "conversationId" in flt:
            raise AssertionError("Graph rejects $orderby with a conversationId filter (InefficientFilter)")
        for op, value in re.findall(r"receivedDateTime (ge|le) (\S+)", flt):
            bound = datetime.fromisoformat(value.replace("Z", "+00:00"))
            rows = [m for m in rows if (_dt(m.received) >= bound if op == "ge" else _dt(m.received) <= bound)]
        if cid := re.search(r"conversationId eq '((?:[^']|'')*)'", flt):
            rows = [m for m in rows if m.conversation == cid[1].replace("''", "'")]
        if "$search" in params:
            rows = self.search_matches(params["$search"].strip('"'), rows)
        else:
            rows.sort(key=lambda m: m.received, reverse=True)
        return self.paged([m.json(text_body=text_body) for m in rows], path, params)

    def search_matches(self, query: str, rows: list[FakeMessage] | None) -> list[FakeMessage]:
        rows = list(self.messages.values()) if rows is None else rows
        terms = [t.split(":", 1)[-1].lower() for t in query.split()]
        return [m for m in rows if all(t in f"{m.subject} {m.text}".lower() for t in terms)]

    def paged(self, items: list[dict[str, Any]], path: str, params: dict[str, str]) -> dict[str, Any]:
        top = int(params.get("$top", 10))
        skip = int(params.get("$skip", 0))
        page = items[skip : skip + top]
        body: dict[str, Any] = {"value": page}
        if skip + top < len(items):
            body["@odata.nextLink"] = f"{ROOT}{path}?{urlencode({**params, '$skip': skip + top})}"
        return body

    def batch(self, payload: dict[str, Any]) -> dict[str, Any]:
        assert len(payload["requests"]) <= 20, "Graph allows at most 20 requests per batch"
        responses = []
        for sub in payload["requests"]:
            url = urlsplit(sub["url"])
            params = {k: v[0] for k, v in parse_qs(url.query).items()}
            prefer = sub.get("headers", {}).get("Prefer", "")
            assert 'IdType="ImmutableId"' in prefer
            status, body, _ = self.route(sub["method"], url.path, params, prefer, None)
            responses.append({"id": sub["id"], "status": status, "body": body})
        return {"responses": responses}


def _dt(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


class StaticTokens:
    """TokenSource stand-in."""

    class _T:
        value = "test-token"

    def get_token(self, profile: str) -> Any:
        return self._T()


def sample_mailbox() -> FakeGraph:
    g = FakeGraph()
    g.add_folder("f-inbox", "Inbox", alias="inbox")
    g.add_folder("f-sent", "Sent Items", alias="sentitems")
    g.add_folder("f-deleted", "Deleted Items", alias="deleteditems")
    g.add_folder("f-junk", "Junk Email", alias="junkemail")
    g.add_folder("f-archive", "Archive", alias="archive")
    g.add_folder("f-proj", "Projects", parent="f-inbox")
    g.add_folder("f-rie", "RIE", parent="f-proj")
    g.add(
        FakeMessage(
            "m1",
            "Relatório BE semanal",
            "f-inbox",
            "2026-09-28T09:00:00Z",
            conversation="c-rel",
            text="First report\n\nregards",
            unique_text="First report",
        )
    )
    g.add(
        FakeMessage(
            "m2",
            "RE: Relatório BE semanal",
            "f-sent",
            "2026-09-28T10:00:00Z",
            conversation="c-rel",
            sender="me@example.com",
            to=("alice@example.com",),
            text="Thanks!\n\n> First report",
            unique_text="Thanks!",
        )
    )
    g.add(
        FakeMessage(
            "m3",
            "RE: Relatório BE semanal",
            "f-rie",
            "2026-09-29T08:00:00Z",
            conversation="c-rel",
            text="Follow-up with numbers",
            attachments=[
                FakeAttachment("a1", "numbers.xlsx", b"xlsx-bytes"),
                FakeAttachment("a2", "image001.png", b"png", "image/png", inline=True, content_id="img1"),
                FakeAttachment("a3", "logo.png", b"png2", "image/png", inline=True, content_id="sig"),
            ],
            html='<p>Follow-up</p><img src="cid:img1">',
            unique_html='<p>Follow-up</p><img src="cid:img1">',
        )
    )
    g.add(
        FakeMessage(
            "m4", "Spam offer", "f-junk", "2026-09-29T09:00:00Z", conversation="c-rel", text="buy now"
        )
    )
    g.add(
        FakeMessage(
            "m5",
            "Lunch?",
            "f-inbox",
            "2026-09-30T12:00:00Z",
            conversation="c-lunch",
            text="Lunch at noon?",
            is_read=False,
        )
    )
    return g

"""An in-memory Microsoft Graph mailbox served through httpx.MockTransport.

It understands exactly the request shapes the connector sends (research §3): folder listing,
message listing with receivedDateTime / conversationId filters, $search, $top paging,
$batch, attachments and $value downloads. It also answers the OWS write actions (research §4)
on the same mailbox, so a write can be read back through Graph. All data is synthetic.
"""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from datetime import UTC, datetime
from html import unescape
from typing import Any
from urllib.parse import parse_qs, unquote, urlencode, urlsplit

import httpx

ROOT = "https://graph.microsoft.com/v1.0"
REST_PREFIX = "rest."  # marks the regular (move-sensitive) id form that $search returns


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
    cc: tuple[str, ...] = ()
    bcc: tuple[str, ...] = ()
    is_draft: bool = False
    flagged: bool = False
    categories: list[str] = field(default_factory=list)
    text: str = "Hello"
    unique_text: str | None = None
    html: str = "<p>Hello</p>"
    unique_html: str | None = None
    is_read: bool = True
    attachments: list[FakeAttachment] = field(default_factory=list)
    internet_id: str | None = None  # copies of one message (e.g. sent to yourself) share it
    meeting: dict[str, Any] = field(default_factory=dict)  # eventMessage fields of meeting mail

    def change_key(self) -> str:
        """Like Exchange's change key: a new value whenever the item changes."""
        state = (self.subject, self.folder, self.to, self.cc, self.bcc, self.text, self.html, self.is_draft)
        return hashlib.sha256(repr(state).encode()).hexdigest()[:16]

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
            "ccRecipients": [{"emailAddress": {"name": t, "address": t}} for t in self.cc],
            "bccRecipients": [{"emailAddress": {"name": t, "address": t}} for t in self.bcc],
            "receivedDateTime": self.received,
            "sentDateTime": self.received,
            "isRead": self.is_read,
            "isDraft": self.is_draft,
            "hasAttachments": any(not a.inline for a in self.attachments),  # false when inline-only
            "importance": "normal",
            "categories": list(self.categories),
            "flag": {"flagStatus": "flagged" if self.flagged else "notFlagged"},
            "bodyPreview": self.text[:50],
            "internetMessageId": self.internet_id or f"<{self.id}@example.com>",
            "changeKey": self.change_key(),
            "body": {"contentType": kind, "content": body},
            "uniqueBody": {"contentType": kind, "content": unique},
            **self.meeting,
        }


@dataclass
class FakeGraph:
    folders: list[dict[str, Any]] = field(default_factory=list)
    aliases: dict[str, str] = field(default_factory=dict)
    messages: dict[str, FakeMessage] = field(default_factory=dict)
    calls: list[str] = field(default_factory=list)
    throttle_next: int = 0  # respond 429 to this many upcoming top-level requests
    throttle_items: int = 0  # respond 429 to this many upcoming $batch sub-requests
    reject_tokens: int = 0  # respond 401 (token rejected) to this many upcoming top-level requests
    drop_downloads: int = 0  # the connection drops (read timeout) on this many upcoming $value downloads
    claims_challenge: str | None = None  # base64 claims sent with those 401s (CAE)
    fail: dict[str, int] = field(default_factory=dict)  # path regex -> HTTP status to answer instead
    batch_sizes: list[int] = field(default_factory=list)
    sent_drafts: list[str] = field(default_factory=list)  # drafts sent with UpdateItem
    ows_calls: list[tuple[str, dict[str, Any]]] = field(default_factory=list)  # (action, request body)
    inbox_rules: list[dict[str, Any]] = field(default_factory=list)
    ows_next: list[Any] = field(default_factory=list)  # scripted answers for upcoming OWS calls:
    # "no-answer" (connection drops after sending), "done-no-answer" (applied, then dropped),
    # "no-items" / "not-json" (HTTP 200 without readable item results),
    # an int (that HTTP status), or a dict (that item result; for a bare-request call, the answer)
    me: str = "me@example.com"
    reply_drops_history: bool = False  # simulate a reply draft that lost the quoted original
    reply_flattens_html: bool = False  # simulate a reply whose quoted original lost its formatting
    display_name: str = "Doe, Jane"
    photo: bytes | None = None  # the user's 48x48 profile photo; None: no photo set

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
        if self.reject_tokens:
            self.reject_tokens -= 1
            challenge = f'Bearer error="insufficient_claims", claims="{self.claims_challenge}"'
            return httpx.Response(
                401,
                headers={"www-authenticate": challenge} if self.claims_challenge else {},
                json={"error": {"code": "InvalidAuthenticationToken"}},
            )
        if self.throttle_next:
            self.throttle_next -= 1
            return httpx.Response(
                429, headers={"retry-after": "0"}, json={"error": {"code": "TooManyRequests"}}
            )
        if self.drop_downloads and request.url.path.endswith("/$value"):
            self.drop_downloads -= 1
            raise httpx.ReadTimeout("Injected read timeout.", request=request)
        if request.url.host == "outlook.cloud.microsoft":
            return self.handle_ows(request)
        assert request.headers.get("authorization", "").startswith("Bearer "), "missing bearer token"
        prefer = request.headers.get("prefer", "")
        assert 'IdType="ImmutableId"' in prefer, "every Graph request must ask for immutable ids"
        path = request.url.path.removeprefix("/v1.0")
        params = {k: v[0] for k, v in parse_qs(urlsplit(str(request.url)).query).items()}
        status, body, content = self.route(request.method, path, params, prefer, request)
        request_id = {"request-id": f"req-{len(self.calls)}"}
        if content is not None:
            return httpx.Response(
                status, content=content, headers={"content-type": "application/octet-stream", **request_id}
            )
        return httpx.Response(status, json=body, headers=request_id)

    # ------------------------------------------------------------------ OWS (writes)
    def handle_ows(self, request: httpx.Request) -> httpx.Response:
        assert request.url.path == "/owa/service.svc", request.url.path
        assert request.headers.get("authorization", "").startswith("Bearer ")
        assert 'IdType="ImmutableId"' in request.headers.get("prefer", "")
        action = request.url.params["action"]
        assert request.headers.get("action") == action
        posted = request.headers.get("x-owa-urlpostdata")
        envelope = json.loads(unquote(posted)) if posted else json.loads(request.content)
        if envelope["__type"] == f"{action}Request:#Exchange":  # bare request style (inbox rules)
            return self._ows_bare(action, envelope, request)
        assert envelope["__type"] == f"{action}JsonRequest:#Exchange"
        body = envelope["Body"]
        assert body["__type"] == f"{action}Request:#Exchange"
        self.ows_calls.append((action, body))
        script = self.ows_next.pop(0) if self.ows_next else None
        if script == "no-answer":
            raise httpx.ReadTimeout("no answer", request=request)
        if script == "no-items":
            return httpx.Response(200, json={"Body": {}})
        if script == "not-json":
            return httpx.Response(200, content=b"<html>", headers={"content-type": "text/html"})
        if isinstance(script, int):
            return httpx.Response(script, headers={"x-owa-error": "FakeError"}, json={})
        if isinstance(script, dict):
            return httpx.Response(200, json={"Body": {"ResponseMessages": {"Items": [script]}}})
        items = getattr(self, f"ows_{action}")(body)
        if script == "done-no-answer":
            raise httpx.ReadTimeout("no answer", request=request)
        return httpx.Response(200, json={"Body": {"ResponseMessages": {"Items": items}}})

    def _ows_bare(
        self, action: str, request_object: dict[str, Any], request: httpx.Request
    ) -> httpx.Response:
        """The inbox-rule style: no wrapper, no Body; the answer carries ``WasSuccessful``."""
        assert request_object["Header"]["__type"] == "JsonRequestHeaders:#Exchange"
        self.ows_calls.append((action, request_object))
        script = self.ows_next.pop(0) if self.ows_next else None
        if script == "no-answer":
            raise httpx.ReadTimeout("no answer", request=request)
        if isinstance(script, int):
            return httpx.Response(script, headers={"x-owa-error": "FakeError"}, json={})
        answer = script if isinstance(script, dict) else self._rule_action(action, request_object)
        if script == "done-no-answer":
            raise httpx.ReadTimeout("no answer", request=request)
        return httpx.Response(200, json=answer)

    def _rule_action(self, action: str, request: dict[str, Any]) -> dict[str, Any]:
        import copy

        answer: dict[str, Any] = {"WasSuccessful": True, "ErrorCode": 0}
        if action == "GetInboxRule":
            answer["InboxRuleCollection"] = {"InboxRules": copy.deepcopy(self.inbox_rules)}
        elif action == "NewInboxRule":
            rule = copy.deepcopy(request["InboxRule"])
            rule["Identity"] = {
                "RawIdentity": f"rule-{len(self.inbox_rules) + 1}",
                "DisplayName": rule["Name"],
            }
            rule["Enabled"] = True
            self.inbox_rules.insert(0, rule)
            answer["InboxRule"] = copy.deepcopy(rule)
        elif action == "SetInboxAndSweepRules":
            by_id = {r["Identity"]["RawIdentity"]: r for r in self.inbox_rules}
            self.inbox_rules = [
                by_id[r["Identity"]["RawIdentity"]] for r in request["EnableDisableInboxRules"]
            ]
            for rule, entry in zip(self.inbox_rules, request["EnableDisableInboxRules"], strict=True):
                rule["Enabled"] = entry["IsEnabled"]
        elif action in ("SetInboxRule", "EnableInboxRule", "DisableInboxRule", "RemoveInboxRule"):
            fields = request.get("InboxRule") or request
            rule = next(
                (
                    r
                    for r in self.inbox_rules
                    if r["Identity"]["RawIdentity"] == fields["Identity"]["RawIdentity"]
                ),
                None,
            )
            if rule is None:
                return {"WasSuccessful": False, "ErrorCode": 5, "ErrorMessage": "No such rule"}
            if action == "SetInboxRule":
                rule.update(copy.deepcopy(fields))
            elif action == "RemoveInboxRule":
                self.inbox_rules.remove(rule)
            else:
                rule["Enabled"] = action == "EnableInboxRule"
        for priority, rule in enumerate(self.inbox_rules, 1):
            rule["Priority"] = priority
        return answer

    def ows_CreateItem(self, body: dict[str, Any]) -> list[dict[str, Any]]:  # noqa: N802
        disposition = body["MessageDisposition"]
        (item,) = body["Items"]
        addresses = lambda key: tuple(r["EmailAddress"] for r in item.get(key, []))  # noqa: E731
        content = item.get("Body") or item.get("NewBodyContent")
        text = content["Value"]
        page = content["Value"] if content["BodyType"] == "HTML" else f"<p>{content['Value']}</p>"
        if content["BodyType"] == "HTML":  # what the HTML shows, as Graph's text view would
            text = unescape(re.sub(r"<[^>]+>", "", text.replace("<br>", "\n")))
        conversation = f"conv-new-{len(self.messages)}"
        inline: list[FakeAttachment] = []
        if item["__type"] in ("ReplyToItem:#Exchange", "ReplyAllToItem:#Exchange"):
            assert content["BodyType"] == "HTML", (
                "a reply body must be HTML, or Exchange flattens the history"
            )
            original = self.messages.get(_graph_id(item["ReferenceItemId"]["Id"]))
            if original is None:
                return [{"ResponseClass": "Error", "ResponseCode": "ErrorItemNotFound"}]
            conversation = original.conversation
            if not self.reply_drops_history:
                text += "\n\nFrom: " + original.sender + "\n" + original.text
                inline = [a for a in original.attachments if a.inline]
                quoted = f"<p>{original.text}</p>" if self.reply_flattens_html else original.html
                page += f"<div><b>From:</b> {original.sender}</div>{quoted}"  # Exchange re-wraps it
        else:
            assert item["__type"] == "Message:#Exchange" and item["MessageDisposition"] == disposition
        new_id = f"w{len(self.messages)}-x_y"  # has "-" and "_", so the id mapping is exercised
        folder = self.aliases["drafts" if disposition == "SaveOnly" else "sentitems"]
        now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
        self.add(
            FakeMessage(new_id, item.get("Subject") or "", folder, now, conversation=conversation,
                        sender=self.me, to=addresses("ToRecipients"), cc=addresses("CcRecipients"),
                        bcc=addresses("BccRecipients"), is_draft=disposition == "SaveOnly", text=text,
                        attachments=list(inline), html=page)
        )  # fmt: skip
        created = [{"ItemId": {"Id": _ows_id(new_id)}}] if disposition == "SaveOnly" else []
        return [{"ResponseClass": "Success", "ResponseCode": "NoError", "Items": created}]

    def _ows_target(self, item_id: dict[str, Any]) -> FakeMessage | None:
        return self.messages.get(_graph_id(item_id["Id"]))

    def ows_UpdateItem(self, body: dict[str, Any]) -> list[dict[str, Any]]:  # noqa: N802
        assert body["SuppressReadReceipts"] is True
        if body["MessageDisposition"] == "SendAndSaveCopy":  # Outlook Web sends a draft this way
            (change,) = body["ItemChanges"]
            assert body["ConflictResolution"] == "NeverOverwrite" and change["ItemId"].get("ChangeKey")
            draft = self._ows_target(change["ItemId"])
            if draft is None or not draft.is_draft:
                return [{"ResponseClass": "Error", "ResponseCode": "ErrorItemNotFound"}]
            if change["ItemId"]["ChangeKey"] != draft.change_key():  # changed since it was read
                return [{"ResponseClass": "Error", "ResponseCode": "ErrorIrresolvableConflict"}]
            draft.folder, draft.is_draft = self.aliases["sentitems"], False
            self.sent_drafts.append(draft.id)
            return [{"ResponseClass": "Success", "ResponseCode": "NoError"}]
        assert body["MessageDisposition"] == "SaveOnly"
        out = []
        for change in body["ItemChanges"]:
            message = self._ows_target(change["ItemId"])
            if message is None:
                out.append({"ResponseClass": "Error", "ResponseCode": "ErrorItemNotFound"})
                continue
            for update in change["Updates"]:
                field_uri, props = update["Path"]["FieldURI"], update["Item"]
                if field_uri == "message:IsRead":
                    message.is_read = props["IsRead"]
                elif field_uri == "item:Flag":
                    message.flagged = props["Flag"]["FlagStatus"] == "Flagged"
                elif field_uri == "item:Subject":
                    message.subject = props["Subject"]
                elif field_uri == "item:Body":
                    message.html = props["Body"]["Value"]
                    message.text = unescape(re.sub(r"<[^>]+>", "", message.html.replace("<br>", "\n")))
                elif field_uri in ("message:ToRecipients", "message:CcRecipients", "message:BccRecipients"):
                    key = field_uri.split(":")[1]
                    values = tuple(r["EmailAddress"] for r in props[key])
                    if key == "ToRecipients":
                        message.to = values
                    elif key == "CcRecipients":
                        message.cc = values
                    else:
                        message.bcc = values
                else:
                    raise AssertionError(field_uri)
            out.append({"ResponseClass": "Success", "ResponseCode": "NoError"})
        return out

    def _ows_move(self, item_ids: list[dict[str, Any]], folder: str) -> list[dict[str, Any]]:
        out = []
        for item_id in item_ids:
            message = self._ows_target(item_id)
            if message is None:
                out.append({"ResponseClass": "Error", "ResponseCode": "ErrorItemNotFound"})
                continue
            message.folder = folder
            out.append(
                {"ResponseClass": "Success", "ResponseCode": "NoError", "Items": [{"ItemId": item_id}]}
            )
        return out

    def ows_MoveItem(self, body: dict[str, Any]) -> list[dict[str, Any]]:  # noqa: N802
        target = body["ToFolderId"]["BaseFolderId"]
        if target["__type"] == "DistinguishedFolderId:#Exchange":
            folder = self.aliases[target["Id"]]
        else:
            folder = _graph_id(target["Id"])
            assert any(f["id"] == folder for f in self.folders), folder
        return self._ows_move(body["ItemIds"], folder)

    def ows_DeleteItem(self, body: dict[str, Any]) -> list[dict[str, Any]]:  # noqa: N802
        assert body["DeleteType"] == "MoveToDeletedItems", "never a hard delete"
        return self._ows_move(body["ItemIds"], self.aliases["deleteditems"])

    def route(
        self, method: str, path: str, params: dict[str, str], prefer: str, request: httpx.Request | None
    ):
        if method == "POST" and path == "/$batch":
            return self.batch(json.loads(request.content))
        for pattern, status in self.fail.items():
            if re.fullmatch(pattern, path):
                code = "ErrorAccessDenied" if status == 403 else f"Failure{status}"
                return status, {"error": {"code": code, "message": "Injected failure."}}, None
        if method == "POST" and path == "/me/translateExchangeIds":
            return self.translate_ids(json.loads(request.content))
        text_body = 'outlook.body-content-type="text"' in prefer

        if path == "/me":
            return (
                200,
                {"displayName": self.display_name, "mail": self.me, "userPrincipalName": self.me},
                None,
            )
        if path == "/me/photos/48x48/$value":
            if self.photo is None:
                return 404, {"error": {"code": "ImageNotFound", "message": "No photo."}}, None
            return 200, None, self.photo
        if m := re.fullmatch(r"/me/mailFolders", path):
            return (
                200,
                self.paged(
                    [self.folder_json(f) for f in self.folders if f["parentFolderId"] == "root"], path, params
                ),
                None,
            )
        if m := re.fullmatch(r"/me/mailFolders/([^/]+)/childFolders", path):
            kids = [self.folder_json(f) for f in self.folders if f["parentFolderId"] == m[1]]
            return 200, self.paged(kids, path, params), None
        if m := re.fullmatch(r"/me/mailFolders/([^/]+)", path):
            fid = self.aliases.get(m[1], m[1])
            match = [f for f in self.folders if f["id"] == fid]
            return (
                (200, self.folder_json(match[0]), None)
                if match
                else (404, {"error": {"code": "ErrorItemNotFound"}}, None)
            )
        if m := re.fullmatch(
            r"(?:/me/mailFolders/([^/]+))?/me/messages|/me/mailFolders/([^/]+)/messages", path
        ):
            folder = m[1] or m[2]
            folder = self.aliases.get(folder, folder) if folder else None
            if folder and not any(f["id"] == folder for f in self.folders):
                return 404, {"error": {"code": "ErrorItemNotFound", "message": "Folder not found."}}, None
            return 200, self.list_messages(folder, params, path, text_body), None
        if m := re.fullmatch(r"/me/messages/([^/]+)", path):
            msg = self.messages.get(m[1].removeprefix(REST_PREFIX))  # either id form is readable
            if not msg:
                return 404, {"error": {"code": "ErrorItemNotFound"}}, None
            # like Graph (live 2026-10-03): the id comes back in the form it was asked with
            return 200, msg.json(text_body=text_body) | {"id": m[1]}, None
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

    def folder_json(self, f: dict[str, Any]) -> dict[str, Any]:
        """A folder with its current message count (Graph's totalItemCount)."""
        return {**f, "totalItemCount": sum(1 for m in self.messages.values() if m.folder == f["id"])}

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
            # Like Graph (live 2026-10-03): $search ignores Prefer: IdType="ImmutableId" and returns
            # the regular id, which changes when the message moves.
            items = [m.json(text_body=text_body) | {"id": REST_PREFIX + m.id} for m in rows]
            return self.paged(items, path, params)
        rows.sort(key=lambda m: m.received, reverse=True)
        return self.paged([m.json(text_body=text_body) for m in rows], path, params)

    def search_matches(self, query: str, rows: list[FakeMessage] | None) -> list[FakeMessage]:
        """Words must all appear; KQL received>=/<= date terms filter by (UTC) received date."""
        rows = list(self.messages.values()) if rows is None else rows
        words, dates = [], []
        for term in query.split():
            if bound := re.fullmatch(r"received(>=|<=)(\d{4}-\d{2}-\d{2})", term):
                dates.append((bound[1], bound[2]))
            else:
                words.append(term.split(":", 1)[-1].lower())

        def in_dates(m: FakeMessage) -> bool:
            day = m.received[:10]
            return all(day >= d if op == ">=" else day <= d for op, d in dates)

        return [m for m in rows if in_dates(m) and all(w in f"{m.subject} {m.text}".lower() for w in words)]

    def paged(self, items: list[dict[str, Any]], path: str, params: dict[str, str]) -> dict[str, Any]:
        top = int(params.get("$top", 10))
        skip = int(params.get("$skip", 0))
        page = items[skip : skip + top]
        body: dict[str, Any] = {"value": page}
        if params.get("$count") == "true":
            body["@odata.count"] = len(items)
        if skip + top < len(items):
            body["@odata.nextLink"] = f"{ROOT}{path}?{urlencode({**params, '$skip': skip + top})}"
        return body

    def translate_ids(self, payload: dict[str, Any]) -> tuple[int, dict[str, Any], None]:
        """Like Graph (live 2026-10-04): regular ids -> immutable ids; an input that is already
        immutable fails the whole call."""
        ids = payload["inputIds"]
        assert payload["sourceIdType"] == "restId" and payload["targetIdType"] == "restImmutableEntryId"
        if any(not i.startswith(REST_PREFIX) for i in ids):
            error = {"code": "InvalidArgument", "message": "Invalid value for arg: storeObjectId.IdType"}
            return 400, {"error": error}, None
        return 200, {"value": [{"sourceId": i, "targetId": i.removeprefix(REST_PREFIX)} for i in ids]}, None

    def batch(self, payload: dict[str, Any]) -> tuple[int, dict[str, Any], None]:
        """Like Graph: a batch over 20 requests, or with ids equal ignoring case, is rejected whole."""
        requests = payload["requests"]
        self.batch_sizes.append(len(requests))
        if len(requests) > 20:
            return 400, {"error": {"code": "BadRequest", "message": "Too many requests in batch."}}, None
        if len({str(sub["id"]).lower() for sub in requests}) < len(requests):
            return 400, {"error": {"code": "BadRequest", "message": "Duplicate request id."}}, None
        responses = []
        for sub in requests:
            if self.throttle_items:
                self.throttle_items -= 1
                responses.append(
                    {
                        "id": sub["id"],
                        "status": 429,
                        "headers": {"Retry-After": "1"},
                        "body": {"error": {"code": "ApplicationThrottled", "message": "Too many requests."}},
                    }
                )
                continue
            url = urlsplit(sub["url"])
            params = {k: v[0] for k, v in parse_qs(url.query).items()}
            prefer = sub.get("headers", {}).get("Prefer", "")
            assert 'IdType="ImmutableId"' in prefer
            status, body, _ = self.route(sub["method"], url.path, params, prefer, None)
            responses.append({"id": sub["id"], "status": status, "body": body})
        return 200, {"responses": responses}, None


def _dt(value: str) -> datetime:
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


class StaticTokens:
    """TokenProvider stand-in (the methods Transport and Ows use)."""

    class _T:
        value = "test-token"

        def claims(self) -> dict[str, Any]:
            return {"tid": "tenant-x", "oid": "user-x", "upn": "me@example.com"}

    def __init__(self) -> None:
        self.renewals: list[dict[str, Any]] = []

    def get_token(self, profile: str, **renewal: Any) -> Any:
        if renewal:
            self.renewals.append(renewal)
        return self._T()

    def sign_in_command(self, profile: str) -> str:
        return f"outlook-connector auth {profile}"


def sample_mailbox() -> FakeGraph:
    g = FakeGraph()
    g.add_folder("f-inbox", "Inbox", alias="inbox")
    g.add_folder("f-sent", "Sent Items", alias="sentitems")
    g.add_folder("f-drafts", "Drafts", alias="drafts")
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


def _ows_id(graph_id: str) -> str:
    return graph_id.replace("-", "/").replace("_", "+")


def _graph_id(ows_id: str) -> str:
    return ows_id.replace("/", "-").replace("+", "_")

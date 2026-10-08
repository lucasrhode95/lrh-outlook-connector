"""Draft-first mail: explicit bodies, mandatory Graph read-back, existing-draft sending."""

from __future__ import annotations

import asyncio
import hashlib
import re
import tempfile
from collections import Counter
from html import escape, unescape
from html.parser import HTMLParser
from pathlib import Path

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import (
    AccountMismatch,
    ConnectorError,
    InvalidRequest,
    Upstream,
    WriteOutcomeUnknown,
)
from outlook_connector.domain.models import (
    ADDRESS_PATTERN,
    MAX_BODY_CHARS,
    MAX_RECIPIENTS,
    Attachment,
    DraftMessage,
    DraftResult,
    Message,
    OutgoingMessage,
    Recipient,
    SendResult,
)
from outlook_connector.remote.ports import MailWriter
from outlook_connector.service.conversations import base_subject
from outlook_connector.service.mailbox import Mailbox

ADDRESS = re.compile(ADDRESS_PATTERN)


class Writes:
    def __init__(self, mailbox: Mailbox, writer: MailWriter, account: Account) -> None:
        self.mailbox = mailbox
        self.writer = writer
        self.account = account

    async def create_draft(self, message: OutgoingMessage) -> DraftResult:
        """Entry point: validate body, envelope and reply arguments; save once and read back."""
        resolved = await self._resolve(message)
        self.check_account()
        draft_id = await self.writer.create_draft(resolved)
        if not draft_id:
            raise WriteOutcomeUnknown(
                "Outlook did not report a draft id; check Drafts before creating again."
            )
        result = await self._read_back(draft_id)
        if resolved.reply_to_message_id and result.verified:
            try:
                problem = await self._reply_problem(draft_id, resolved.reply_to_message_id)
            except ConnectorError as exc:
                problem = f"Reply history verification failed: {exc}"
            result.history_intact, result.history_problem = problem is None, problem
        return result

    async def send_draft(self, draft_id: str) -> SendResult:
        """Entry point: require a server draft and the bound account. Only an id, no content.

        The host obtains human approval; the connector never reconstructs or retries the message.
        """
        draft = await self._draft(draft_id)
        if not draft.revision:
            raise Upstream("Outlook did not report the draft's version, so it was not sent.")
        self.check_account()
        try:
            await self.writer.send_draft(draft_id, draft.revision)
        except WriteOutcomeUnknown as exc:
            try:
                item = await self.mailbox.message(draft_id)
                sent = await self.mailbox.resolve_folder("sentitems")
                if item.is_draft is False and item.folder_id == sent.id:
                    return SendResult(
                        status="sent", sent_item_id=item.id, detail="The exact draft is in Sent Items."
                    )
            except ConnectorError:
                pass
            return SendResult(
                status="unknown", detail=f"{exc} Do not send again before checking Sent Items and Outbox."
            )
        return SendResult(status="sent", detail="Sent the existing draft; a copy is kept in Sent Items.")

    async def _draft(self, draft_id: str) -> Message:
        """Validate a public draft id against current server state."""
        if not draft_id.strip():
            raise InvalidRequest("draft_id must not be empty.")
        message = await self.mailbox.message(draft_id)
        if not message.is_draft:
            raise InvalidRequest("The message is not an existing Outlook draft.")
        return message

    async def _resolve(self, message: OutgoingMessage) -> DraftMessage:
        """Assumes (not re-checked here): called only by create_draft; validates its outgoing input."""
        page = _body(message.text_body, message.html_body)
        to, cc, bcc = (
            _addresses(message.to, "to"),
            _addresses(message.cc, "cc"),
            _addresses(message.bcc, "bcc"),
        )
        subject = message.subject
        if message.reply_to_message_id:
            original = await self.mailbox.message(message.reply_to_message_id)
            if not (to or cc or bcc):
                to, cc = _reply_recipients(
                    original, me=(self.account.username or "").lower(), reply_all=message.reply_all
                )
            if subject is None:
                subject = f"RE: {base_subject(original.subject) or ''}".rstrip()
        elif message.reply_all:
            raise InvalidRequest("reply_all needs reply_to_message_id.")
        _validate_envelope(to, cc, bcc, subject)
        return DraftMessage(
            to=to,
            cc=cc,
            bcc=bcc,
            subject=subject or "",
            html_body=page,
            reply_to_message_id=message.reply_to_message_id,
            reply_all=message.reply_all,
        )

    async def _read_back(self, draft_id: str, *, expected: list[Attachment] | None = None) -> DraftResult:
        """Assumes (not re-checked here): reliable id from a single successful write. Retries reads only."""
        for attempt in range(3):
            try:
                text = await self.mailbox.reader.get_message(draft_id)
                page = await self.mailbox.reader.get_message(draft_id, body_format="html")
                findings = []
                if not (text.body_text or "").strip():
                    findings.append("The server draft body is empty or missing.")
                cids = re.findall(r"cid:([^\"' >]+)", page.body_html or "", re.IGNORECASE)
                if cids or expected:
                    attachments = await self.mailbox.reader.list_attachments(draft_id)
                    if expected and Counter(a.name for a in expected) - Counter(a.name for a in attachments):
                        findings.append("The server draft is missing expected attachments.")
                    found = await self.mailbox.reader.attachment_content_ids(
                        {draft_id: [a.id for a in attachments if a.is_inline]}
                    )
                    present = set(found.get(draft_id, {}).values())
                    if any(cid not in present for cid in cids):
                        findings.append("The server draft references missing inline images.")
                return DraftResult(
                    id=draft_id,
                    status="saved",
                    verified=True,
                    message=text,
                    text_body=text.body_text,
                    html_body=page.body_html,
                    findings=findings,
                )
            except ConnectorError as exc:
                if attempt == 2:
                    return DraftResult(
                        id=draft_id,
                        status="failed",
                        verified=False,
                        findings=[f"Draft saved, but Graph read-back failed: {exc}"],
                    )
                await asyncio.sleep(0.2 * (attempt + 1))
        raise AssertionError("unreachable")

    async def _reply_problem(self, draft_id: str, original_id: str) -> str | None:
        """Why the reply draft does not carry the original as received, or None when it does.

        Assumes (not re-checked here): ``draft_id`` is a reply draft this call just saved and
        ``original_id`` the message it replies to.
        """
        reader = self.mailbox.reader
        texts = await reader.get_messages([original_id, draft_id])
        pages = await reader.get_messages([original_id, draft_id], body_format="html")
        original, draft = texts.messages.get(original_id), texts.messages.get(draft_id)
        original_page, page = pages.messages.get(original_id), pages.messages.get(draft_id)
        if original is None or draft is None or original_page is None or page is None:
            return "the draft or the original could not be read back"
        if _flat(original.body_text) not in _flat(draft.body_text):
            return "the quoted original is not the original's full text"
        # Exchange re-wraps the original's HTML when it quotes it, so the HTML is not compared as
        # text; its structure is: every list, table, emphasis, link and image reference must survive.
        lost = _structure(original_page.body_html) - _structure(page.body_html)
        if lost:
            return "the quoted original lost formatting (" + ", ".join(sorted(lost.elements())) + ")"
        found, failed = await reader.list_attachments_many([original_id, draft_id])
        if failed:
            return "its attachments could not be listed"
        hashes = {mid: await self._inline_hashes(mid, found.get(mid, [])) for mid in (original_id, draft_id)}
        missing = Counter(hashes[original_id]) - Counter(hashes[draft_id])
        if missing:
            return f"{missing.total()} inline image(s) of the original are missing or changed"
        if "[cid:" in _flat(re.sub(r"<[^>]+>", " ", page.body_html or "")):
            return 'an inline image became "[cid:...]" text'
        return None

    async def _inline_hashes(self, message_id: str, attachments: list[Attachment]) -> list[str]:
        """SHA-256 of each inline image's bytes (downloaded to a temporary folder, then removed).

        Assumes (not re-checked here): ``attachments`` is the listing of ``message_id``.
        """
        out = []
        with tempfile.TemporaryDirectory(prefix="outlook-reply-check-") as folder:
            for index, attachment in enumerate(a for a in attachments if a.is_inline and a.kind == "file"):
                target = Path(folder) / str(index)
                await self.mailbox.reader.download_attachment(message_id, attachment.id, target)
                out.append(hashlib.sha256(target.read_bytes()).hexdigest())
        return out

    def check_account(self) -> None:
        claims = self.writer.account()
        if (claims.get("tid"), claims.get("oid")) != (self.account.tenant_id, self.account.object_id):
            raise AccountMismatch(
                "The write sign-in belongs to a different Microsoft account than the one this "
                "connector reads; nothing was written."
            )


def _reply_recipients(original: Message, *, me: str, reply_all: bool) -> tuple[list[str], list[str]]:
    """Outlook's reply recipients: the sender (the original recipients when you sent it); with
    reply-all also the original To (in To) and Cc (in Cc). You are left out, except when nobody
    else is left: a reply to mail you sent only to yourself goes back to you.

    Assumes (not re-checked here): ``original`` was read from the server and ``me`` is the signed-in
    address, lower case.
    """

    def plain(recipients: list[Recipient]) -> list[str]:
        return [r.address for r in recipients if r.address and r.address.lower() != me]

    sender = original.sender.address if original.sender and original.sender.address else None
    to = plain(original.to) if sender is None or sender.lower() == me else [sender]
    cc: list[str] = []
    if reply_all:
        to += plain(original.to)
        cc = plain(original.cc)
    if not to and not cc and sender:
        to = [sender]
    to = _unique(_addresses(to, "to"))
    return to, [a for a in _unique(_addresses(cc, "cc")) if a.lower() not in {t.lower() for t in to}]


FORMATTING = (
    "ul",
    "ol",
    "li",
    "table",
    "tr",
    "td",
    "th",
    "b",
    "strong",
    "i",
    "em",
    "u",
    "a",
    "img",
    "blockquote",
    "pre",
)


def _structure(page: str | None) -> Counter[str]:
    """How many of each formatting element, and of each cid: image reference, a page holds."""
    tags = Counter(t.lower() for t in re.findall(r"<([a-zA-Z0-9]+)\b", page or "") if t.lower() in FORMATTING)
    tags["image reference"] = len(re.findall(r"cid:", page or "", re.IGNORECASE))
    return tags


def _flat(text: str | None) -> str:
    """Text with its whitespace collapsed, for comparing a quoted copy with its original."""
    return re.sub(r"\s+", " ", text or "").strip()


def _addresses(values: list[str], field: str) -> list[str]:
    """The addresses of one field, validated and de-duplicated (ignoring case). Callers need not de-duplicate
    again.
    """
    out = []
    for value in values:
        address = value.strip()
        if not ADDRESS.match(address):
            raise InvalidRequest(f"Not an email address in {field}: {address!r}.")
        out.append(address)
    return _unique(out)


def _unique(addresses: list[str]) -> list[str]:
    """Drop repeats, ignoring case; the first spelling wins."""
    first: dict[str, str] = {}
    for address in addresses:
        first.setdefault(address.lower(), address)
    return list(first.values())


def _validate_envelope(to: list[str], cc: list[str], bcc: list[str], subject: str | None) -> None:
    recipients = to + cc + bcc
    if not recipients:
        raise InvalidRequest("Name at least one recipient (to, cc or bcc).")
    if len(recipients) > MAX_RECIPIENTS:
        raise InvalidRequest(f"At most {MAX_RECIPIENTS} recipients per message.")
    if len({a.lower() for a in recipients}) != len(recipients):
        raise InvalidRequest("A recipient is listed twice (across to, cc and bcc).")
    if not (subject or "").strip():
        raise InvalidRequest("A new message needs a subject.")
    if len(subject or "") > 255:
        raise InvalidRequest("Subject must be at most 255 characters.")


def _body(text: str | None, page: str | None) -> str:
    if (text is None) == (page is None):
        raise InvalidRequest("Supply exactly one of text_body or html_body.")
    value = text if text is not None else page or ""
    if not value.strip():
        raise InvalidRequest("The body is empty.")
    if len(value) > MAX_BODY_CHARS:
        raise InvalidRequest(f"Body must be at most {MAX_BODY_CHARS} characters.")
    if text is not None:
        lines = escape(text).replace("\u00a0", "&nbsp;").replace("\r\n", "\n").replace("\r", "\n").split("\n")
        return "<div>" + "<br>".join(_keep_spaces(line) for line in lines) + "</div>"
    parser = _PassiveHtml()
    parser.feed(value)
    return value


def _keep_spaces(line: str) -> str:
    line = line.replace("\t", "&nbsp;" * 4)
    line = re.sub(r" {2,}", lambda run: "&nbsp;" * (len(run.group()) - 1) + " ", line)
    return "&nbsp;" + line[1:] if line.startswith(" ") else line


class _PassiveHtml(HTMLParser):
    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "iframe", "object", "embed", "form", "input", "button", "textarea", "select"}:
            raise InvalidRequest(f"Active HTML content is not allowed: {tag}.")
        for name, value in attrs:
            compact = re.sub(r"[\s\x00-\x1f]+", "", unescape(value or "")).lower()
            if name.startswith("on") or compact.startswith(("javascript:", "vbscript:")):
                raise InvalidRequest("Active HTML URLs and event handlers are not allowed.")

    handle_startendtag = handle_starttag

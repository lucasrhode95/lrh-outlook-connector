"""Drafts and sending (requirements v4 §11.1, architecture §5.8).

Draft first: ``create_draft`` saves into Drafts and never sends, so it needs no confirmation.

Sending is two steps, and the second re-derives everything from scratch (processes are
short-lived, so nothing is remembered in between):

1. ``propose`` resolves the message exactly as it would be sent (sender, recipients, subject,
   body, reply defaults) and returns a confirmation code: a hash of all of it and of the account.
   The agent shows the proposal to the user.
2. ``send`` takes the same message and the code the user confirmed. It resolves the message
   again, refuses it unless the code matches (any change to the account, recipients, subject or
   body changes the code), checks that the write credential is the bound account, and sends
   once. With no clear answer it looks in Sent Items instead of retrying.

A reply is sent through a draft (W6): the reply is saved into Drafts, read back and compared with
the original (its whole text quoted, its formatting kept, its inline images kept with the same
bytes, no image turned into "[cid:...]" text), and only that checked draft is sent.
``create_draft`` runs the same check on a reply draft and reports it. If the check fails, nothing
is sent and the draft stays in Drafts for the user to look at.
"""

from __future__ import annotations

import hashlib
import json
import re
import tempfile
from collections import Counter
from datetime import UTC, datetime, timedelta
from pathlib import Path

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import (
    AccountMismatch,
    ConnectorError,
    InvalidRequest,
    NotFound,
    WriteOutcomeUnknown,
)
from outlook_connector.domain.models import (
    ADDRESS_PATTERN,
    MAX_RECIPIENTS,
    Attachment,
    DraftResult,
    EmailProposal,
    Message,
    MessageSummary,
    OutgoingMessage,
    Recipient,
    SendResult,
)
from outlook_connector.remote.ports import MailWriter
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.threads import base_subject

ADDRESS = re.compile(ADDRESS_PATTERN)
SENT_LOOKBACK = timedelta(minutes=5)


class Writes:
    def __init__(self, mailbox: Mailbox, writer: MailWriter, account: Account) -> None:
        self.mailbox = mailbox
        self.writer = writer
        self.account = account

    # ---------------------------------------------------------------- proposal

    async def propose(self, message: OutgoingMessage) -> EmailProposal:
        """Exactly what would be sent, validated, with its confirmation code. Changes nothing."""
        me = (self.account.username or "").lower()
        to, cc, bcc = (
            _addresses(message.to, "to"),
            _addresses(message.cc, "cc"),
            _addresses(message.bcc, "bcc"),
        )
        subject = message.subject
        if message.reply_to_message_id:
            original = await self.mailbox.message(message.reply_to_message_id)
            if not (to or cc or bcc):  # Outlook's defaults
                to, cc = _reply_recipients(original, me=me, reply_all=message.reply_all)
            if subject is None:
                subject = f"RE: {base_subject(original.subject) or ''}".rstrip()
        elif message.reply_all:
            raise InvalidRequest("reply_all needs reply_to_message_id.")
        recipients = to + cc + bcc
        if not recipients:
            raise InvalidRequest("Name at least one recipient (to, cc or bcc).")
        if len(recipients) > MAX_RECIPIENTS:
            raise InvalidRequest(f"At most {MAX_RECIPIENTS} recipients per message.")
        if len({a.lower() for a in recipients}) < len(recipients):
            raise InvalidRequest("A recipient is listed twice (across to, cc and bcc).")
        if not (subject or "").strip():
            raise InvalidRequest("A new message needs a subject.")
        if not message.body.strip():
            raise InvalidRequest("The body is empty.")
        fields = {
            "sender": self.account.username or "",
            "to": to,
            "cc": cc,
            "bcc": bcc,
            "subject": subject or "",
            "body": message.body,
            "reply_to_message_id": message.reply_to_message_id,
            "reply_all": message.reply_all,
            "quotes_original": bool(message.reply_to_message_id),
        }
        return EmailProposal.model_validate({**fields, "confirmation": self._code(fields)})

    def _code(self, fields: dict[str, object]) -> str:
        material = json.dumps(
            {"account": self.account.fingerprint, **fields}, sort_keys=True, ensure_ascii=False
        )
        return "SEND-" + hashlib.sha256(material.encode()).hexdigest()[:8].upper()

    # ---------------------------------------------------------------- draft

    async def create_draft(self, message: OutgoingMessage) -> DraftResult:
        """Save the message (or reply) into Drafts. Nothing is sent."""
        proposal = await self.propose(message)
        self.check_account()
        draft_id = await self.writer.create_draft(proposal)
        if not draft_id:
            raise WriteOutcomeUnknown(
                "Outlook saved the draft but did not report its id; look in Drafts before saving it again."
            )
        try:
            await self.mailbox.message(draft_id)
            verified = True
        except NotFound:
            verified = False
        result = DraftResult(id=draft_id, proposal=proposal, verified=verified)
        if proposal.reply_to_message_id and verified:
            problem = await self._reply_problem(draft_id, proposal.reply_to_message_id)
            result.history_intact, result.history_problem = problem is None, problem
        return result

    # ---------------------------------------------------------------- send

    async def send(self, message: OutgoingMessage, user_confirmation: str) -> SendResult:
        proposal = await self.propose(message)
        if user_confirmation.strip().upper() != proposal.confirmation:
            raise InvalidRequest(
                "Not sent: user_confirmation does not match this exact message (any change to the "
                "recipients, subject or body changes the code). Call propose_email, show the proposal "
                "to the user, and send only the message they confirmed, with its code."
            )
        self.check_account()
        started = datetime.now(UTC)
        try:
            if proposal.reply_to_message_id:
                await self._send_reply(proposal, proposal.reply_to_message_id)
            else:
                await self.writer.send(proposal)
        except WriteOutcomeUnknown as exc:
            try:
                found = await self._find_sent(proposal, since=started - SENT_LOOKBACK)
            except ConnectorError:  # the check failed: stay with "unknown", never a plain error
                found = None
            if found:
                return SendResult(
                    status="sent", sent_item_id=found.id,
                    detail="Outlook gave no clear answer, but the message is in Sent Items.",
                )  # fmt: skip
            return SendResult(
                status="unknown",
                detail=f"{exc} No copy is in Sent Items yet. Do not send again before the user has "
                "checked Outlook (Sent Items and Outbox).",
            )
        return SendResult(status="sent", detail="Sent; a copy is kept in Sent Items.")

    async def _send_reply(self, proposal: EmailProposal, original_id: str) -> None:
        """Save the reply as a draft, check it against the original, send exactly that draft."""
        try:
            draft_id = await self.writer.create_draft(proposal)
        except WriteOutcomeUnknown:
            raise ConnectorError(
                "Not sent: Outlook gave no clear answer while saving the reply as a draft. Look in "
                "Drafts before trying again."
            ) from None
        if not draft_id:
            raise ConnectorError(
                "Not sent: Outlook saved the reply as a draft but did not report its id. Look in Drafts."
            )
        problem = await self._reply_problem(draft_id, original_id)
        if problem:
            raise ConnectorError(
                f"Not sent: the reply draft did not keep the original message intact ({problem}). "
                f"The draft is in Drafts (id {draft_id}) for the user to check in Outlook."
            )
        await self.writer.send_draft(draft_id, proposal.subject)

    async def _reply_problem(self, draft_id: str, original_id: str) -> str | None:
        """Why the reply draft does not carry the original as received, or None when it does."""
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
        """SHA-256 of each inline image's bytes (downloaded to a temporary folder, then removed)."""
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

    async def _find_sent(self, proposal: EmailProposal, *, since: datetime) -> MessageSummary | None:
        """The proposal's copy in Sent Items: same subject and exactly the same To, Cc and Bcc."""
        sent = await self.mailbox.resolve_folder("sentitems")
        items, _ = await self.mailbox.reader.list_messages(
            folder_id=sent.id, since=since, until=None, page_size=50, page=None
        )
        candidates = [m.id for m in items if m.subject == proposal.subject]
        if not candidates:
            return None
        fetched = await self.mailbox.reader.get_messages(candidates)  # with Bcc, which summaries lack

        def addresses(recipients: list[Recipient]) -> set[str]:
            return {r.address.lower() for r in recipients if r.address}

        wanted = [{a.lower() for a in field} for field in (proposal.to, proposal.cc, proposal.bcc)]
        for message in fetched.messages.values():
            if message and [addresses(message.to), addresses(message.cc), addresses(message.bcc)] == wanted:
                return message
        return None


def _reply_recipients(original: Message, *, me: str, reply_all: bool) -> tuple[list[str], list[str]]:
    """Outlook's reply recipients: the sender (the original recipients when you sent it); with
    reply-all also the original To (in To) and Cc (in Cc). You are left out, except when nobody
    else is left: a reply to mail you sent only to yourself goes back to you."""

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

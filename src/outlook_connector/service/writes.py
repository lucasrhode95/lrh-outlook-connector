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
"""

from __future__ import annotations

import hashlib
import json
import re
from datetime import UTC, datetime, timedelta

from outlook_connector.auth.tokens import Account
from outlook_connector.domain.errors import AccountMismatch, InvalidRequest, NotFound, WriteOutcomeUnknown
from outlook_connector.domain.models import (
    ADDRESS_PATTERN,
    MAX_RECIPIENTS,
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
        self._check_account()
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
        return DraftResult(id=draft_id, proposal=proposal, verified=verified)

    # ---------------------------------------------------------------- send

    async def send(self, message: OutgoingMessage, user_confirmation: str) -> SendResult:
        proposal = await self.propose(message)
        if user_confirmation.strip().upper() != proposal.confirmation:
            raise InvalidRequest(
                "Not sent: user_confirmation does not match this exact message (any change to the "
                "recipients, subject or body changes the code). Call propose_email, show the proposal "
                "to the user, and send only the message they confirmed, with its code."
            )
        self._check_account()
        started = datetime.now(UTC)
        try:
            await self.writer.send(proposal)
        except WriteOutcomeUnknown as exc:
            found = await self._find_sent(proposal, since=started - SENT_LOOKBACK)
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

    def _check_account(self) -> None:
        claims = self.writer.account()
        if (claims.get("tid"), claims.get("oid")) != (self.account.tenant_id, self.account.object_id):
            raise AccountMismatch(
                "The write sign-in belongs to a different Microsoft account than the one this "
                "connector reads; nothing was written."
            )

    async def _find_sent(self, proposal: EmailProposal, *, since: datetime) -> MessageSummary | None:
        sent = await self.mailbox.resolve_folder("sentitems")
        items, _ = await self.mailbox.reader.list_messages(
            folder_id=sent.id, since=since, until=None, page_size=50, page=None
        )
        wanted = {a.lower() for a in proposal.to + proposal.cc}
        for item in items:
            got = {r.address.lower() for r in item.to + item.cc if r.address}
            if item.subject == proposal.subject and wanted <= got:
                return item
        return None


def _reply_recipients(original: Message, *, me: str, reply_all: bool) -> tuple[list[str], list[str]]:
    """Outlook's reply recipients: the sender (the original recipients when you sent it); with
    reply-all also the original To (in To) and Cc (in Cc). You are never among them."""

    def plain(recipients: list[Recipient]) -> list[str]:
        return [r.address for r in recipients if r.address and r.address.lower() != me]

    sender = original.sender.address if original.sender and original.sender.address else None
    to = plain(original.to) if sender is None or sender.lower() == me else [sender]
    cc: list[str] = []
    if reply_all:
        to += plain(original.to)
        cc = plain(original.cc)
    to = _unique(_addresses(to, "to"))
    return to, [a for a in _unique(_addresses(cc, "cc")) if a.lower() not in {t.lower() for t in to}]


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

"""Draft-first mail: explicit bodies, mandatory Graph read-back, existing-draft sending."""

from __future__ import annotations

import asyncio
import hashlib
import re
import tempfile
from collections import Counter
from html import escape
from pathlib import Path

from m365_auth import Account

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
    InlineImage,
    LocalAttachment,
    Message,
    OutgoingMessage,
    Recipient,
    SendResult,
)
from outlook_connector.remote.ports import MailWriter
from outlook_connector.service.conversations import base_subject
from outlook_connector.service.mailbox import Mailbox
from outlook_connector.service.passive_html import passive_start_tags
from outlook_connector.service.signatures import Signatures

ADDRESS = re.compile(ADDRESS_PATTERN)
MAX_SEND_BYTES = 34 * 1024 * 1024


class Writes:
    def __init__(
        self, mailbox: Mailbox, writer: MailWriter, account: Account, signatures: Signatures
    ) -> None:
        self.mailbox = mailbox
        self.writer = writer
        self.account = account
        self.signatures = signatures

    async def create_draft(self, message: OutgoingMessage) -> DraftResult:
        """Entry point: validate body, envelope and reply arguments; save once and read back."""
        resolved, original = await self._resolve(message)
        self.check_account()
        draft_id = await self.writer.create_draft(resolved)
        if not draft_id:
            raise WriteOutcomeUnknown(
                "Outlook did not report a draft id; check Drafts before creating again."
            )
        failed_attachment: LocalAttachment | None = None
        try:
            for attachment in resolved.local_attachments:
                failed_attachment = attachment
                current = await self.mailbox.message(draft_id)
                if not current.revision:
                    raise Upstream("Outlook did not report the saved draft version for attachment upload.")
                await self.writer.add_local_attachment(draft_id, current.revision, attachment)
        except ConnectorError as exc:
            if failed_attachment is None:
                raise
            deleted = await self._discard_draft(draft_id)
            disposition = (
                "The draft was moved to Deleted Items."
                if deleted
                else ("The draft could not be moved to Deleted Items; check Drafts before continuing.")
            )
            raise Upstream(
                f"Could not attach local file '{failed_attachment.name}': {exc} {disposition}"
            ) from exc
        result = await self._read_back(
            draft_id,
            expected=resolved.inline_images,
            expected_local=resolved.local_attachments,
        )
        if resolved.local_attachments:
            failed = next(
                (item for item in resolved.local_attachments if not _attachment_is_confirmed(item, result)),
                None,
            )
            if failed:
                deleted = await self._discard_draft(draft_id)
                disposition = (
                    "The draft was moved to Deleted Items."
                    if deleted
                    else ("The draft could not be moved to Deleted Items; check Drafts before continuing.")
                )
                detail = next(
                    (finding for finding in result.findings if failed.name in finding),
                    "Graph could not confirm the attachment in the saved draft.",
                )
                raise Upstream(f"Could not verify local file '{failed.name}': {detail} {disposition}")
        if resolved.reply_to_message_id and result.verified:
            try:
                assert original is not None
                problem = await self._reply_problem(original, result)
            except ConnectorError as exc:
                problem = f"Reply history verification failed: {exc}"
            result.history_intact, result.history_problem = problem is None, problem
        return result

    async def send_draft(self, draft_id: str) -> SendResult:
        """Entry point: require a server draft and the bound account. Only an id, no content.

        The host obtains human approval; the connector never reconstructs or retries the message.
        """
        draft = await self._draft(draft_id)
        if draft.size is None:
            raise Upstream("Outlook did not report this draft's size, so it was not sent.")
        if draft.size > MAX_SEND_BYTES:
            raise InvalidRequest(
                f"This draft is too large to send safely ({draft.size:,} bytes; connector limit is 34 MiB)."
            )
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

    async def _resolve(self, message: OutgoingMessage) -> tuple[DraftMessage, Message | None]:
        """Assumes (not re-checked here): called only by create_draft; validates its outgoing input."""
        page = _body(message.text_body, message.html_body)
        to, cc, bcc = (
            _addresses(message.to, "to"),
            _addresses(message.cc, "cc"),
            _addresses(message.bcc, "bcc"),
        )
        subject = message.subject
        original: Message | None = None
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
        inline_images: list[InlineImage] = []
        local_attachments = _local_attachments(message.attachments)
        if message.include_signature:
            selected = await self.signatures.resolve_for_draft(
                message.signature, is_reply=bool(message.reply_to_message_id)
            )
            if selected:
                page += selected.html
                inline_images = selected.images
                if len(page) > MAX_BODY_CHARS:
                    raise InvalidRequest(f"Body must be at most {MAX_BODY_CHARS} characters.")
        draft = DraftMessage(
            to=to,
            cc=cc,
            bcc=bcc,
            subject=subject or "",
            html_body=page,
            reply_to_message_id=message.reply_to_message_id,
            reply_all=message.reply_all,
            inline_images=inline_images,
            local_attachments=local_attachments,
        )
        return draft, original

    async def _read_back(
        self,
        draft_id: str,
        *,
        expected: list[InlineImage] | None = None,
        expected_local: list[LocalAttachment] | None = None,
    ) -> DraftResult:
        """Assumes (not re-checked here): reliable id from a single successful write. Retries reads only."""
        for attempt in range(3):
            try:
                text = await self.mailbox.reader.get_message(draft_id)
                page = await self.mailbox.reader.get_message(draft_id, body_format="html")
                findings = []
                if not (text.body_text or "").strip():
                    findings.append("The server draft body is empty or missing.")
                cids = re.findall(r"cid:([^\"' >]+)", page.body_html or "", re.IGNORECASE)
                if cids or expected or expected_local:
                    attachments = await self.mailbox.reader.list_attachments(draft_id)
                    present = {
                        attachment.content_id
                        for attachment in attachments
                        if attachment.is_inline and attachment.content_id
                    }
                    if any(cid not in present for cid in cids):
                        findings.append("The server draft references missing inline images.")
                    if expected:
                        by_cid = {
                            attachment.content_id: attachment
                            for attachment in attachments
                            if attachment.is_inline and attachment.content_id
                        }
                        with tempfile.TemporaryDirectory(prefix="outlook-signature-check-") as folder:
                            for index, image in enumerate(expected):
                                attachment = by_cid.get(image.content_id)
                                if attachment is None:
                                    findings.append(
                                        "The server draft is missing an expected signature image."
                                    )
                                    continue
                                target = Path(folder) / str(index)
                                await self.mailbox.reader.download_attachment(draft_id, attachment.id, target)
                                if (
                                    hashlib.sha256(target.read_bytes()).digest()
                                    != hashlib.sha256(image.content).digest()
                                ):
                                    findings.append("A signature image changed in the server draft.")
                    if expected_local:
                        found = Counter(
                            (attachment.name, attachment.size)
                            for attachment in attachments
                            if not attachment.is_inline
                        )
                        for local in expected_local:
                            key = (local.name, local.size)
                            if found[key]:
                                found[key] -= 1
                            else:
                                findings.append(
                                    f"Attachment '{local.name}' is missing or has a different size."
                                )
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

    async def _discard_draft(self, draft_id: str) -> bool:
        """Move a partial draft to Deleted Items; never purge it or retry an unclear result."""
        try:
            return (await self.writer.delete([draft_id])).get(draft_id) is None
        except ConnectorError:
            return False

    async def _reply_problem(self, original: Message, result: DraftResult) -> str | None:
        """Check quoted history using the original and read-back draft already fetched this call.

        Assumes (not re-checked here): original was read while resolving the reply and result is the
        verified read-back of the draft this call just saved.
        """
        if result.message is None:
            return "the draft or the original could not be read back"
        original_page = await self.mailbox.reader.get_message(original.id, body_format="html")
        if _flat(original.body_text) not in _flat(result.text_body):
            return "the quoted original is not the original's full text"
        # Exchange re-wraps the original's HTML when it quotes it, so the HTML is not compared as
        # text; its structure is: every list, table, emphasis, link and image reference must survive.
        lost = _structure(original_page.body_html) - _structure(result.html_body)
        if lost:
            return "the quoted original lost formatting (" + ", ".join(sorted(lost.elements())) + ")"
        found, failed = await self.mailbox.reader.list_attachments_many([original.id, result.id])
        if failed:
            return "its attachments could not be listed"
        hashes = {
            original.id: await self._inline_hashes(original.id, found.get(original.id, [])),
            result.id: await self._inline_hashes(result.id, found.get(result.id, [])),
        }
        missing = Counter(hashes[original.id]) - Counter(hashes[result.id])
        if missing:
            return f"{missing.total()} inline image(s) of the original are missing or changed"
        if "[cid:" in _flat(re.sub(r"<[^>]+>", " ", result.html_body or "")):
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
                "The Outlook sign-in belongs to a different Microsoft account than the one this "
                "connector reads; nothing was written."
            )


def _local_attachments(paths: list[str]) -> list[LocalAttachment]:
    """Validate every path before CreateItem so a missing local file never leaves a draft behind."""
    out: list[LocalAttachment] = []
    for value in paths:
        path = Path(value)
        if not value.strip() or not path.name:
            raise InvalidRequest("Attachment paths must name a local file.")
        try:
            if not path.is_file():
                raise InvalidRequest(f"Cannot read local attachment '{path}': it is not a file.")
            size = path.stat().st_size
            with path.open("rb"):
                pass
        except OSError as exc:
            raise InvalidRequest(f"Cannot read local attachment '{path}': {exc}") from None
        out.append(LocalAttachment(path=str(path), name=path.name, size=size))
    return out


def _attachment_is_confirmed(item: LocalAttachment, result: DraftResult) -> bool:
    missing = f"Attachment '{item.name}' is missing or has a different size."
    return result.verified and missing not in result.findings


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
    to = _addresses(to, "to")
    return to, [a for a in _addresses(cc, "cc") if a.lower() not in {t.lower() for t in to}]


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
    passive_start_tags(value)
    return value


def _keep_spaces(line: str) -> str:
    line = line.replace("\t", "&nbsp;" * 4)
    line = re.sub(r" {2,}", lambda run: "&nbsp;" * (len(run.group()) - 1) + " ", line)
    return "&nbsp;" + line[1:] if line.startswith(" ") else line

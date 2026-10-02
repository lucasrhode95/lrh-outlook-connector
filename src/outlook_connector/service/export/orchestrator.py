"""The single export path (requirements v4 §10), used by both the local UI and MCP.

selection (threads + messages) → bodies (Graph $batch, retained copies for server-deleted mail)
→ attachments (policy) → grouping (per thread / all / none) → TXT rendering → packaging.
"""

from __future__ import annotations

import asyncio
import shutil
import tempfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path, PurePath

from outlook_connector.domain.errors import ConnectorError, InvalidRequest, NotFound, Upstream
from outlook_connector.domain.models import (
    EXPORT_MAX_MESSAGES,
    Attachment,
    ExportArtifact,
    ExportRequest,
    Message,
    MessageSummary,
)
from outlook_connector.service.export import attachments as policy
from outlook_connector.service.export.formatter import RenderedMessage, body_text, render_file
from outlook_connector.service.export.packaging import TextFile, package
from outlook_connector.service.threads import Threads, base_subject, oldest_first

RANGE_PAGE = 200
DELETED_FOLDERS = ("deleteditems", "junkemail")
SENT_FOLDERS = ("sentitems", "drafts", "outbox")
DELETED_OR_JUNK = "in Deleted Items or Junk Email (include_deleted_items=false)"
NOT_RECEIVED = "in Sent Items, Drafts or Outbox (received_only=true)"
NEVER_RETAINED = "deleted on the server and never retained by this app"

Downloaded = list[tuple[Attachment, Path | None]]  # None: unavailable


@dataclass
class Selection:
    summaries: list[MessageSummary]
    excluded: dict[str, int] = field(default_factory=dict)  # reason -> messages left out


@dataclass
class Fetched:
    bodies: dict[str, Message | None]
    body_failures: dict[str, str]  # message id -> why its body could not be fetched
    attachments: dict[str, list[Attachment]]
    attachment_failures: dict[str, str]  # message id -> why its attachments could not be listed


class Exports:
    def __init__(self, threads: Threads) -> None:
        self.threads = threads
        self.mailbox = threads.mailbox
        self.reader = threads.mailbox.reader

    async def export(self, request: ExportRequest) -> ExportArtifact:
        if not request.conversation_ids and not request.message_ids and not request.by_range:
            raise InvalidRequest(
                "Select at least one conversation or message, or a range (since/until/folder/received_only)."
            )
        selection = await self._select(request)
        summaries = selection.summaries
        if not summaries:
            raise InvalidRequest(
                f"Nothing to export: the selection holds no messages{_excluded_note(selection)}."
            )
        bodies, body_failures = await self.threads.bodies(summaries)
        for summary in summaries:  # deleted on the server before this app ever read the body
            if bodies.get(summary.id) is None and summary.id not in body_failures:
                body_failures[summary.id] = NEVER_RETAINED
        found, attachment_failures = await self._attachments(
            summaries, bodies, inline=request.include_attachments, skip=set(body_failures)
        )
        fetched = Fetched(bodies, body_failures, found, attachment_failures)
        workdir = Path(tempfile.mkdtemp(prefix="outlook-export-"))
        try:
            downloads = (
                await self._download(summaries, bodies, found, request, workdir)
                if request.include_attachments
                else {}
            )
            files = self._render(summaries, fetched, downloads, request, selection.excluded)
            base = files[0].folder if len(files) == 1 else _range_name(summaries)
            result = package(files, base_name=base)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
        unavailable = sum(1 for items in downloads.values() for _, path in items if path is None)
        return ExportArtifact(
            path=str(result.path),
            filename=result.filename,
            content_type=result.content_type,
            size=result.size,
            message_count=len(summaries),
            text_files=len(files),
            attachment_files=sum(len(f.attachments) for f in files),
            attachments_unavailable=unavailable + len(attachment_failures),
            messages_excluded=selection.excluded,
            messages_unavailable=len(body_failures),
            unavailable_message_ids=list(body_failures),
        )

    # ---------------------------------------------------------------- selection

    async def _select(self, request: ExportRequest) -> Selection:
        selected: dict[str, MessageSummary] = {}
        excluded: dict[str, int] = defaultdict(int)
        for conversation_id in dict.fromkeys(request.conversation_ids):
            items, left_out, truncated = await self.threads.messages(
                conversation_id, include_deleted_items=request.include_deleted_items
            )
            if truncated:
                raise InvalidRequest(
                    f"Conversation {conversation_id} is larger than the server listing limit; "
                    "export its messages by message id instead."
                )
            if left_out:
                excluded[DELETED_OR_JUNK] += left_out
            for item in items:
                selected.setdefault(item.id, item)
            self._check_limit(selected, request.limit)
        if request.by_range:
            await self._select_range(request, selected, excluded)
        explicit = [mid for mid in dict.fromkeys(request.message_ids) if mid not in selected]
        known = self.mailbox.store.summaries(explicit)
        selected.update(known)
        missing = [mid for mid in explicit if mid not in known]
        if missing:
            selected.update(await self._summaries(missing))
        self._check_limit(selected, request.limit)
        summaries = sorted(await self.mailbox.decorate(list(selected.values())), key=oldest_first)
        return Selection(summaries, dict(excluded))

    async def _select_range(
        self, request: ExportRequest, selected: dict[str, MessageSummary], excluded: dict[str, int]
    ) -> None:
        """Every message in the window (newest first from the server), minus the excluded folders."""
        folders = await self.mailbox.folder_map()
        skip: dict[str, str] = {}
        if not request.include_deleted_items:
            skip |= {f.id: DELETED_OR_JUNK for f in folders.values() if f.well_known in DELETED_FOLDERS}
        if request.received_only:
            skip |= {f.id: NOT_RECEIVED for f in folders.values() if f.well_known in SENT_FOLDERS}
        cursor: str | None = None
        while True:
            page = await self.mailbox.list_messages(
                folder=request.folder,
                since=request.since,
                until=request.until,
                limit=RANGE_PAGE,
                cursor=cursor,
            )
            for item in page.items:
                reason = skip.get(item.folder_id or "")
                if reason:
                    excluded[reason] += 1
                else:
                    selected.setdefault(item.id, item)
            self._check_limit(selected, request.limit, more=page.cursor is not None)
            cursor = page.cursor
            if cursor is None:
                return

    @staticmethod
    def _check_limit(selected: dict[str, MessageSummary], limit: int, *, more: bool = False) -> None:
        if len(selected) > limit:
            raise InvalidRequest(
                f"The selection holds {'more than ' if more else ''}{len(selected)} messages, above "
                f"the limit of {limit} (at most {EXPORT_MAX_MESSAGES}). Narrow the date range or "
                "split the export."
            )

    async def _summaries(self, message_ids: list[str]) -> dict[str, MessageSummary]:
        """Summaries for ids this app has never seen, in batches (bodies are stored on the way)."""
        fetched = await self.reader.get_messages(message_ids)
        if fetched.failed:
            first = next(iter(fetched.failed.values()))
            raise Upstream(f"{len(fetched.failed)} selected message(s) could not be read. First: {first}")
        gone = [mid for mid, m in fetched.messages.items() if m is None]
        if gone:
            raise NotFound(
                f"{len(gone)} selected message(s) exist neither on the server nor locally: {gone[0]}"
            )
        messages = [m for m in fetched.messages.values() if m is not None]
        self.mailbox.store.save_messages(messages)
        return {m.id: MessageSummary.model_validate(m.model_dump()) for m in messages}

    async def _attachments(
        self,
        summaries: list[MessageSummary],
        bodies: dict[str, Message | None],
        *,
        inline: bool,
        skip: set[str],
    ) -> tuple[dict[str, list[Attachment]], dict[str, str]]:
        # Graph reports hasAttachments=false when a message has only inline attachments, so when
        # files are exported (inline images included) every live message is asked, in batches.
        # ``bodies`` has already marked messages that disappeared from the server as deleted.
        live = [
            m.id for m in summaries if not m.is_deleted and m.id not in skip and (inline or m.has_attachments)
        ]
        found, failed = await self.reader.list_attachments_many(live) if live else ({}, {})
        retain = []
        for summary in summaries:
            body = bodies.get(summary.id)
            if summary.id in found and body is not None:
                body.attachments = found[summary.id]
                retain.append(body)
            elif body is not None and body.attachments and summary.id not in failed:
                found[summary.id] = body.attachments  # retained metadata of a server-deleted message
        self.mailbox.store.save_messages(retain)
        return found, failed

    # ---------------------------------------------------------------- attachments

    async def _download(
        self,
        summaries: list[MessageSummary],
        bodies: dict[str, Message | None],
        found: dict[str, list[Attachment]],
        request: ExportRequest,
        workdir: Path,
    ) -> dict[str, Downloaded]:
        inline_ids = {
            mid: [a.id for a in items if a.is_inline and a.kind == "file"] for mid, items in found.items()
        }
        needs_html = [m.id for m in summaries if inline_ids.get(m.id) and not m.is_deleted]
        html = (await self.reader.get_messages(needs_html, body_format="html")).messages if needs_html else {}
        content_ids = dict(
            zip(
                needs_html,
                await asyncio.gather(
                    *(self.reader.attachment_content_ids(mid, inline_ids[mid]) for mid in needs_html)
                ),
                strict=True,
            )
        )

        jobs: list[tuple[str, Attachment, Path | None]] = []
        for index, summary in enumerate(summaries):
            page = html.get(summary.id)
            page_html = None
            if page is not None:
                page_html = page.unique_body_html if request.body == "unique" else page.body_html
            elif (retained := bodies.get(summary.id)) is not None:
                page_html = retained.unique_body_html if request.body == "unique" else retained.body_html
            for position, attachment in enumerate(found.get(summary.id, [])):
                cid = content_ids.get(summary.id, {}).get(attachment.id)
                if policy.wanted(attachment, content_id=cid, body_html=page_html):
                    target = None if summary.is_deleted else workdir / f"{index}-{position}"
                    jobs.append((summary.id, attachment, target))

        async def fetch(message_id: str, attachment: Attachment, target: Path | None) -> Path | None:
            if target is None:
                return None
            try:
                await self.reader.download_attachment(message_id, attachment.id, target)
            except ConnectorError:
                return None
            return target

        results = await asyncio.gather(*(fetch(*job) for job in jobs))
        downloads: dict[str, Downloaded] = defaultdict(list)
        for (message_id, attachment, _), path in zip(jobs, results, strict=True):
            downloads[message_id].append((attachment, path))
        return downloads

    # ---------------------------------------------------------------- rendering

    def _render(
        self,
        summaries: list[MessageSummary],
        fetched: Fetched,
        downloads: dict[str, Downloaded],
        request: ExportRequest,
        excluded: dict[str, int],
    ) -> list[TextFile]:
        found = fetched.attachments
        notes = [f"Left out: {count} message(s) {reason}." for reason, count in excluded.items() if count]
        if fetched.body_failures:
            notes.append(
                f"Unavailable: {len(fetched.body_failures)} message body(ies); they are marked below."
            )
        files: list[TextFile] = []
        taken: set[str] = set()
        for title, base, members in _groups(summaries, request.combine):
            name = policy.dedupe(policy.safe_name(f"{base}.txt", fallback="export.txt"), taken)
            folder = PurePath(name).stem
            file = TextFile(name=name, text="")
            names: set[str] = set()
            rendered = []
            for message in members:
                lines = []
                if message.id in fetched.attachment_failures:
                    lines.append(
                        f"[Attachments could not be listed: {fetched.attachment_failures[message.id]}]"
                    )
                if request.include_attachments:
                    for attachment, path in downloads.get(message.id, []):
                        label = attachment.name or "attachment"
                        if path is None:
                            lines.append(f"[Attachment unavailable: {label}]")
                            continue
                        final = policy.dedupe(
                            policy.safe_name(
                                attachment.name,
                                fallback=f"attachment-{len(names) + 1}",
                                eml=attachment.kind == "item",
                            ),
                            names,
                        )
                        file.attachments.append((final, path))
                        lines.append(f"{folder}/{final} ({policy.size_label(path.stat().st_size)})")
                    lines += [
                        f"{a.name} (cloud link, not downloaded)"
                        for a in found.get(message.id, [])
                        if a.kind == "reference"
                    ]
                else:
                    lines += [
                        f"{a.name} ({policy.size_label(a.size)})"
                        for a in found.get(message.id, [])
                        if policy.listed(a)
                    ]
                failure = fetched.body_failures.get(message.id)
                text = (
                    f"(Content unavailable: {failure})"
                    if failure
                    else body_text(fetched.bodies.get(message.id), request.body)
                )
                rendered.append(RenderedMessage(message, text, lines))
            file.text = render_file(
                title, rendered, body_kind=request.body, sections=request.combine == "all", notes=notes
            )
            files.append(file)
        return files


def _excluded_note(selection: Selection) -> str:
    parts = [f"{count} {reason}" for reason, count in selection.excluded.items() if count]
    return f" ({'; '.join(parts)} left out)" if parts else ""


def _day(message: MessageSummary) -> str:
    stamp = message.received_at or message.sent_at
    return stamp.strftime("%Y-%m-%d") if stamp else "undated"


def _range_name(summaries: list[MessageSummary]) -> str:
    return f"outlook-export {_day(summaries[0])} to {_day(summaries[-1])}" if summaries else "outlook-export"


def _groups(summaries: list[MessageSummary], combine: str) -> list[tuple[str, str, list[MessageSummary]]]:
    """(title, file base name, members oldest first), groups ordered by their first message."""
    if combine == "all":
        conversations = len({m.conversation_id or m.id for m in summaries})
        return [(f"Outlook export: {conversations} conversation(s)", _range_name(summaries), summaries)]
    if combine == "none":
        return [
            (f"Message: {m.subject or '(no subject)'}", f"{_day(m)} {m.subject or 'message'}", [m])
            for m in summaries
        ]
    groups: dict[str, list[MessageSummary]] = defaultdict(list)
    for message in summaries:
        groups[message.conversation_id or message.id].append(message)
    out = []
    for members in sorted(groups.values(), key=lambda g: oldest_first(g[0])):
        subject = base_subject(members[0].subject) or "conversation"
        out.append((f"Conversation: {subject}", f"{_day(members[0])} {subject}", members))
    return out

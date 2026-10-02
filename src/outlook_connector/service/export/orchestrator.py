"""The single export path (requirements v4 §10), used by both the local UI and MCP.

selection (conversations + messages + a range) → copies merged → bodies (Graph $batch, retained
copies for server-deleted mail) → attachments (policy) → grouping (per thread / all / none) →
TXT or JSONL rendering → packaging.
"""

from __future__ import annotations

import asyncio
import hashlib
import shutil
import tempfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from outlook_connector.domain.errors import ConnectorError, InvalidRequest, NotFound, Upstream
from outlook_connector.domain.models import (
    EXCLUSION_TEXT,
    EXPORT_MAX_MESSAGES,
    Attachment,
    ExportArtifact,
    ExportRequest,
    Message,
    MessageSummary,
)
from outlook_connector.service.export import attachments as policy
from outlook_connector.service.export.formatter import RenderedMessage, body_text, jsonl_record, render_file
from outlook_connector.service.export.packaging import TextFile, package
from outlook_connector.service.mailbox import unavailable
from outlook_connector.service.threads import Threads, base_subject, oldest_first

RANGE_PAGE = 200

Downloaded = list[tuple[Attachment, Path | None]]  # None: unavailable


@dataclass
class Selection:
    summaries: list[MessageSummary]
    excluded: dict[str, int] = field(default_factory=dict)  # ExclusionReason -> messages left out
    known: dict[str, Message] = field(default_factory=dict)  # fetched with bodies while selecting


@dataclass
class Fetched:
    bodies: dict[str, Message]
    missing: dict[str, str]  # message id -> why it has no body
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
                f"Nothing to export: the selection holds no messages{_excluded_note(selection.excluded)}."
            )
        bodies, missing = await self.threads.bodies(summaries, known=selection.known)
        found, attachment_failures = await self._attachments(
            summaries, bodies, inline=request.include_attachments, skip=set(missing)
        )
        fetched = Fetched(bodies, missing, found, attachment_failures)
        workdir = Path(tempfile.mkdtemp(prefix="outlook-export-"))
        try:
            downloads = (
                await self._download(summaries, bodies, found, request, workdir)
                if request.include_attachments
                else {}
            )
            files = self._render(summaries, fetched, downloads, request, selection)
            base = files[0].folder if len(files) == 1 else _range_name(summaries)
            result = package(files, base_name=base)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
        unavailable_files = sum(1 for items in downloads.values() for _, path in items if path is None)
        return ExportArtifact(
            path=str(result.path),
            filename=result.filename,
            content_type=result.content_type,
            size=result.size,
            message_count=len(summaries),
            text_files=len(files),
            attachment_files=sum(len(f.attachments) for f in files),
            attachments_unavailable=unavailable_files,
            attachment_listing_failures=len(attachment_failures),
            messages_excluded=selection.excluded,
            messages_unavailable=len(missing),
            unavailable_message_ids=list(missing),
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
            _add(excluded, left_out)
            for item in items:
                selected.setdefault(item.id, item)
            _check_limit(selected, request.limit)
        if request.by_range:
            await self._select_range(request, selected, excluded)
        explicit = [mid for mid in dict.fromkeys(request.message_ids) if mid not in selected]
        cached = self.mailbox.store.summaries(explicit)
        selected.update(cached)
        known = await self._fetch_unknown([mid for mid in explicit if mid not in cached])
        selected.update({mid: MessageSummary.model_validate(m.model_dump()) for mid, m in known.items()})
        merged, _ = await self.mailbox.finish(selected.values())  # copies across conversations and pages
        _check_limit({m.id: m for m in merged}, request.limit)
        return Selection(sorted(merged, key=oldest_first), dict(excluded), known)

    async def _select_range(
        self, request: ExportRequest, selected: dict[str, MessageSummary], excluded: dict[str, int]
    ) -> None:
        """Every message in the window, page by page, with the listing's scope rules. Copies on
        different pages are all kept here, so the final merge sees them and names every folder in
        ``also_in``."""
        cursor: str | None = None
        while True:
            page = await self.mailbox.list_messages(
                folder=request.folder,
                since=request.since,
                until=request.until,
                limit=RANGE_PAGE,
                cursor=cursor,
                received_only=request.received_only,
                include_deleted_items=request.include_deleted_items,
                skip_returned_copies=False,
            )
            _add(excluded, page.coverage.excluded)
            for item in page.items:
                selected.setdefault(item.id, item)
            unique = {m.internet_message_id or m.id: m for m in selected.values()}  # copies count once
            _check_limit(unique, request.limit, more=page.cursor is not None)
            cursor = page.cursor
            if cursor is None:
                return

    async def _fetch_unknown(self, message_ids: list[str]) -> dict[str, Message]:
        """Messages this app has never seen, in batches; their bodies are reused for the export."""
        if not message_ids:
            return {}
        fetched = await self.reader.get_messages(message_ids)
        if fetched.failed:
            first = next(iter(fetched.failed.values()))
            raise Upstream(f"{len(fetched.failed)} selected message(s) could not be read. First: {first}")
        gone = [mid for mid, m in fetched.messages.items() if m is None]
        if gone:
            raise NotFound(
                f"{len(gone)} selected message(s) exist neither on the server nor locally: {gone[0]}"
            )
        messages = {mid: m for mid, m in fetched.messages.items() if m is not None}
        self.mailbox.store.save_messages(messages.values())
        return messages

    async def _attachments(
        self,
        summaries: list[MessageSummary],
        bodies: dict[str, Message],
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
        bodies: dict[str, Message],
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
            page = html.get(summary.id) or bodies.get(summary.id)
            page_html = None
            if page is not None:
                page_html = page.unique_body_html if request.body == "unique" else page.body_html
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
        selection: Selection,
    ) -> list[TextFile]:
        notes = [
            f"Left out: {count} message(s) {EXCLUSION_TEXT[key]}."
            for key, count in selection.excluded.items()
        ]
        if fetched.missing:
            notes.append(f"Unavailable: {len(fetched.missing)} message body(ies); they are marked below.")
        combine = "all" if request.format == "jsonl" else request.combine
        suffix = ".jsonl" if request.format == "jsonl" else ".txt"
        files: list[TextFile] = []
        taken: set[str] = set()
        for title, base, members in _groups(summaries, combine):
            name = policy.dedupe(policy.safe_name(f"{base}{suffix}", fallback=f"export{suffix}"), taken)
            file = TextFile(name=name, text="")
            names: set[str] = set()
            stored: dict[str, str] = {}  # content digest -> name in this file's folder
            rendered = [self._message(m, fetched, downloads, request, file, names, stored) for m in members]
            if request.format == "jsonl":
                file.text = "".join(jsonl_record(r, body_kind=request.body) + "\n" for r in rendered)
            else:
                file.text = render_file(
                    title, rendered, body_kind=request.body, sections=combine == "all", notes=notes
                )
            files.append(file)
        return files

    def _message(
        self,
        message: MessageSummary,
        fetched: Fetched,
        downloads: dict[str, Downloaded],
        request: ExportRequest,
        file: TextFile,
        names: set[str],
        stored: dict[str, str],
    ) -> RenderedMessage:
        """One message's text, attachment lines (TXT) and attachment records (JSONL)."""
        lines: list[str] = []
        records: list[dict[str, Any]] = []
        if message.id in fetched.attachment_failures:
            lines.append(f"[Attachments could not be listed: {fetched.attachment_failures[message.id]}]")
        listed = fetched.attachments.get(message.id, [])
        if request.include_attachments:
            for attachment, path in downloads.get(message.id, []):
                label = attachment.name or "attachment"
                if path is None:
                    lines.append(f"[Attachment unavailable: {label}]")
                    records.append(_attachment_record(attachment, unavailable=True))
                    continue
                # The same bytes (a signature logo on every message) are stored once; every
                # message that carries them points to that file.
                digest = hashlib.sha256(path.read_bytes()).hexdigest()
                final = stored.get(digest)
                if final is None:
                    final = policy.dedupe(
                        policy.safe_name(
                            attachment.name,
                            fallback=f"attachment-{len(names) + 1}",
                            eml=attachment.kind == "item",
                        ),
                        names,
                    )
                    stored[digest] = final
                    file.attachments.append((final, path))
                lines.append(f"{file.folder}/{final} ({policy.size_label(path.stat().st_size)})")
                records.append(_attachment_record(attachment, file=f"{file.folder}/{final}"))
            for a in listed:
                if a.kind == "reference":
                    lines.append(f"{a.name} (cloud link, not downloaded)")
                    records.append(_attachment_record(a))
        else:
            for a in listed:
                if policy.listed(a):
                    lines.append(f"{a.name} ({policy.size_label(a.size)})")
                    records.append(_attachment_record(a))
        reason = fetched.missing.get(message.id)
        text = unavailable(reason) if reason else body_text(fetched.bodies[message.id], request.body)
        return RenderedMessage(message, text, lines, unavailable=reason, attachments=records)


def _attachment_record(
    a: Attachment, *, file: str | None = None, unavailable: bool = False
) -> dict[str, Any]:
    record: dict[str, Any] = {"name": a.name, "size": a.size, "kind": a.kind, "inline": a.is_inline}
    if file:
        record["file"] = file
    if unavailable:
        record["unavailable"] = True
    return record


def _add(total: dict[str, int], counts: dict[str, int]) -> None:
    for key, count in counts.items():
        total[key] += count


def _check_limit(selected: dict[str, MessageSummary], limit: int, *, more: bool = False) -> None:
    if len(selected) > limit:
        raise InvalidRequest(
            f"The selection holds {'more than ' if more else ''}{len(selected)} messages, above "
            f"the limit of {limit} (at most {EXPORT_MAX_MESSAGES}). Narrow the date range or "
            "split the export."
        )


def _excluded_note(excluded: dict[str, int]) -> str:
    parts = [f"{count} {EXCLUSION_TEXT[key]}" for key, count in excluded.items() if count]
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

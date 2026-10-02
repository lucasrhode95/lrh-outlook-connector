"""The single export path (requirements v4 §10), used by both the local UI and MCP.

selection (threads + messages) → bodies (Graph $batch, retained copies for server-deleted mail)
→ attachments (policy) → grouping (per thread / all / none) → TXT rendering → packaging.
"""

from __future__ import annotations

import asyncio
import shutil
import tempfile
from collections import defaultdict
from pathlib import Path, PurePath

from outlook_connector.domain.errors import ConnectorError, InvalidRequest
from outlook_connector.domain.models import Attachment, ExportArtifact, ExportRequest, Message, MessageSummary
from outlook_connector.service.export import attachments as policy
from outlook_connector.service.export.formatter import RenderedMessage, body_text, render_file
from outlook_connector.service.export.packaging import TextFile, package
from outlook_connector.service.threads import Threads, base_subject, oldest_first

MAX_MESSAGES = 2000

Downloaded = list[tuple[Attachment, Path | None]]  # None: unavailable


class Exports:
    def __init__(self, threads: Threads) -> None:
        self.threads = threads
        self.mailbox = threads.mailbox
        self.reader = threads.mailbox.reader

    async def export(self, request: ExportRequest) -> ExportArtifact:
        if not request.conversation_ids and not request.message_ids:
            raise InvalidRequest("Select at least one conversation or message to export.")
        summaries = await self._select(request)
        bodies = await self.threads.bodies(summaries)
        found = await self._attachments(summaries, bodies, inline=request.include_attachments)
        workdir = Path(tempfile.mkdtemp(prefix="outlook-export-"))
        try:
            downloads = (
                await self._download(summaries, bodies, found, request, workdir)
                if request.include_attachments
                else {}
            )
            files = self._render(summaries, bodies, found, downloads, request)
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
            attachments_unavailable=unavailable,
        )

    # ---------------------------------------------------------------- selection

    async def _select(self, request: ExportRequest) -> list[MessageSummary]:
        selected: dict[str, MessageSummary] = {}
        for conversation_id in dict.fromkeys(request.conversation_ids):
            items, _, truncated = await self.threads.messages(
                conversation_id, include_deleted_items=request.include_deleted_items
            )
            if truncated:
                raise InvalidRequest(
                    f"Conversation {conversation_id} is larger than the server listing limit; "
                    "export its messages by message id instead."
                )
            for item in items:
                selected.setdefault(item.id, item)
        explicit = [mid for mid in dict.fromkeys(request.message_ids) if mid not in selected]
        known = self.mailbox.store.summaries(explicit)
        for message_id in explicit:
            selected[message_id] = known.get(message_id) or MessageSummary.model_validate(
                (await self.mailbox.message(message_id)).model_dump()
            )
        if len(selected) > MAX_MESSAGES:
            raise InvalidRequest(f"An export is limited to {MAX_MESSAGES} messages; narrow the selection.")
        return sorted(await self.mailbox.decorate(list(selected.values())), key=oldest_first)

    async def _attachments(
        self, summaries: list[MessageSummary], bodies: dict[str, Message | None], *, inline: bool
    ) -> dict[str, list[Attachment]]:
        # Graph reports hasAttachments=false when a message has only inline attachments, so when
        # files are exported (inline images included) every live message is asked, in batches.
        # ``bodies`` has already marked messages that disappeared from the server as deleted.
        live = [m.id for m in summaries if not m.is_deleted and (inline or m.has_attachments)]
        found = await self.reader.list_attachments_many(live) if live else {}
        retain = []
        for summary in summaries:
            body = bodies.get(summary.id)
            if summary.id in found and body is not None:
                body.attachments = found[summary.id]
                retain.append(body)
            elif body is not None and body.attachments:
                found[summary.id] = body.attachments  # retained metadata of a server-deleted message
        self.mailbox.store.save_messages(retain)
        return found

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
        html = await self.reader.get_messages(needs_html, body_format="html") if needs_html else {}
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
        bodies: dict[str, Message | None],
        found: dict[str, list[Attachment]],
        downloads: dict[str, Downloaded],
        request: ExportRequest,
    ) -> list[TextFile]:
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
                rendered.append(
                    RenderedMessage(message, body_text(bodies.get(message.id), request.body), lines)
                )
            file.text = render_file(
                title, rendered, body_kind=request.body, sections=request.combine == "all"
            )
            files.append(file)
        return files


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

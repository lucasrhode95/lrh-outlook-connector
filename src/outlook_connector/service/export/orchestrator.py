"""The single export path (requirements v4 §10), used by both the local UI and MCP.

selection (conversations + messages + a range) → copies merged → bodies (Graph $batch) →
attachments (policy) → grouping (per conversation / all / none) →
TXT or JSONL rendering → packaging.
"""

from __future__ import annotations

import asyncio
import hashlib
import shutil
import tempfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from outlook_connector.domain.errors import ConnectorError, InvalidRequest, NotFound, Throttled, Upstream
from outlook_connector.domain.models import (
    EXCLUSION_TEXT,
    EXPORT_MAX_MESSAGES,
    Attachment,
    ExportArtifact,
    ExportError,
    ExportRequest,
    Message,
    MessageSummary,
)
from outlook_connector.service.conversations import BODY_MISSING, Conversations, base_subject, oldest_first
from outlook_connector.service.export import attachments as policy
from outlook_connector.service.export.formatter import RenderedMessage, body_text, jsonl_record, render_file
from outlook_connector.service.export.packaging import TextFile, package
from outlook_connector.service.failures import (
    THROTTLING,
    error_block,
    error_from,
    error_summary,
    export_error,
)
from outlook_connector.service.mailbox import OUTGOING_FOLDERS, merge_copies
from outlook_connector.service.scope import validate_scope

RANGE_PAGE = 200
SHOWN_FAILURES = 3
UNREADABLE_HINT = (
    "They may have been deleted or moved in Outlook, or Microsoft is throttling requests. "
    "Refresh the list (or list/search again) and retry the export. Nothing was exported."
)

Downloaded = list[tuple[Attachment, Path | ExportError]]  # the file, or why it could not be downloaded


@dataclass
class Selection:
    summaries: list[MessageSummary]
    excluded: dict[str, int] = field(default_factory=dict)  # ExclusionReason -> messages left out
    known: dict[str, Message] = field(default_factory=dict)  # fetched with bodies while selecting


@dataclass
class Fetched:
    bodies: dict[str, Message]
    missing: dict[str, ExportError]  # message id -> why it has no body
    attachments: dict[str, list[Attachment]]
    attachment_failures: dict[str, ExportError]  # message id -> why its attachments could not be listed


class Exports:
    def __init__(self, conversations: Conversations) -> None:
        self.conversations = conversations
        self.mailbox = conversations.mailbox
        self.reader = conversations.mailbox.reader

    async def export(self, request: ExportRequest) -> ExportArtifact:
        """Export a selection to one local file (requirements v4 §10).

        Entry point: ``request`` is an ExportRequest (its fields validated by the model); this method
        checks that something is selected and, through ``_select``, the message limit. The steps after
        ``_select`` trust the selection and do not re-check it.
        """
        validate_scope(
            request.scope,
            sent_items=request.by_range,
            meeting_mail=request.by_range,
            deleted_items=request.by_range or bool(request.conversation_ids),
        )
        if not request.conversation_ids and not request.message_ids and not request.by_range:
            raise InvalidRequest(
                "Select at least one conversation or message, or a range "
                "(since/until/folder/scope.sent_items=false)."
            )
        selection = await self._select(request)
        summaries = selection.summaries
        if not summaries:
            raise InvalidRequest(
                f"Nothing to export: the selection holds no messages{_excluded_note(selection.excluded)}."
            )
        bodies, missing = await self.conversations.bodies(summaries, known=selection.known)
        found, attachment_failures = await self._attachments(
            summaries, inline=request.include_attachments, skip=set(missing)
        )
        fetched = Fetched(bodies, missing, found, attachment_failures)
        workdir = Path(tempfile.mkdtemp(prefix="outlook-export-"))
        try:
            downloads = (
                await self._download(summaries, bodies, found, request, workdir)
                if request.include_attachments
                else {}
            )
            failed_files = [p for items in downloads.values() for _, p in items if isinstance(p, ExportError)]
            summary = error_summary(list(missing.values()), failed_files, list(attachment_failures.values()))
            files = self._render(summaries, fetched, downloads, request, selection, summary)
            base = files[0].folder if len(files) == 1 else _range_name(summaries)
            result = package(files, base_name=base)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)
        counts = Counter(e.step for e in [*missing.values(), *failed_files, *attachment_failures.values()])
        return ExportArtifact(
            path=str(result.path),
            filename=result.filename,
            content_type=result.content_type,
            size=result.size,
            message_count=len(summaries),
            text_files=len(files),
            attachment_files=sum(len(f.attachments) for f in files),
            messages_excluded=selection.excluded,
            unavailable_message_ids=list(missing),
            export_errors=dict(counts),
            error_summary=summary,
        )

    # ---------------------------------------------------------------- selection

    async def _select(self, request: ExportRequest) -> Selection:
        """Conversations, the range and messages selected by id, copies merged, within ``request.limit``.

        Assumes (not re-checked here): ``request`` was validated by ``export``.
        """
        selected: dict[str, MessageSummary] = {}
        excluded: dict[str, int] = defaultdict(int)
        for conversation_id in dict.fromkeys(request.conversation_ids):
            items, left_out, truncated = await self.conversations.messages(
                conversation_id, scope=request.scope
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
        known = await self._fetch(explicit)
        selected.update({mid: MessageSummary.model_validate(m.model_dump()) for mid, m in known.items()})
        # Conversations/ranges already applied scope. Explicit ids are authoritative, like get_message.
        folders, _ = await self.mailbox.reach(m.folder_id for m in selected.values())
        outgoing = {f.id for f in folders.values() if f.well_known in OUTGOING_FOLDERS}
        merged = merge_copies(await self.mailbox.decorate(list(selected.values())), outgoing)
        _check_limit({m.id: m for m in merged}, request.limit)
        return Selection(sorted(merged, key=oldest_first), dict(excluded), known)

    async def _select_range(
        self, request: ExportRequest, selected: dict[str, MessageSummary], excluded: dict[str, int]
    ) -> None:
        """Every message in the window, page by page, with the listing's scope rules. Copies on
        different pages are all kept here, so the final merge sees them and names every folder in
        ``also_in``.

        Assumes (not re-checked here): ``request`` was validated by ``export``; ``list_messages``
        validates the window.
        """
        cursor: str | None = None
        while True:
            page = await self.mailbox.list_messages(
                folder=request.folder,
                since=request.since,
                until=request.until,
                limit=RANGE_PAGE,
                cursor=cursor,
                scope=request.scope,
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

    async def _fetch(self, message_ids: list[str]) -> dict[str, Message]:
        """Messages selected by id, read from the server in batches; their bodies are reused.

        Assumes (not re-checked here): ``message_ids`` are unique and not already selected (``_select``
        removes those).
        """
        if not message_ids:
            return {}
        fetched = await self.reader.get_messages(message_ids)
        gone = [mid for mid, m in fetched.messages.items() if m is None]
        if fetched.failed or gone:
            cases = [f"{mid} (failed: {f.describe()})" for mid, f in fetched.failed.items()]
            cases += [f"{mid} (not found: deleted, moved out of reach, or wrong id)" for mid in gone]
            shown = "; ".join(cases[:SHOWN_FAILURES]) + ("; ..." if len(cases) > SHOWN_FAILURES else "")
            text = (
                f"{len(cases)} of the {len(message_ids)} message(s) selected by id could not be read: "
                f"{shown}. {UNREADABLE_HINT}"
            )
            if not fetched.failed:
                raise NotFound(text)
            if any(f.status in THROTTLING for f in fetched.failed.values()):
                raise Throttled(text)
            raise Upstream(text)
        return {mid: m for mid, m in fetched.messages.items() if m is not None}

    async def _attachments(
        self,
        summaries: list[MessageSummary],
        *,
        inline: bool,
        skip: set[str],
    ) -> tuple[dict[str, list[Attachment]], dict[str, ExportError]]:
        """Attachments per message, and why listing failed for others. ``skip``: messages without
        a body (deleted on the server, or not fetched). Graph reports hasAttachments=false when a
        message has only inline attachments, so when files are exported (inline images included)
        every message is asked, in batches.

        Assumes (not re-checked here): ``summaries`` is the final selection from ``_select``, and ``skip``
        the messages without a body from ``Conversations.bodies``.
        """
        wanted = [m.id for m in summaries if m.id not in skip and (inline or m.has_attachments)]
        if not wanted:
            return {}, {}
        found, failed = await self.reader.list_attachments_many(wanted)
        return found, {mid: export_error("listing attachments", f) for mid, f in failed.items()}

    # ---------------------------------------------------------------- attachments

    async def _download(
        self,
        summaries: list[MessageSummary],
        bodies: dict[str, Message],
        found: dict[str, list[Attachment]],
        request: ExportRequest,
        workdir: Path,
    ) -> dict[str, Downloaded]:
        """Download the attachments the policy wants, each to its own file in ``workdir``; a failed download
        becomes an ExportError.

        Assumes (not re-checked here): ``summaries`` is the final selection and ``found`` the attachment
        listing from ``_attachments`` for it.
        """
        inline_ids = {
            mid: [a.id for a in items if a.is_inline and a.kind == "file"] for mid, items in found.items()
        }
        needs_html = [m.id for m in summaries if inline_ids.get(m.id)]
        html = (await self.reader.get_messages(needs_html, body_format="html")).messages if needs_html else {}

        jobs: list[tuple[str, Attachment, Path]] = []
        for index, summary in enumerate(summaries):
            page = html.get(summary.id) or bodies.get(summary.id)
            page_html = None
            if page is not None:
                page_html = page.unique_body_html if request.body == "unique" else page.body_html
            for position, attachment in enumerate(found.get(summary.id, [])):
                if policy.wanted(attachment, body_html=page_html):
                    jobs.append((summary.id, attachment, workdir / f"{index}-{position}"))

        async def fetch(message_id: str, attachment: Attachment, target: Path) -> Path | ExportError:
            try:
                await self.reader.download_attachment(message_id, attachment.id, target)
            except ConnectorError as exc:
                return error_from("downloading an attachment", exc)
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
        summary: str | None,
    ) -> list[TextFile]:
        """The text files: one per group (per conversation, all, none), TXT or JSONL.

        Assumes (not re-checked here): ``fetched`` covers every summary (each has a body or an entry in
        ``missing``), and ``downloads`` comes from ``_download`` for the same selection.
        """
        notes = [
            f"Left out: {count} message(s) {EXCLUSION_TEXT[key]}."
            for key, count in selection.excluded.items()
        ]
        if summary:
            notes.append(summary)
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
        """One message's text, attachment lines (TXT) and attachment records (JSONL).

        Assumes (not re-checked here): the message has a body in ``fetched.bodies`` or an error in
        ``fetched.missing`` (``Conversations.bodies`` guarantees one of the two).
        """
        lines: list[str] = []
        blocks: list[str] = []
        records: list[dict[str, Any]] = []
        listing_error = fetched.attachment_failures.get(message.id)
        if listing_error:
            blocks.append(error_block("The attachments of this message could not be listed.", listing_error))
        listed = fetched.attachments.get(message.id, [])
        if request.include_attachments:
            for attachment, path in downloads.get(message.id, []):
                if isinstance(path, ExportError):
                    name = attachment.name or "(unnamed)"
                    blocks.append(error_block(f"The attachment {name} could not be downloaded.", path))
                    records.append(_attachment_record(attachment, error=path))
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
        error = fetched.missing.get(message.id)
        text = (
            error_block(BODY_MISSING, error) if error else body_text(fetched.bodies[message.id], request.body)
        )
        return RenderedMessage(
            message,
            text,
            lines,
            export_error=error,
            error_blocks=blocks,
            attachments=records,
            attachments_export_error=listing_error,
        )


def _attachment_record(
    a: Attachment, *, file: str | None = None, error: ExportError | None = None
) -> dict[str, Any]:
    record: dict[str, Any] = {"name": a.name, "size": a.size, "kind": a.kind, "inline": a.is_inline}
    if file:
        record["file"] = file
    if error:
        record["export_error"] = error.model_dump(mode="json")
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

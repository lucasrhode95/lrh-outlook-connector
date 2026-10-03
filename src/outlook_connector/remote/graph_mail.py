"""MailReader over Microsoft Graph (research §3)."""

from __future__ import annotations

import asyncio
import functools
from collections.abc import Callable, Coroutine, Mapping
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, TypeVar, cast

from outlook_connector.domain.errors import Failure
from outlook_connector.domain.models import Attachment, Folder, Message, MessageSummary
from outlook_connector.remote import graph_mapping as mapping
from outlook_connector.remote.graph import (
    PREFER_TEXT_BODY,
    Graph,
    failure_of,
    raise_for_failures,
    relative,
    sub_failure,
)
from outlook_connector.remote.ports import BodyFormat, FetchedMessages, FetchedSummaries
from outlook_connector.remote.transport import operation

WELL_KNOWN = (
    "inbox", "sentitems", "drafts", "outbox", "deleteditems", "junkemail", "archive",
    "recoverableitemsdeletions", "searchfolders", "conversationhistory",
    "syncissues", "conflicts", "localfailures", "serverfailures",
)  # fmt: skip
MAX_CONVERSATION = 1000


def _iso(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _window(since: datetime | None, until: datetime | None) -> str | None:
    conditions = []
    if since:
        conditions.append(f"receivedDateTime ge {_iso(since)}")
    if until:
        conditions.append(f"receivedDateTime le {_iso(until)}")
    return " and ".join(conditions) or None


def _odata_string(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


F = TypeVar("F", bound=Callable[..., Coroutine[Any, Any, Any]])


def _named(name: str) -> Callable[[F], F]:
    """Errors raised inside the decorated call say what was being done."""

    def wrap(func: F) -> F:
        @functools.wraps(func)
        async def run(*args: Any, **kwargs: Any) -> Any:
            with operation(name):
                return await func(*args, **kwargs)

        return cast(F, run)

    return wrap


class GraphMailReader:
    def __init__(self, graph: Graph) -> None:
        self._graph = graph

    @_named("listing folders")
    async def list_folders(self) -> list[Folder]:
        params = {"$top": 100, "includeHiddenFolders": "true", "$select": mapping.FOLDER_FIELDS}

        async def level(paths: list[str]) -> list[dict]:
            pages = await asyncio.gather(*(self._graph.collect(p, params, max_items=1000) for p in paths))
            return [item for items, _ in pages for item in items]

        # Aliases and the folder tree are independent; each tree level is fetched in parallel.
        aliases_task = asyncio.create_task(self._well_known_ids())
        raw: list[dict] = []
        paths = ["/me/mailFolders"]
        while paths:
            items = await level(paths)
            raw.extend(items)
            paths = [f"/me/mailFolders/{i['id']}/childFolders" for i in items if i.get("childFolderCount")]
        aliases = await aliases_task
        return [mapping.folder(item, aliases.get(item["id"])) for item in raw]

    async def _well_known_ids(self) -> dict[str, str]:
        responses = await self._graph.batch(
            {alias: relative(f"/me/mailFolders/{alias}", {"$select": "id"}) for alias in WELL_KNOWN}
        )
        return {r.body["id"]: alias for alias, r in responses.items() if r.status == 200 and "id" in r.body}

    @_named("listing messages")
    async def list_messages(
        self,
        *,
        folder_id: str | None,
        since: datetime | None,
        until: datetime | None,
        page_size: int,
        page: str | None,
    ) -> tuple[list[MessageSummary], str | None]:
        if page:
            items, link = await self._graph.page(page)
        else:
            path = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
            params = {
                "$select": mapping.SUMMARY_FIELDS,
                "$top": page_size,
                "$orderby": "receivedDateTime desc",
                "$filter": _window(since, until),
            }
            items, link = await self._graph.page(path, params)
        return [mapping.summary(i) for i in items], link

    @_named("searching the mailbox")
    async def search(
        self, *, query: str, folder_id: str | None, page_size: int, page: str | None
    ) -> tuple[list[MessageSummary], str | None]:
        if page:
            items, link = await self._graph.page(page)
        else:
            path = f"/me/mailFolders/{folder_id}/messages" if folder_id else "/me/messages"
            params = {
                "$search": '"' + query.replace('"', '\\"') + '"',
                "$top": page_size,
                "$select": mapping.SUMMARY_FIELDS,
            }
            items, link = await self._graph.page(path, params)
        return await self._immutable([mapping.summary(i) for i in items]), link

    async def _immutable(self, found: list[MessageSummary]) -> list[MessageSummary]:
        """$search ignores ``Prefer: IdType="ImmutableId"`` (live 2026-10-03) and returns regular ids,
        which change when a message moves. Read each hit's id back (a GET honours the header), so
        search gives the same ids as every other call. A hit gone meanwhile is dropped; one whose
        lookup failed keeps its search id."""
        requests = {str(i): relative(f"/me/messages/{m.id}", {"$select": "id"}) for i, m in enumerate(found)}
        responses = await self._graph.batch(requests) if requests else {}
        out = []
        for index, message in enumerate(found):
            response = responses[str(index)]
            if response.status == 404:
                continue
            if response.ok and isinstance(response.body.get("id"), str):
                message.id = response.body["id"]
            out.append(message)
        return out

    @_named("listing a conversation")
    async def conversation(self, conversation_id: str) -> tuple[list[MessageSummary], bool]:
        # $orderby cannot be combined with this filter (InefficientFilter, research §3.4): sort locally.
        params = {
            "$filter": f"conversationId eq {_odata_string(conversation_id)}",
            "$select": mapping.SUMMARY_FIELDS,
            "$top": 100,
        }
        items, truncated = await self._graph.collect("/me/messages", params, max_items=MAX_CONVERSATION)
        return [mapping.summary(i) for i in items], truncated

    @_named("counting conversation messages")
    async def conversation_folders(
        self, conversation_ids: list[str]
    ) -> dict[str, tuple[list[tuple[str | None, str | None]], bool]]:
        requests = {
            str(index): relative(
                "/me/messages",
                {
                    "$filter": f"conversationId eq {_odata_string(cid)}",
                    "$select": "id,parentFolderId,internetMessageId",
                    "$top": MAX_CONVERSATION,
                },
            )
            for index, cid in enumerate(conversation_ids)
        }
        out: dict[str, tuple[list[tuple[str | None, str | None]], bool]] = {}
        responses = await self._graph.batch(requests)
        raise_for_failures(responses)
        for key, response in responses.items():
            items = [i for i in response.body.get("value", []) if isinstance(i, dict)]
            out[conversation_ids[int(key)]] = (
                [(i.get("parentFolderId"), i.get("internetMessageId")) for i in items],
                "@odata.nextLink" in response.body,
            )
        return out

    @_named("counting messages")
    async def count_messages(
        self, *, folder_ids: list[str], since: datetime | None, until: datetime | None
    ) -> dict[str, int] | None:
        """Messages in the window per folder (that folder only, not its subfolders), in $batch."""
        params = {"$count": "true", "$top": 1, "$select": "id", "$filter": _window(since, until)}
        requests = {
            str(index): relative(f"/me/mailFolders/{fid}/messages", params)
            for index, fid in enumerate(folder_ids)
        }
        responses = await self._graph.batch(requests, headers={"ConsistencyLevel": "eventual"})
        counts: dict[str, int] = {}
        for key, response in responses.items():
            count = response.body.get("@odata.count")
            if not response.ok or not isinstance(count, int):
                return None
            counts[folder_ids[int(key)]] = count
        return counts

    @_named("reading a message")
    async def get_message(self, message_id: str, *, body_format: BodyFormat = "text") -> Message:
        prefer = (PREFER_TEXT_BODY,) if body_format == "text" else ()
        data = await self._graph.get(
            f"/me/messages/{message_id}", {"$select": mapping.MESSAGE_FIELDS}, prefer=prefer
        )
        result = mapping.message(data, html=body_format == "html")
        if result.has_attachments:
            result.attachments = await self.list_attachments(message_id)
        return result

    @_named("fetching message bodies")
    async def get_messages(
        self, message_ids: list[str], *, body_format: BodyFormat = "text"
    ) -> FetchedMessages:
        prefer = (PREFER_TEXT_BODY,) if body_format == "text" else ()
        requests = {
            mid: relative(f"/me/messages/{mid}", {"$select": mapping.MESSAGE_FIELDS}) for mid in message_ids
        }
        responses = await self._graph.batch(requests, prefer=prefer)
        out = FetchedMessages()
        for mid, response in responses.items():
            if response.status == 404:
                out.messages[mid] = None
            elif response.ok:
                out.messages[mid] = mapping.message(response.body, html=body_format == "html")
            else:
                out.failed[mid] = failure_of(response)
        return out

    @_named("reading message state")
    async def get_summaries(self, message_ids: list[str]) -> FetchedSummaries:
        requests = {
            mid: relative(f"/me/messages/{mid}", {"$select": mapping.SUMMARY_FIELDS}) for mid in message_ids
        }
        out = FetchedSummaries()
        for mid, response in (await self._graph.batch(requests)).items() if requests else []:
            if response.status == 404:
                out.summaries[mid] = None
            elif response.ok:
                out.summaries[mid] = mapping.summary(response.body)
            else:
                out.failed[mid] = str(sub_failure(response))
        return out

    @_named("listing attachments")
    async def list_attachments(self, message_id: str) -> list[Attachment]:
        items, _ = await self._graph.collect(
            f"/me/messages/{message_id}/attachments", {"$select": mapping.ATTACHMENT_FIELDS}, max_items=500
        )
        return [mapping.attachment(i, message_id) for i in items]

    @_named("listing attachments")
    async def list_attachments_many(
        self, message_ids: list[str]
    ) -> tuple[dict[str, list[Attachment]], dict[str, Failure]]:
        requests = {
            mid: relative(f"/me/messages/{mid}/attachments", {"$select": mapping.ATTACHMENT_FIELDS})
            for mid in message_ids
        }
        responses = await self._graph.batch(requests)
        out: dict[str, list[Attachment]] = {}
        failed: dict[str, Failure] = {}
        for mid, response in responses.items():
            if not response.ok:
                failed[mid] = failure_of(response)
            elif "@odata.nextLink" in response.body:  # rare: more than one page of attachments
                out[mid] = await self.list_attachments(mid)
            else:
                values = response.body.get("value", [])
                out[mid] = [mapping.attachment(i, mid) for i in values if isinstance(i, dict)]
        return out, failed

    @_named("reading inline attachment ids")
    async def attachment_content_ids(
        self, attachments: Mapping[str, list[str]]
    ) -> dict[str, dict[str, str | None]]:
        """Content ids of attachments, {message id: [attachment ids]}, looked up together: one
        $batch item per attachment, 20 per batch across messages."""
        pairs = [(mid, aid) for mid, aids in attachments.items() for aid in aids]
        requests = {
            str(index): relative(
                f"/me/messages/{mid}/attachments/{aid}",
                {"$select": "microsoft.graph.fileAttachment/contentId"},
            )
            for index, (mid, aid) in enumerate(pairs)
        }
        responses = await self._graph.batch(requests) if requests else {}
        out: dict[str, dict[str, str | None]] = {mid: {} for mid in attachments}
        for index, (mid, aid) in enumerate(pairs):
            response = responses[str(index)]
            out[mid][aid] = response.body.get("contentId") if response.status == 200 else None
        return out

    @_named("downloading an attachment")
    async def download_attachment(self, message_id: str, attachment_id: str, dest: Path) -> int:
        _, size = await self._graph.download(
            f"/me/messages/{message_id}/attachments/{attachment_id}/$value", dest
        )
        return size

    @_named("downloading a message")
    async def download_mime(self, message_id: str, dest: Path) -> int:
        _, size = await self._graph.download(f"/me/messages/{message_id}/$value", dest)
        return size

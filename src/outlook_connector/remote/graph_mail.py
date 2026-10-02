"""MailReader over Microsoft Graph (research §3)."""

from __future__ import annotations

import asyncio
from datetime import UTC, datetime
from pathlib import Path

from outlook_connector.domain.models import Attachment, Folder, Message, MessageSummary
from outlook_connector.remote import graph_mapping as mapping
from outlook_connector.remote.graph import PREFER_TEXT_BODY, Graph, raise_for_sub_status, relative
from outlook_connector.remote.ports import BodyFormat

WELL_KNOWN = (
    "inbox", "sentitems", "drafts", "outbox", "deleteditems", "junkemail", "archive",
    "recoverableitemsdeletions", "searchfolders", "conversationhistory", "syncissues",
)  # fmt: skip
MAX_CONVERSATION = 1000


def _iso(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _odata_string(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


class GraphMailReader:
    def __init__(self, graph: Graph) -> None:
        self._graph = graph

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
        return {
            body["id"]: alias for alias, (status, body) in responses.items() if status == 200 and "id" in body
        }

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
            conditions = []
            if since:
                conditions.append(f"receivedDateTime ge {_iso(since)}")
            if until:
                conditions.append(f"receivedDateTime le {_iso(until)}")
            params = {
                "$select": mapping.SUMMARY_FIELDS,
                "$top": page_size,
                "$orderby": "receivedDateTime desc",
                "$filter": " and ".join(conditions) or None,
            }
            items, link = await self._graph.page(path, params)
        return [mapping.summary(i) for i in items], link

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
        return [mapping.summary(i) for i in items], link

    async def search_total(self, query: str) -> int | None:
        body = {
            "requests": [{"entityTypes": ["message"], "query": {"queryString": query}, "from": 0, "size": 1}]
        }
        data = await self._graph.post("/search/query", body)
        for value in data.get("value", []):
            for container in value.get("hitsContainers", []):
                if isinstance(container.get("total"), int):
                    return container["total"]
        return None

    async def conversation(self, conversation_id: str) -> list[MessageSummary]:
        # $orderby cannot be combined with this filter (InefficientFilter, research §3.4): sort locally.
        params = {
            "$filter": f"conversationId eq {_odata_string(conversation_id)}",
            "$select": mapping.SUMMARY_FIELDS,
            "$top": 100,
        }
        items, _ = await self._graph.collect("/me/messages", params, max_items=MAX_CONVERSATION)
        return [mapping.summary(i) for i in items]

    async def get_message(self, message_id: str, *, body_format: BodyFormat = "text") -> Message:
        prefer = (PREFER_TEXT_BODY,) if body_format == "text" else ()
        data = await self._graph.get(
            f"/me/messages/{message_id}", {"$select": mapping.MESSAGE_FIELDS}, prefer=prefer
        )
        result = mapping.message(data, html=body_format == "html")
        if result.has_attachments:
            result.attachments = await self.list_attachments(message_id)
        return result

    async def get_messages(
        self, message_ids: list[str], *, body_format: BodyFormat = "text"
    ) -> dict[str, Message | None]:
        prefer = (PREFER_TEXT_BODY,) if body_format == "text" else ()
        requests = {
            mid: relative(f"/me/messages/{mid}", {"$select": mapping.MESSAGE_FIELDS}) for mid in message_ids
        }
        responses = await self._graph.batch(requests, prefer=prefer)
        out: dict[str, Message | None] = {}
        for mid, (status, body) in responses.items():
            if status == 404:
                out[mid] = None
                continue
            raise_for_sub_status(status, body)
            out[mid] = mapping.message(body, html=body_format == "html")
        return out

    async def locate(self, message_ids: list[str]) -> dict[str, str | None]:
        requests = {
            mid: relative(f"/me/messages/{mid}", {"$select": "id,parentFolderId"}) for mid in message_ids
        }
        out: dict[str, str | None] = {}
        for mid, (status, body) in (await self._graph.batch(requests)).items():
            if status == 404:
                out[mid] = None
            else:
                raise_for_sub_status(status, body)
                out[mid] = body.get("parentFolderId")
        return out

    async def list_attachments(self, message_id: str) -> list[Attachment]:
        items, _ = await self._graph.collect(
            f"/me/messages/{message_id}/attachments", {"$select": mapping.ATTACHMENT_FIELDS}, max_items=500
        )
        return [mapping.attachment(i, message_id) for i in items]

    async def attachment_content_ids(
        self, message_id: str, attachment_ids: list[str]
    ) -> dict[str, str | None]:
        requests = {
            aid: relative(
                f"/me/messages/{message_id}/attachments/{aid}",
                {"$select": "microsoft.graph.fileAttachment/contentId"},
            )
            for aid in attachment_ids
        }
        out: dict[str, str | None] = {}
        for aid, (status, body) in (await self._graph.batch(requests)).items():
            out[aid] = body.get("contentId") if status == 200 else None
        return out

    async def download_attachment(self, message_id: str, attachment_id: str, dest: Path) -> int:
        _, size = await self._graph.download(
            f"/me/messages/{message_id}/attachments/{attachment_id}/$value", dest
        )
        return size

    async def download_mime(self, message_id: str, dest: Path) -> int:
        _, size = await self._graph.download(f"/me/messages/{message_id}/$value", dest)
        return size

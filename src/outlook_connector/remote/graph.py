"""Generic Microsoft Graph plumbing: URLs, paging, JSON batching, downloads.

Every request asks for immutable ids. Continuation links are opaque and must stay on the Graph host.
"""

from __future__ import annotations

import asyncio
from collections.abc import Mapping
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlencode

from outlook_connector.domain.errors import InvalidRequest, NotFound, Throttled, Upstream
from outlook_connector.remote.transport import Transport

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
PREFER_IMMUTABLE = 'IdType="ImmutableId"'
PREFER_TEXT_BODY = 'outlook.body-content-type="text"'
BATCH_LIMIT = 20  # Graph JSON batching maximum per request
MAX_DOWNLOAD_BYTES = 150 * 1024 * 1024


def relative(path: str, params: Mapping[str, Any] | None = None) -> str:
    """A version-relative URL such as ``/me/messages?$select=id`` (used directly in $batch)."""
    query = urlencode({k: v for k, v in (params or {}).items() if v is not None}, quote_via=quote)
    return f"{path}?{query}" if query else path


def _prefer(*extra: str) -> dict[str, str]:
    return {"Prefer": ", ".join((PREFER_IMMUTABLE, *extra))}


class Graph:
    def __init__(self, transport: Transport, *, profile: str = "read") -> None:
        self._transport = transport
        self._profile = profile

    def _absolute(self, path_or_url: str) -> str:
        if path_or_url.startswith("https://"):
            if not path_or_url.startswith(GRAPH_ROOT + "/"):
                raise InvalidRequest("Continuation link does not point at Microsoft Graph.")
            return path_or_url
        return GRAPH_ROOT + path_or_url

    async def get(
        self, path: str, params: Mapping[str, Any] | None = None, *, prefer: tuple[str, ...] = ()
    ) -> dict[str, Any]:
        url = self._absolute(relative(path, params) if not path.startswith("https://") else path)
        data = await self._transport.json("GET", url, profile=self._profile, headers=_prefer(*prefer))
        return data if isinstance(data, dict) else {}

    async def post(self, path: str, body: Any) -> dict[str, Any]:
        """POST for read-style APIs (search). Idempotent, so retried like a GET."""
        data = await self._transport.json(
            "POST", self._absolute(path), profile=self._profile, headers=_prefer(), json_body=body, retry=True
        )
        return data if isinstance(data, dict) else {}

    async def page(
        self, path_or_link: str, params: Mapping[str, Any] | None = None, *, prefer: tuple[str, ...] = ()
    ) -> tuple[list[dict[str, Any]], str | None]:
        """One page of a collection and the opaque link to the next page (if any)."""
        data = await self.get(
            path_or_link, None if path_or_link.startswith("https://") else params, prefer=prefer
        )
        items = [x for x in data.get("value", []) if isinstance(x, dict)]
        link = data.get("@odata.nextLink")
        return items, link if isinstance(link, str) else None

    async def collect(
        self,
        path: str,
        params: Mapping[str, Any] | None = None,
        *,
        max_items: int,
        prefer: tuple[str, ...] = (),
    ) -> tuple[list[dict[str, Any]], bool]:
        """Follow pages up to ``max_items``. Returns (items, truncated)."""
        items, link = await self.page(path, params, prefer=prefer)
        while link and len(items) < max_items:
            more, link = await self.page(link, prefer=prefer)
            items.extend(more)
        return items[:max_items], bool(link) or len(items) > max_items

    async def batch(
        self, requests: Mapping[str, str], *, prefer: tuple[str, ...] = ()
    ) -> dict[str, tuple[int, dict[str, Any]]]:
        """GET many relative URLs via $batch. Returns {request id: (status, body)}.

        Sub-requests throttled with 429 are retried once (individually) after the advised delay.
        """
        keys = list(requests)
        chunks = [keys[i : i + BATCH_LIMIT] for i in range(0, len(keys), BATCH_LIMIT)]
        results: dict[str, tuple[int, dict[str, Any]]] = {}
        for part in await asyncio.gather(
            *(self._batch_once({k: requests[k] for k in c}, prefer) for c in chunks)
        ):
            results.update(part)
        throttled = [k for k, (status, _) in results.items() if status == 429]
        if throttled:
            await asyncio.sleep(5)
            retried = await self._batch_once({k: requests[k] for k in throttled}, prefer)
            results.update(retried)
            if any(status == 429 for status, _ in retried.values()):
                raise Throttled("Microsoft Graph is throttling batch requests. Retry later.")
        return results

    async def _batch_once(
        self, requests: Mapping[str, str], prefer: tuple[str, ...]
    ) -> dict[str, tuple[int, dict[str, Any]]]:
        body = {
            "requests": [
                {"id": key, "method": "GET", "url": url, "headers": _prefer(*prefer)}
                for key, url in requests.items()
            ]
        }
        data = await self.post("/$batch", body)
        out: dict[str, tuple[int, dict[str, Any]]] = {}
        for response in data.get("responses", []):
            body_part = response.get("body")
            out[str(response.get("id"))] = (
                int(response.get("status", 0)),
                body_part if isinstance(body_part, dict) else {},
            )
        missing = set(requests) - set(out)
        if missing:
            raise Upstream(f"Graph batch response is missing {len(missing)} sub-responses.")
        return out

    async def download(
        self, path: str, dest: Path, *, max_bytes: int = MAX_DOWNLOAD_BYTES
    ) -> tuple[str | None, int]:
        dest.parent.mkdir(parents=True, exist_ok=True)
        with dest.open("wb") as handle:
            return await self._transport.download(
                self._absolute(path), handle, profile=self._profile, headers=_prefer(), max_bytes=max_bytes
            )


def raise_for_sub_status(status: int, body: dict[str, Any]) -> None:
    """Map a failed $batch sub-response to a domain error (404 is usually handled by the caller)."""
    if 200 <= status < 300:
        return
    error = body.get("error") if isinstance(body.get("error"), dict) else {}
    code = error.get("code", "no error code")
    if status == 404:
        raise NotFound(f"Not found ({code}).")
    if status == 429:
        raise Throttled(f"Microsoft Graph is throttling requests ({code}).")
    if status == 400:
        raise InvalidRequest(f"Microsoft Graph rejected the request ({code}).")
    raise Upstream(f"Microsoft Graph returned HTTP {status} ({code}).")

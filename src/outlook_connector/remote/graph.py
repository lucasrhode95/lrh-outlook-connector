"""Generic Microsoft Graph plumbing: URLs, paging, JSON batching, downloads.

Every request asks for immutable ids. Continuation links are opaque and must stay on the Graph host.
"""

from __future__ import annotations

import asyncio
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlencode

from outlook_connector.domain.errors import InvalidRequest, Upstream
from outlook_connector.remote.transport import Transport, describe_failure, request_id, service_error

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
PREFER_IMMUTABLE = 'IdType="ImmutableId"'
PREFER_TEXT_BODY = 'outlook.body-content-type="text"'
BATCH_LIMIT = 20  # Graph JSON batching maximum per request
BATCH_CONCURRENCY = 2  # batches in flight; each sub-request counts against the mailbox's 4 concurrent
BATCH_RETRIES = 4  # rounds of re-sending throttled sub-requests
DEFAULT_RETRY_WAIT = 5.0
MAX_RETRY_WAIT = 30.0
MAX_DOWNLOAD_BYTES = 150 * 1024 * 1024


def relative(path: str, params: Mapping[str, Any] | None = None) -> str:
    """A version-relative URL such as ``/me/messages?$select=id`` (used directly in $batch)."""
    query = urlencode({k: v for k, v in (params or {}).items() if v is not None}, quote_via=quote)
    return f"{path}?{query}" if query else path


@dataclass
class SubResponse:
    status: int
    body: dict[str, Any]
    headers: dict[str, str] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return 200 <= self.status < 300


def _prefer(*extra: str) -> dict[str, str]:
    return {"Prefer": ", ".join((PREFER_IMMUTABLE, *extra))}


class Graph:
    def __init__(self, transport: Transport, *, profile: str = "read") -> None:
        self._transport = transport
        self._profile = profile
        self._batch_gate = asyncio.Semaphore(BATCH_CONCURRENCY)

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
    ) -> dict[str, SubResponse]:
        """GET many relative URLs via $batch. Returns {caller key: sub-response}, one per request.

        - Batch request ids are numbers assigned here, never the caller's keys: Graph compares them
          case-insensitively, and immutable ids can differ only by case.
        - At most BATCH_LIMIT requests per batch and BATCH_CONCURRENCY batches in flight across all
          concurrent calls, because
          Exchange Online throttles more than a few concurrent requests per mailbox (sub-requests count).
        - Throttled sub-requests (429) are re-sent in new batches of at most BATCH_LIMIT after the
          advised delay, up to BATCH_RETRIES rounds. What is still throttled is returned as 429.
        """
        results: dict[str, SubResponse] = {}
        pending = list(requests)

        async def send(keys: list[str]) -> dict[str, SubResponse]:
            async with self._batch_gate:  # shared by every batch() call of this process
                return await self._batch_once({k: requests[k] for k in keys}, prefer)

        for attempt in range(BATCH_RETRIES + 1):
            chunks = [pending[i : i + BATCH_LIMIT] for i in range(0, len(pending), BATCH_LIMIT)]
            for part in await asyncio.gather(*(send(c) for c in chunks)):
                results.update(part)
            pending = [k for k in pending if results[k].status == 429]
            if not pending or attempt == BATCH_RETRIES:
                break
            await self._transport.sleep(max(_retry_after(results[k].headers) for k in pending))
        return results

    async def _batch_once(
        self, requests: Mapping[str, str], prefer: tuple[str, ...]
    ) -> dict[str, SubResponse]:
        keys = list(requests)
        body = {
            "requests": [
                {"id": str(index), "method": "GET", "url": requests[key], "headers": _prefer(*prefer)}
                for index, key in enumerate(keys)
            ]
        }
        data = await self.post("/$batch", body)
        out: dict[str, SubResponse] = {}
        for response in data.get("responses", []):
            try:
                key = keys[int(response.get("id"))]
            except (TypeError, ValueError, IndexError):
                raise Upstream("Graph batch response has an unknown request id.") from None
            body_part = response.get("body")
            headers = response.get("headers")
            out[key] = SubResponse(
                int(response.get("status", 0)),
                body_part if isinstance(body_part, dict) else {},
                {str(k): str(v) for k, v in headers.items()} if isinstance(headers, dict) else {},
            )
        missing = set(keys) - set(out)
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


def _retry_after(headers: Mapping[str, str]) -> float:
    value = next((v for k, v in headers.items() if k.lower() == "retry-after"), "")
    try:
        return min(max(float(value), 1.0), MAX_RETRY_WAIT)
    except ValueError:
        return DEFAULT_RETRY_WAIT


def sub_failure(response: SubResponse, *, detail: str | None = None) -> Exception:
    """The domain error for a failed $batch sub-response."""
    code, message = service_error(response.body)
    return describe_failure(
        status=response.status,
        code=code,
        message=message,
        request=request_id(response.headers),
        detail=detail,
    )


def raise_for_failures(responses: Mapping[str, SubResponse], *, allow: tuple[int, ...] = ()) -> None:
    """Raise for the first failed sub-response (statuses in ``allow`` are fine), with a failure count."""
    failed = [r for r in responses.values() if not r.ok and r.status not in allow]
    if failed:
        raise sub_failure(failed[0], detail=f"{len(failed)} of {len(responses)} batch item(s) failed.")

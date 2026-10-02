"""The one HTTP client for Microsoft calls: auth header, host allowlist, retries, error mapping.

- One ``httpx.AsyncClient`` per process (keep-alive within a call).
- No redirects, and only allowlisted hosts.
- Retries only when the caller marks the request idempotent (GETs by default), honoring
  429 / Retry-After.
- Concurrency limiter to keep parallel fetches polite.
- Logs metadata only: no URLs with query text, no bodies, no tokens.
"""

from __future__ import annotations

import asyncio
import json
import logging
import time
from collections.abc import Awaitable, Callable
from typing import Any, BinaryIO, Protocol
from urllib.parse import urlsplit

import httpx

from outlook_connector.domain.errors import InvalidRequest, NotFound, Throttled, Upstream

log = logging.getLogger(__name__)

ALLOWED_HOSTS = frozenset({"graph.microsoft.com", "outlook.cloud.microsoft", "outlook.office.com"})
RETRY_STATUSES = frozenset({429, 502, 503, 504})
MAX_JSON_BYTES = 32 * 1024 * 1024


class _Token(Protocol):
    value: str


class TokenSource(Protocol):
    def get_token(self, profile: str) -> _Token: ...


class Transport:
    def __init__(
        self,
        tokens: TokenSource,
        *,
        client: httpx.AsyncClient | None = None,
        max_concurrency: int = 6,
        max_attempts: int = 4,
        sleep: Callable[[float], Awaitable[None]] = asyncio.sleep,
    ) -> None:
        self._tokens = tokens
        self._client = client or httpx.AsyncClient(
            timeout=httpx.Timeout(60.0, connect=15.0), follow_redirects=False
        )
        self._limit = asyncio.Semaphore(max_concurrency)
        self._max_attempts = max_attempts
        self._sleep = sleep

    async def aclose(self) -> None:
        await self._client.aclose()

    async def request(
        self,
        method: str,
        url: str,
        *,
        profile: str,
        headers: dict[str, str] | None = None,
        json_body: Any = None,
        retry: bool | None = None,
    ) -> httpx.Response:
        """Send a request and return a successful response, or raise a domain error."""
        _check_host(url)
        retry = (method.upper() == "GET") if retry is None else retry
        attempts = self._max_attempts if retry else 1
        for attempt in range(1, attempts + 1):
            request_headers = {
                "Authorization": f"Bearer {self._tokens.get_token(profile).value}",
                **(headers or {}),
            }
            started = time.monotonic()
            try:
                async with self._limit:
                    response = await self._client.request(
                        method, url, headers=request_headers, json=json_body
                    )
            except (httpx.ConnectError, httpx.ConnectTimeout) as exc:
                if attempt < attempts:
                    await self._sleep(min(2**attempt, 20))
                    continue
                raise Upstream(f"Could not reach {urlsplit(url).hostname} ({type(exc).__name__}).") from None
            except httpx.HTTPError as exc:  # sent, but no complete response
                if attempt < attempts:
                    await self._sleep(min(2**attempt, 20))
                    continue
                raise Upstream(
                    f"No complete response from {urlsplit(url).hostname} ({type(exc).__name__})."
                ) from None
            log.debug(
                "http %s %s -> %s in %dms",
                method,
                _path(url),
                response.status_code,
                (time.monotonic() - started) * 1000,
            )
            if response.status_code in RETRY_STATUSES and attempt < attempts:
                await self._sleep(_retry_after(response, attempt))
                continue
            if response.is_success:
                return response
            raise _error_for(response)
        raise AssertionError("unreachable")

    async def json(self, method: str, url: str, **kwargs: Any) -> Any:
        response = await self.request(method, url, **kwargs)
        if len(response.content) > MAX_JSON_BYTES:
            raise Upstream("Response too large.")
        if not response.content:
            return None
        try:
            return response.json()
        except json.JSONDecodeError:
            raise Upstream("Expected a JSON response.") from None

    async def download(
        self, url: str, dest: BinaryIO, *, profile: str, headers: dict[str, str] | None = None, max_bytes: int
    ) -> tuple[str | None, int]:
        """Stream a GET response body into ``dest``. Returns (content type, size)."""
        _check_host(url)
        request_headers = {
            "Authorization": f"Bearer {self._tokens.get_token(profile).value}",
            **(headers or {}),
        }
        async with self._limit, self._client.stream("GET", url, headers=request_headers) as response:
            if not response.is_success:
                await response.aread()
                raise _error_for(response)
            size = 0
            async for chunk in response.aiter_bytes():
                size += len(chunk)
                if size > max_bytes:
                    raise InvalidRequest(f"Download exceeds the {max_bytes // (1024 * 1024)} MiB limit.")
                dest.write(chunk)
            return response.headers.get("content-type"), size


def _check_host(url: str) -> None:
    parts = urlsplit(url)
    if parts.scheme != "https" or parts.hostname not in ALLOWED_HOSTS:
        raise InvalidRequest(f"Refusing to call a non-allowlisted URL host: {parts.hostname}")


def _path(url: str) -> str:
    return urlsplit(url).path


def _retry_after(response: httpx.Response, attempt: int) -> float:
    value = response.headers.get("retry-after", "")
    try:
        return min(float(value), 30.0)
    except ValueError:
        return float(min(2**attempt, 20))


def error_code(response: httpx.Response) -> str | None:
    owa = response.headers.get("x-owa-error")
    if owa:
        return owa[:120]
    try:
        data = response.json()
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None
    error = data.get("error") if isinstance(data, dict) else None
    if isinstance(error, dict) and isinstance(error.get("code"), str):
        return error["code"]
    return None


def _error_for(response: httpx.Response) -> Exception:
    code = error_code(response) or "no error code"
    status = response.status_code
    host = urlsplit(str(response.request.url)).hostname if response.request else "server"
    if status == 404:
        return NotFound(f"Not found ({code}).")
    if status == 429:
        return Throttled(f"{host} is throttling requests ({code}). Retry later.")
    if status == 400:
        return InvalidRequest(f"{host} rejected the request ({code}).")
    if status in (401, 403):
        return Upstream(f"{host} refused access (HTTP {status}, {code}).")
    return Upstream(f"{host} returned HTTP {status} ({code}).")

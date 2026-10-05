"""The one HTTP client for Microsoft calls: auth header, host allowlist, retries, error mapping.

- One ``httpx.AsyncClient`` per process (keep-alive within a call).
- No redirects, and only allowlisted hosts.
- Retries only when the caller marks the request idempotent (GETs by default), honoring
  429 / Retry-After. Writes (``write=True``) are sent once: an answer that never completes, or a
  server error, raises ``WriteOutcomeUnknown`` because the write may have happened.
- Concurrency limiter: Exchange Online allows 4 concurrent requests per app and mailbox.
- Logs metadata only: no URLs with query text, no bodies, no tokens.
- Errors name the operation in progress (``operation()``), the service error code and message
  (shortened), and the request id, so a failure can be traced without exposing mail content.
"""

from __future__ import annotations

import asyncio
import base64
import contextlib
import json
import logging
import re
import time
from collections.abc import Awaitable, Callable, Iterator, Mapping
from contextvars import ContextVar
from typing import Any, BinaryIO, Protocol
from urllib.parse import urlsplit

import httpx

from outlook_connector.domain.errors import (
    AuthenticationRequired,
    ConnectorError,
    Failure,
    InvalidRequest,
    NotFound,
    Throttled,
    Upstream,
    WriteOutcomeUnknown,
)

log = logging.getLogger(__name__)

ALLOWED_HOSTS = frozenset({"graph.microsoft.com", "outlook.cloud.microsoft", "outlook.office.com"})
RETRY_STATUSES = frozenset({429, 502, 503, 504})
MAX_JSON_BYTES = 32 * 1024 * 1024
MAX_CONCURRENCY = 4  # Exchange Online: concurrent requests per app per mailbox
MAX_ERROR_TEXT = 200

_operation: ContextVar[str | None] = ContextVar("operation", default=None)


@contextlib.contextmanager
def operation(name: str) -> Iterator[None]:
    """Name what the code below is doing; errors raised inside it say so ("fetching message bodies")."""
    token = _operation.set(name)
    try:
        yield
    finally:
        _operation.reset(token)


class _Token(Protocol):
    @property
    def value(self) -> str: ...

    def claims(self) -> dict[str, Any]: ...


class TokenSource(Protocol):
    """What the transport and OWS need from the token provider (``auth.tokens.TokenProvider``)."""

    def get_token(
        self, profile: str, *, force_refresh: bool = False, claims_challenge: str | None = None
    ) -> _Token: ...

    def sign_in_command(self, profile: str) -> str: ...


class Transport:
    def __init__(
        self,
        tokens: TokenSource,
        *,
        client: httpx.AsyncClient | None = None,
        max_concurrency: int = MAX_CONCURRENCY,
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

    async def sleep(self, seconds: float) -> None:
        """Wait before a retry (injectable for tests)."""
        await self._sleep(seconds)

    async def request(
        self,
        method: str,
        url: str,
        *,
        profile: str,
        headers: dict[str, str] | None = None,
        json_body: Any = None,
        retry: bool | None = None,
        write: bool = False,
    ) -> httpx.Response:
        """Send a request and return a successful response, or raise a domain error.

        Assumes (not re-checked here): ``url`` is built by ``remote/`` (Graph or OWS); the host allowlist
        below is the one guard every URL passes.
        """
        _check_host(url)
        retry = False if write else (method.upper() == "GET") if retry is None else retry
        attempts = self._max_attempts if retry else 1
        renewal: dict[str, Any] | None = None  # after a 401: how to renew the token, for one request
        renewed = False
        attempt = 0
        while attempt < attempts:
            attempt += 1
            request_headers = {"Authorization": self._bearer(profile, renewal), **(headers or {})}
            renewal = None
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
                raise _no_response(
                    f"Could not reach {urlsplit(url).hostname} ({type(exc).__name__})."
                ) from None
            except httpx.HTTPError as exc:  # sent, but no complete response
                if write:
                    raise _unknown(url, type(exc).__name__) from None
                if attempt < attempts:
                    await self._sleep(min(2**attempt, 20))
                    continue
                raise _no_response(
                    f"No complete response from {urlsplit(url).hostname} ({type(exc).__name__})."
                ) from None
            log.debug(
                "http %s %s -> %s in %dms",
                method,
                _path(url),
                response.status_code,
                (time.monotonic() - started) * 1000,
            )
            if response.status_code == 401:
                if not renewed:  # the token was rejected (revoked, expired early): renew once
                    renewed, renewal = True, _renewal(response)
                    attempt -= 1
                    continue
                raise self._sign_in_required(profile, response)
            if response.status_code in RETRY_STATUSES and attempt < attempts:
                await self._sleep(_retry_after(response, attempt))
                continue
            if response.is_success:
                return response
            if write and response.status_code >= 500:
                raise _unknown(url, f"HTTP {response.status_code}, {error_code(response) or 'no error code'}")
            raise _error_for(response)
        raise AssertionError("unreachable")

    def _bearer(self, profile: str, renewal: dict[str, Any] | None) -> str:
        return f"Bearer {self._tokens.get_token(profile, **(renewal or {})).value}"

    def _sign_in_required(self, profile: str, response: httpx.Response) -> AuthenticationRequired:
        code = error_code(response) or "no error code"
        return AuthenticationRequired(
            f"Microsoft rejected the sign-in for this request (HTTP 401, {code}), also after renewing it.",
            command=self._tokens.sign_in_command(profile),
        )

    async def json(self, method: str, url: str, **kwargs: Any) -> Any:
        response = await self.request(method, url, **kwargs)
        problem = None
        if len(response.content) > MAX_JSON_BYTES:
            problem = "Response too large."
        elif response.content:
            try:
                return response.json()
            except json.JSONDecodeError:
                problem = "Expected a JSON response."
        else:
            return None
        if kwargs.get("write"):  # a write was accepted, but its answer is unreadable
            raise _unknown(url, problem)
        raise Upstream(problem)

    async def download(
        self, url: str, dest: BinaryIO, *, profile: str, headers: dict[str, str] | None = None, max_bytes: int
    ) -> tuple[str | None, int]:
        """Stream a GET response body into ``dest``. Returns (content type, size).

        Retried like any GET: throttling (429/503, after Retry-After), a gateway error, or a
        connection that fails or drops mid-download starts the download again from scratch. What
        still fails is a domain error, never a raw HTTP client exception.

        Assumes (not re-checked here): ``url`` is built by ``remote/``; the host allowlist is checked as
        for ``request``.
        """
        _check_host(url)
        renewal: dict[str, Any] | None = None
        renewed = False
        attempt = 0
        while True:
            attempt += 1
            request_headers = {"Authorization": self._bearer(profile, renewal), **(headers or {})}
            renewal = None
            try:
                async with self._limit, self._client.stream("GET", url, headers=request_headers) as response:
                    if response.status_code == 401:
                        await response.aread()
                        if not renewed:
                            renewed, renewal = True, _renewal(response)
                            attempt -= 1
                            continue
                        raise self._sign_in_required(profile, response)
                    if response.status_code in RETRY_STATUSES and attempt < self._max_attempts:
                        await response.aread()
                        wait = _retry_after(response, attempt)
                    elif not response.is_success:
                        await response.aread()
                        raise _error_for(response)
                    else:
                        dest.seek(0)
                        dest.truncate()
                        size = 0
                        async for chunk in response.aiter_bytes():
                            size += len(chunk)
                            if size > max_bytes:
                                raise InvalidRequest(
                                    f"Download exceeds the {max_bytes // (1024 * 1024)} MiB limit."
                                )
                            dest.write(chunk)
                        return response.headers.get("content-type"), size
            except httpx.HTTPError as exc:
                if attempt >= self._max_attempts:
                    connect = isinstance(exc, httpx.ConnectError | httpx.ConnectTimeout)
                    what = "Could not reach" if connect else "No complete response from"
                    raise _no_response(f"{what} {urlsplit(url).hostname} ({type(exc).__name__}).") from None
                wait = float(min(2**attempt, 20))
            await self._sleep(wait)  # outside the concurrency limit, like request()


def _no_response(text: str) -> Upstream:
    error = Upstream(text)
    error.failure = Failure(status=None, message=text)
    return error


def _unknown(url: str, what: str) -> WriteOutcomeUnknown:
    prefix = f"While {_operation.get()}: " if _operation.get() else ""
    return WriteOutcomeUnknown(
        f"{prefix}{urlsplit(url).hostname} gave no clear answer ({what}). The change may or may not "
        "have been made; it was not retried. Check the mailbox before trying again."
    )


def _check_host(url: str) -> None:
    parts = urlsplit(url)
    if parts.scheme != "https" or parts.hostname not in ALLOWED_HOSTS:
        raise InvalidRequest(f"Refusing to call a non-allowlisted URL host: {parts.hostname}")


def _path(url: str) -> str:
    """The URL path for logs, with ids (long segments) replaced: logs never carry ids in clear."""
    return "/".join("{id}" if len(part) > 32 else part for part in urlsplit(url).path.split("/"))


_CLAIMS = re.compile(r'claims="([^"]+)"')


def _renewal(response: httpx.Response) -> dict[str, Any]:
    """How to renew a token the service rejected: force a refresh, with the CAE claims challenge if any."""
    challenge = _CLAIMS.search(response.headers.get("www-authenticate", ""))
    if challenge:
        try:
            claims = base64.b64decode(challenge.group(1) + "=" * (-len(challenge.group(1)) % 4)).decode()
            return {"claims_challenge": claims}
        except (ValueError, UnicodeDecodeError):
            pass
    return {"force_refresh": True}


def _retry_after(response: httpx.Response, attempt: int) -> float:
    value = response.headers.get("retry-after", "")
    try:
        return min(float(value), 30.0)
    except ValueError:
        return float(min(2**attempt, 20))


def service_error(body: Any) -> tuple[str | None, str | None]:
    """(code, message) of a Microsoft error body ``{"error": {"code", "message"}}``."""
    error = body.get("error") if isinstance(body, dict) else None
    if not isinstance(error, dict):
        return None, None
    code = error.get("code") if isinstance(error.get("code"), str) else None
    message = error.get("message") if isinstance(error.get("message"), str) else None
    return code, message


def error_code(response: httpx.Response) -> str | None:
    owa = response.headers.get("x-owa-error")
    if owa:
        return owa[:120]
    try:
        return service_error(response.json())[0]
    except (json.JSONDecodeError, UnicodeDecodeError):
        return None


def request_id(headers: Mapping[str, str]) -> str | None:
    lowered = {k.lower(): v for k, v in headers.items()}
    value = lowered.get("request-id") or lowered.get("client-request-id")
    return value[:64] if value else None


def shorten(message: str | None) -> str | None:
    """A service message flattened to one line and cut to MAX_ERROR_TEXT."""
    return re.sub(r"\s+", " ", message).strip()[:MAX_ERROR_TEXT] if message else None


def describe_failure(
    *,
    status: int,
    code: str | None,
    message: str | None,
    request: str | None,
    host: str | None = None,
    detail: str | None = None,
) -> ConnectorError:
    """The domain error for a failed Microsoft response, with sanitized diagnostics.

    The service message is shortened and flattened to one line; it never contains tokens, and
    Graph error messages do not echo mail content.
    """
    where = host or "Microsoft Graph"
    parts = [f"HTTP {status}", code or "no error code"]
    text = ", ".join(parts)
    if message:
        text += ": " + (shorten(message) or "")
    if request:
        text += f"; request-id {request}"
    error = _failure_error(
        status, f"While {_operation.get()}: " if _operation.get() else "", where, text, detail
    )
    error.failure = Failure(status=status, code=code, message=shorten(message), request_id=request)
    return error


def _failure_error(status: int, prefix: str, where: str, text: str, detail: str | None) -> ConnectorError:
    suffix = f" {detail}" if detail else ""
    if status == 404:
        return NotFound(f"{prefix}Not found ({text}).{suffix}")
    if status == 429:
        return Throttled(
            f"{prefix}{where} is throttling requests ({text}).{suffix} Outlook allows about 4 "
            "concurrent requests and 10,000 requests per 10 minutes per mailbox; wait and retry, "
            "and avoid parallel calls."
        )
    if status == 400:
        return InvalidRequest(f"{prefix}{where} rejected the request ({text}).{suffix}")
    if status == 403:  # authorization denied for this item or action; the sign-in itself is fine
        return Upstream(f"{prefix}{where} denied access ({text}).{suffix}")
    return Upstream(f"{prefix}{where} returned an error ({text}).{suffix}")


def _error_for(response: httpx.Response) -> ConnectorError:
    host = urlsplit(str(response.request.url)).hostname if response.request else None
    message = None
    with contextlib.suppress(json.JSONDecodeError, UnicodeDecodeError):
        message = service_error(response.json())[1]
    return describe_failure(
        status=response.status_code,
        code=error_code(response),
        message=message,
        request=request_id(response.headers),
        host=host,
    )

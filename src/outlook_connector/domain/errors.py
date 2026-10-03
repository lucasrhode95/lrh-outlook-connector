"""Domain errors. Each surface maps these to its own protocol (CLI exit codes, MCP tool errors, HTTP)."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Failure:
    """What Microsoft answered to one failed request: the HTTP status (None: no response), the
    service error code, its message (shortened, one line) and the request id."""

    status: int | None
    code: str | None = None
    message: str | None = None
    request_id: str | None = None

    def describe(self) -> str:
        """ "HTTP 429 TooManyRequests, request-id <id>" or "no response from Microsoft"."""
        text = (
            f"HTTP {self.status} {self.code or 'no error code'}"
            if self.status
            else "no response from Microsoft"
        )
        return f"{text}, request-id {self.request_id}" if self.request_id else text


class ConnectorError(Exception):
    """Base class for errors the surfaces know how to present.

    ``failure``: the Microsoft answer behind it, when a request failed (set by remote/transport.py).
    """

    failure: Failure | None = None


class ConfigurationError(ConnectorError):
    """The application configuration is invalid (for example a denied client/scope pair)."""


class AuthenticationRequired(ConnectorError):
    """No usable credentials. ``command`` is the exact command the user should run."""

    def __init__(self, message: str, *, command: str) -> None:
        super().__init__(f"{message} Run `{command}` in a terminal, then retry.")
        self.command = command


class AccountMismatch(ConnectorError):
    """Credentials belong to a different Microsoft account than the one already bound."""


class SecureStorageUnavailable(ConnectorError):
    """Encrypted token storage cannot be used. The connector never falls back to plaintext silently."""


class InvalidRequest(ConnectorError):
    """The caller asked for something malformed or unsupported."""


class NotFound(ConnectorError):
    """The requested item does not exist (or no longer exists) on the server or locally."""


class Throttled(ConnectorError):
    """Microsoft is throttling requests; retry later."""


class Upstream(ConnectorError):
    """A Microsoft service could not be reached or returned an unexpected failure."""


class WriteOutcomeUnknown(ConnectorError):
    """A write reached Microsoft but no complete answer came back: it may or may not have happened.

    Never retried automatically. The caller checks the mailbox (e.g. Sent Items) instead.
    """

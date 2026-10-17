"""Domain errors. Each surface maps these to its own protocol (CLI exit codes, MCP tool errors, HTTP)."""

from __future__ import annotations

from dataclasses import dataclass

from m365_auth import (
    AccountMismatch,
    AuthenticationRequired,
    SecureStorageUnavailable,
)
from m365_auth import (
    AuthError as ConnectorError,
)


@dataclass(frozen=True)
class Failure:
    """What Microsoft answered to one failed request: the HTTP status (None: no response), the
    service error code, its message (shortened, one line) and the request id."""

    status: int | None
    code: str | None = None
    message: str | None = None
    request_id: str | None = None


class ConfigurationError(ConnectorError):
    """The application configuration is invalid (for example a denied client/scope pair)."""


class InvalidRequest(ConnectorError):
    """The caller asked for something malformed or unsupported."""


class DownloadLimitExceeded(InvalidRequest):
    """A download exceeded the connector's local size limit."""

    def __init__(self, limit_bytes: int) -> None:
        self.limit_bytes = limit_bytes
        mebibytes, remainder = divmod(limit_bytes, 1024 * 1024)
        self.limit = f"{mebibytes} MB" if remainder == 0 else f"{limit_bytes} bytes"
        limit_text = f"{self.limit} limit" if remainder == 0 else f"limit of {self.limit}"
        super().__init__(f"Download exceeds the connector's local {limit_text}.")


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


__all__ = [
    "AccountMismatch",
    "AuthenticationRequired",
    "ConfigurationError",
    "ConnectorError",
    "DownloadLimitExceeded",
    "Failure",
    "InvalidRequest",
    "NotFound",
    "SecureStorageUnavailable",
    "Throttled",
    "Upstream",
    "WriteOutcomeUnknown",
]

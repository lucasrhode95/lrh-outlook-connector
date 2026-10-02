"""Domain errors. Each surface maps these to its own protocol (CLI exit codes, MCP tool errors, HTTP)."""

from __future__ import annotations


class ConnectorError(Exception):
    """Base class for errors the surfaces know how to present."""


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
    """A write may or may not have happened (for example a timeout after sending). Never retry blindly."""

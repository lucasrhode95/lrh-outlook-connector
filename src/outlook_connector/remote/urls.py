"""Central URL policy for Microsoft Graph and Outlook calls.

Transport validates every outbound URL against the host allowlist. Graph continuation links are
additionally constrained to the versioned Graph root before they reach transport.
"""

from __future__ import annotations

from urllib.parse import urlsplit

from outlook_connector.domain.errors import InvalidRequest

ALLOWED_HOSTS = frozenset({"graph.microsoft.com", "outlook.cloud.microsoft", "outlook.office.com"})
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"


def validate_url(url: str) -> None:
    """Reject outbound URLs outside the HTTPS Microsoft host allowlist."""
    parts = urlsplit(url)
    if parts.scheme != "https" or parts.hostname not in ALLOWED_HOSTS:
        raise InvalidRequest(f"Refusing to call a non-allowlisted URL host: {parts.hostname}")


def graph_url(path_or_url: str) -> str:
    """Resolve Graph paths and accept continuation links only under the Graph API root."""
    parts = urlsplit(path_or_url)
    if parts.scheme or parts.netloc:
        validate_url(path_or_url)
        if not path_or_url.startswith(GRAPH_ROOT + "/"):
            raise InvalidRequest("Continuation link does not point at Microsoft Graph.")
        return path_or_url
    return GRAPH_ROOT + path_or_url

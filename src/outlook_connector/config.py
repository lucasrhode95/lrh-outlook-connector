"""Static configuration: first-party client profiles, denied pairs and local paths.

The profiles encode the tenant-specific capability split established in
docs/outlook-api-research.md §2 (Graph and Outlook API profiles). To adapt to a
tenant with different Graph availability, change the profiles here; see
docs/architecture.md §6.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

APP_NAME = "lrh-outlook-connector"
CLI_NAME = "outlook-connector"
TENANT = "organizations"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT}"

OUTLOOK_MOBILE_CLIENT_ID = "27922004-5251-4030-b22d-91ecd9a37ea4"
ONE_OUTLOOK_WEB_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48"


@dataclass(frozen=True)
class TokenProfile:
    """One named credential: a Microsoft first-party public client and the resource scopes it requests."""

    name: str
    client_id: str
    scopes: tuple[str, ...]
    purpose: str


PROFILES: dict[str, TokenProfile] = {
    "graph": TokenProfile(
        name="graph",
        client_id=OUTLOOK_MOBILE_CLIENT_ID,
        scopes=("https://graph.microsoft.com/Mail.Read",),
        purpose="Microsoft Graph APIs (Outlook Mobile client)",
    ),
    "outlook": TokenProfile(
        name="outlook",
        client_id=ONE_OUTLOOK_WEB_CLIENT_ID,
        scopes=("https://outlook.office.com/.default",),
        purpose="Outlook APIs: OWS and Cloud Settings (One Outlook Web client)",
    ),
}

# Client/scope pairs Microsoft refused with AADSTS65002 (research §2). Never request them.
DENIED_PAIRS: frozenset[tuple[str, str]] = frozenset(
    {
        ("d3590ed6-52b3-4102-aeff-aad2292ab01c", "https://graph.microsoft.com/Mail.Read"),
        (OUTLOOK_MOBILE_CLIENT_ID, "https://graph.microsoft.com/Mail.Send"),
        (OUTLOOK_MOBILE_CLIENT_ID, "https://graph.microsoft.com/Mail.ReadWrite"),
        (OUTLOOK_MOBILE_CLIENT_ID, "https://graph.microsoft.com/Mail.ReadWrite.Shared"),
        (ONE_OUTLOOK_WEB_CLIENT_ID, "https://graph.microsoft.com/Mail.Read"),
        (ONE_OUTLOOK_WEB_CLIENT_ID, "https://graph.microsoft.com/Mail.ReadWrite"),
        (ONE_OUTLOOK_WEB_CLIENT_ID, "https://graph.microsoft.com/Mail.Send"),
    }
)


def data_dir() -> Path:
    """Per-user application data directory. ``OUTLOOK_CONNECTOR_HOME`` overrides it (tests, portability).

    On Windows it is ``%USERPROFILE%\\.lrh-outlook-connector``, outside AppData: Windows redirects
    files that packaged apps (such as the Claude desktop app, and the MCP servers it starts) create
    under AppData into a private per-app copy, which a terminal and Explorer never see. That split
    the token cache, the store and the exports in two.
    """
    override = os.environ.get("OUTLOOK_CONNECTOR_HOME")
    if override:
        return Path(override).expanduser().resolve()
    if _windows():
        return Path.home() / f".{APP_NAME}"
    root = os.environ.get("XDG_DATA_HOME")
    return (Path(root) if root else Path.home() / ".local" / "share") / APP_NAME


def _windows() -> bool:
    return os.name == "nt"


def token_cache_path(*, unsecure: bool) -> Path:
    """Encrypted cache by default. The plaintext development cache is a separate, clearly named file."""
    name = "token-cache.plaintext-dev.json" if unsecure else "token-cache.bin"
    return data_dir() / name

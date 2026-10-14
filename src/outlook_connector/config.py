"""Static configuration: first-party client profiles, denied pairs and local paths.

The profiles encode the tenant-specific capability split established in
docs/api-research.md §2 (Graph and Outlook API profiles). To adapt to a
tenant with different Graph availability, change the profiles here; see
docs/architecture.md §6.
"""

from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "lrh-outlook-connector"
CLI_NAME = "outlook-connector"

# Shared-cache profile whose Graph token carries calendar read scopes (Calendars.Read). The Outlook
# Graph profile has none. Calendar reads still belong to Outlook; this only chooses the token.
# Verified live 2026-10-12 (docs/api-research.md §3.17).
CALENDAR_PROFILE = "teams.graph"


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

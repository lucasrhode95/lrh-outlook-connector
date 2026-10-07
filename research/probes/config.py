"""Editable Microsoft defaults with private overrides in .local/probe-config.json."""

from __future__ import annotations

import json
import os
from pathlib import Path
from urllib.parse import quote, urlsplit

ROOT = Path(__file__).resolve().parents[2]
LOCAL = Path(os.environ.get("OUTLOOK_PROBE_HOME", str(ROOT / ".local"))).expanduser()


def load_config() -> dict:
    """Entry point: validate local configuration before token or network use."""
    path = LOCAL / "probe-config.json"
    settings = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    defaults = {
        "tenant": "organizations",
        "expected_user": "",
        "login_url": "https://login.microsoftonline.com",
        "graph_resource": "https://graph.microsoft.com",
        "graph_url": "https://graph.microsoft.com/v1.0",
        "ows_url": "https://outlook.cloud.microsoft/owa/service.svc",
        "profiles": {
            "read": [
                "27922004-5251-4030-b22d-91ecd9a37ea4",
                "https://graph.microsoft.com/Mail.Read",
            ],
            "write": [
                "9199bf20-a13f-4107-85dc-02114787ef48",
                "https://outlook.office.com/.default",
            ],
            "search": [
                "9199bf20-a13f-4107-85dc-02114787ef48",
                "https://outlook.office.com/search/.default",
            ],
        },
        "substrate_urls": [
            "https://outlook.office.com/searchservice/api/v2/query",
            "https://outlook.office.com/search/api/v2/query",
            "https://substrate.office.com/searchservice/api/v2/query",
        ],
        "denied_pairs": [],
    }
    if not isinstance(settings, dict) or settings.keys() - defaults.keys():
        raise SystemExit(
            "probe-config.json must be an object containing only documented settings."
        )
    defaults.update(settings)
    for key in ("tenant", "expected_user"):
        if not isinstance(defaults[key], str) or (
            key == "tenant" and not defaults[key].strip()
        ):
            raise SystemExit(f"{key} must be a string; tenant must not be empty.")
        defaults[key] = defaults[key].strip()
    urls = [
        defaults[key] for key in ("login_url", "graph_resource", "graph_url", "ows_url")
    ]
    if not isinstance(defaults["substrate_urls"], list):
        raise SystemExit("substrate_urls must be a list of HTTPS URLs.")
    urls.extend(defaults["substrate_urls"])
    for url in urls:
        parsed = urlsplit(url) if isinstance(url, str) else None
        if (
            not parsed
            or parsed.scheme != "https"
            or not parsed.hostname
            or parsed.username
            or parsed.password
            or parsed.query
            or parsed.fragment
        ):
            raise SystemExit(
                "Probe URLs must use HTTPS without credentials, query strings or fragments."
            )
    profiles = defaults["profiles"]
    if not isinstance(profiles, dict) or set(profiles) != {"read", "write", "search"}:
        raise SystemExit(
            "profiles must define read, write and search client/scope pairs."
        )
    pairs = defaults["denied_pairs"]
    if not isinstance(pairs, list):
        raise SystemExit("denied_pairs must be a list of client/scope pairs.")
    for pair in [*profiles.values(), *pairs]:
        if (
            not isinstance(pair, list)
            or len(pair) != 2
            or not all(isinstance(value, str) and value.strip() for value in pair)
        ):
            raise SystemExit(
                "Each client/scope pair must contain two nonempty strings."
            )
    defaults["authority"] = (
        defaults["login_url"].rstrip("/")
        + "/"
        + quote(defaults["tenant"], safe="")
        + "/oauth2/v2.0"
    )
    defaults["allowed_hosts"] = {urlsplit(url).hostname for url in urls}
    return defaults


SETTINGS = load_config()

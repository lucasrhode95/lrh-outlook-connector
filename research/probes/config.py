"""Required private probe configuration in .local/probe-config.json."""

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
    try:
        settings = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise SystemExit(
            f"Missing probe configuration: {path}. "
            "Copy research/probe-config.example.json there and configure your environment."
        ) from None
    except (OSError, json.JSONDecodeError):
        raise SystemExit("Cannot read probe-config.json as valid JSON.") from None
    required = {
        "tenant",
        "expected_user",
        "login_url",
        "graph_resource",
        "graph_url",
        "ows_url",
        "profiles",
        "substrate_urls",
        "denied_pairs",
    }
    if not isinstance(settings, dict):
        raise SystemExit("probe-config.json must contain an object.")
    missing = required - settings.keys()
    unknown = settings.keys() - required
    if missing:
        raise SystemExit(
            "Missing probe configuration settings: " + ", ".join(sorted(missing))
        )
    if unknown:
        raise SystemExit(
            "Unknown probe configuration settings: " + ", ".join(sorted(unknown))
        )
    for key in ("tenant", "expected_user"):
        if not isinstance(settings[key], str) or (
            key == "tenant" and not settings[key].strip()
        ):
            raise SystemExit(f"{key} must be a string; tenant must not be empty.")
        settings[key] = settings[key].strip()
    urls = [
        settings[key] for key in ("login_url", "graph_resource", "graph_url", "ows_url")
    ]
    if not isinstance(settings["substrate_urls"], list):
        raise SystemExit("substrate_urls must be a list of HTTPS URLs.")
    urls.extend(settings["substrate_urls"])
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
    profiles = settings["profiles"]
    if not isinstance(profiles, dict) or set(profiles) != {"graph", "outlook", "search"}:
        raise SystemExit(
            "profiles must define graph, outlook and search client/scope pairs."
        )
    pairs = settings["denied_pairs"]
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
    settings["authority"] = (
        settings["login_url"].rstrip("/")
        + "/"
        + quote(settings["tenant"], safe="")
        + "/oauth2/v2.0"
    )
    settings["allowed_hosts"] = {urlsplit(url).hostname for url in urls}
    return settings


SETTINGS = load_config()

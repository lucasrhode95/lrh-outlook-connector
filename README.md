# Outlook connector

A local connector for one user's Exchange Online mailbox. It has two surfaces: an MCP server for agents and a small local web UI for exports. Reads go through Microsoft Graph. Writes go through Outlook Web, because Graph write access is unavailable to the usable Microsoft first-party clients.

- **What it must do:** [docs/outlook-requirements-v4.md](docs/outlook-requirements-v4.md)
- **How it is built:** [docs/architecture.md](docs/architecture.md)
- **What Microsoft allows:** [docs/outlook-api-research.md](docs/outlook-api-research.md)
- **Work register:** [docs/outlook-roadmap.md](docs/outlook-roadmap.md)

## Setup

Python ≥ 3.12.

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -e . --group dev
```

The project also works with `uv sync`.

## Sign in

Sign-in uses Microsoft's device-code flow and is only ever started from this command. MCP and UI processes never prompt.

```bash
outlook-connector auth read     # Graph reads (needed for everything)
outlook-connector auth write    # send and mailbox changes (later phases)
outlook-connector status        # offline: account, profiles, cache location
outlook-connector status --check  # also refreshes each profile's token
```

Tokens are stored encrypted (DPAPI on Windows, Keychain on macOS, libsecret on Linux) under the user data directory. If encryption is unavailable, the connector refuses to store tokens.

For development only, `--unsecure` uses a separate **plaintext** cache file in the same directory, and prints a warning each time.

## Tests

```bash
.venv/Scripts/python -m pytest
.venv/Scripts/python -m ruff check .
```

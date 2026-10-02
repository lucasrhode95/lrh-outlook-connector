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
outlook-connector status --check  # also refreshes each token (only read is required)
```

Tokens are stored encrypted (DPAPI on Windows, Keychain on macOS, libsecret on Linux) under the user data directory. If encryption is unavailable, the connector refuses to store tokens.

For development only, `--unsecure` uses a separate **plaintext** cache file in the same directory, and prints a warning each time.

## Use it

**Agents (MCP).** Register the stdio server with your MCP client, for example Claude Code:

```bash
claude mcp add lrh-outlook -- C:/Users/<you>/dev/lrh-outlook-connector/.venv/Scripts/outlook-connector.exe mcp
```

The client starts one `outlook-connector mcp` process per session. Tools: `list_folders`, `list_messages`,
`search_messages`, `get_thread`, `get_message`, `list_attachments`, `download_attachment`,
`save_message_mime`, `export_messages`, `auth_status`. They are read-only: nothing changes the mailbox,
not even read state.

**You (local UI).** Start it when you need it. It opens a browser tab and stops after 30 idle minutes:

```bash
outlook-connector ui
```

Browse folders or recent mail grouped by thread, search the mailbox, filter what is loaded, read
messages, tick threads or single messages, and export them as one `.txt` or `.zip`
(per thread, all in one file, or one file per message; attachments optional).

Exports and downloaded attachments go to the data directory (`%LOCALAPPDATA%\lrh-outlook-connector`)
and are removed after a week. The local store there keeps only the folder cache and copies of messages
this app has read, so mail deleted on the server stays readable and is labelled as deleted.

## Tests

```bash
.venv/Scripts/python -m pytest
.venv/Scripts/python -m ruff check .
```

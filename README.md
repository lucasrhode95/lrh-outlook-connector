# Outlook connector

A local connector for one user's Exchange Online mailbox. It has two surfaces: an MCP server for agents and a small local web UI for exports. Reads go through Microsoft Graph. Writes (drafts, sending and mailbox changes) go through Outlook Web, because Graph write access is unavailable to the usable Microsoft first-party clients.

- **What it must do:** [docs/outlook-requirements-v4.md](docs/outlook-requirements-v4.md)
- **How it is built:** [docs/architecture.md](docs/architecture.md)
- **What Microsoft allows:** [docs/outlook-api-research.md](docs/outlook-api-research.md)
- **Work register:** [docs/outlook-roadmap.md](docs/outlook-roadmap.md)

## Setup

Python ≥ 3.12.

```bash
uv sync                      # uses the committed uv.lock
```

Or with pip: `python -m venv .venv` then `.venv/Scripts/python -m pip install -e . --group dev`.

## Sign in

Sign-in uses Microsoft's device-code flow and is only ever started from this command. MCP and UI processes never prompt.

```bash
outlook-connector auth read     # Graph reads (needed for everything)
outlook-connector auth write    # drafts, sending and mailbox changes
outlook-connector status        # offline: account, profiles, cache location
outlook-connector status --check  # also refreshes each token (only read is required)
```

Tokens are stored encrypted (DPAPI on Windows, Keychain on macOS, libsecret on Linux) in the data directory (`%USERPROFILE%\.lrh-outlook-connector` on Windows; `OUTLOOK_CONNECTOR_HOME` overrides it). If encryption is unavailable, the connector refuses to store tokens.

For development only, `--unsecure` uses a separate **plaintext** cache file in the same directory, and prints a warning each time.

## Use it

**Agents (MCP).** Register the stdio server with your MCP client, for example Claude Code:

```bash
claude mcp add lrh-outlook -- C:/Users/<you>/dev/lrh-outlook-connector/.venv/Scripts/outlook-connector.exe mcp
```

The client starts one `outlook-connector mcp` process per session. Read tools: `list_folders`,
`list_messages`, `search_messages`, `get_thread`, `get_message`, `list_attachments`,
`download_attachment`, `save_message_mime`, `export_messages`, `auth_status`, `propose_email`. Reading
never changes the mailbox, not even read state.

Write tools (they need `outlook-connector auth write`): `create_draft` saves a plain-text message or
reply into Drafts and never sends it. `send_email` sends only a message you confirmed: the agent calls
`propose_email`, shows you the exact message and its confirmation code, and passes that code once you
confirm; any change to the recipients, subject or body afterwards is refused. A send is never
retried; if Outlook gives no clear answer, the connector looks in Sent Items and otherwise tells you
to check before anything is sent again.

Mailbox changes (same sign-in): `set_read_state`, `set_flag`, `move_messages` and
`delete_messages` act on explicit message ids (up to 100 per call; read state also per conversation)
and return a result per message: done, unchanged, not found, failed or unknown. Delete moves to
Deleted Items and never deletes permanently.

Every tool follows the same scope rules: Deleted Items, Junk Email and Sync Issues (the copies
Outlook files when two versions of an item collide while syncing) are left out unless
`include_deleted_items=true` (a folder you name is always included; a subfolder counts with its
parent, so a folder you deleted in Outlook counts as Deleted Items). `received_only=true` also leaves
out Sent Items, Drafts and Outbox, and `coverage.excluded` counts what was left out. Copies of one
message (mail sent to yourself or to a list you are on) are shown once, with `also_in` naming the
other folders.

**Out of reach:** hidden folders, and items outside the mail folders (Teams meeting records,
settings and other non-mail items), are never listed, searched, counted, threaded or exported, and
`list_folders` does not show them. Search covers mail only. List and search results are compact by default (`detail="full"` for every field);
search hits carry the conversation's message count, and `list_messages(include_total=true)` returns the
server's count for the window.

`export_messages` takes conversations, message ids and/or a range (`since`, `until`, `folder`,
`received_only`), up to 2,000 messages (`limit` lowers that). `format="jsonl"` writes one JSON record
per message for agents; `txt` is for people. It reports what it left out or could not fetch; every
exported message carries its message, conversation and Internet ids, and a message that exists in
several folders is exported once, with `also_in` naming the other folders.

**Throttling.** Microsoft Graph allows about 4 concurrent requests and 10,000 requests per 10 minutes
per mailbox; each item of a `$batch` (at most 20) counts. The connector keeps at most 4 requests and
2 batches in flight, re-sends throttled batch items in new batches of at most 20 after the advised
delay, and reports items that stay throttled instead of failing the whole call. Agents are told to
avoid parallel tool calls and to prefer one range export over many small calls. Errors name the
operation, the Graph error code and message, and the request id. A rejected token (401) is renewed
once before the connector asks you to sign in again; "access denied" (403) never asks for a sign-in.

**You (local UI).** Start it when you need it. It opens a browser tab; an open tab keeps it running, and
it stops 30 minutes after the last one is closed:

```bash
outlook-connector ui
```

The UI opens on the Inbox. Browse folders or recent mail grouped by thread (a conversation with one
message is a plain row; a real thread shows its message count across all folders, newest message on
top), search the mailbox, filter what is loaded, read messages and download their attachments, tick
threads or single messages, and export them as one `.txt` or `.zip` (one file per thread, one for
everything, or one per message; attachments optional). "export this view" exports the whole current
folder and date range. Deleted Items, Junk and Sync Issues are left out unless you tick
"Deleted / Junk" (they are always shown inside those folders). Hidden folders are not listed.

Exports and downloaded attachments go to the data directory (`%USERPROFILE%\.lrh-outlook-connector`)
and are removed after a week. The local store there keeps only the folder cache and the summaries of
listed messages (for the instant preview). Mail deleted on the server is gone here too.

## Tests

```bash
uv run pytest
uv run ruff check .
uv run ruff format --check .
uv run pyright
```

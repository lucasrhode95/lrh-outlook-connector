# Outlook connector

A local connector for one user's Exchange Online mailbox. It has two surfaces: an MCP server for agents and a small local web UI for exports. Both run as short-lived local processes started on demand (one per agent session; the UI until closed or idle), not as a hosted, long-running MCP or HTTP server. Reads go through Microsoft Graph. Writes (drafts, sending and mailbox changes) go through Outlook Web, because Graph write access is unavailable to the usable Microsoft first-party clients.

- **What it must do:** [docs/outlook-requirements-v4.md](docs/outlook-requirements-v4.md)
- **How it is built:** [docs/architecture.md](docs/architecture.md)
- **What Microsoft allows:** [docs/outlook-api-research.md](docs/outlook-api-research.md)
- **Work register:** [docs/outlook-roadmap.md](docs/outlook-roadmap.md)

**Terminology:** this project uses **conversation** consistently for the Exchange/Graph `conversationId` grouping. A conversation is what email users and many clients commonly call a **thread**; there is no separate thread id in this connector.

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
`list_messages`, `search_messages`, `get_conversation`, `get_message`, `list_attachments`,
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

Every tool follows the same scope rules: Deleted Items and Junk Email are left out unless
`include_deleted_items=true` (a folder you name is always included; a subfolder counts with its
parent, so a folder you deleted in Outlook counts as Deleted Items). `include_sent_items=false`
also leaves out Sent Items, Drafts and Outbox (included by default), and
`include_meeting_mail=false` leaves out invitations, RSVPs and cancellations in list, search and
range exports (a conversation with real replies still shows through them). Every flag points the
same way: true shows more mail, false filters more. The UI's "Invites / RSVPs" switch starts off. `coverage.excluded` counts what was left out. Copies of one
message (mail sent to yourself or to a list you are on) are shown once, with `also_in` naming the
other folders.

**Out of reach:** hidden folders, and items outside the mail folders (Teams meeting records,
settings and other non-mail items), are never listed, searched, counted, grouped into conversations or exported, and
`list_folders` does not show them. Search covers mail only. List and search results are compact by default (`detail="full"` for every field);
search hits carry the conversation's message count, and `list_messages(include_total=true)` returns the
server's count for the window.

`export_messages` takes conversations, message ids and/or a range (`since`, `until`, `folder`,
`include_sent_items=false`), up to 2,000 messages (`limit` lowers that). `format="jsonl"` writes one
JSON record per message for agents; `txt` is for people. Every exported message carries its message,
conversation and Internet ids, and a message that exists in several folders is exported once, with
`also_in` naming the other folders. Messages selected by id are read from the server: if any of them
cannot be read (deleted or moved meanwhile, or still throttled), the export fails, writes nothing and
says which ones and what to do. Anything else that cannot be exported (a body, an attachment, an
attachment listing) is marked in place, and the result's `error_summary` (the file header's "Export
errors" line) says how many and why. In TXT the mark is a block:

```text
[EXPORT ERROR] The body of this message could not be fetched.
  Step:   fetching message bodies
  Error:  HTTP 429 TooManyRequests, request-id <id>
  Likely: Microsoft throttled the mailbox (about 4 parallel requests or 10,000 per 10 minutes); the message itself is fine
  Fix:    export it again in a few minutes
```

In JSONL it is an `export_error` object (step, status, code, message, request id, likely cause,
`retry`, fix) on the message record or on the failed attachment record (`attachments_export_error`
when the attachments could not be listed). `get_conversation` marks a body it cannot fetch the same way, sets `export_error` on that message and
counts them in `body_errors`, so a caller never has to read the text to find out.

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

The UI opens on the Inbox. Browse folders or recent mail grouped by conversation (a conversation with one
message is a plain row; a real conversation shows its message count across all folders, newest message on
top), search the mailbox, filter what is loaded, read messages and download their attachments, tick
conversations or single messages, and export them as one `.txt` or `.zip` (one file per conversation, one for
everything, or one per message; attachments optional). "export this view" exports the whole current
folder and date range. Deleted Items and Junk are left out unless you tick "Deleted / Junk"
(they are always shown inside those folders). Hidden folders and Sync Issues (classic Outlook's
conflict copies) are not listed at all.

Exports and downloaded attachments go to the data directory (`%USERPROFILE%\.lrh-outlook-connector`)
and are removed after a week. The local store there keeps only the account binding and the folder
cache: no message data, so mail deleted on the server is gone here too. After an export with errors,
the UI shows the file's "Export errors" line.

## Tests

```bash
uv run pytest
uv run ruff check .
uv run ruff format --check .
uv run pyright
```


H25 UI reader: selecting a message automatically follows every `next_offset` until the chosen
unique/full body is complete, with no total body-size ceiling. The web endpoint accepts `offset`;
the service retains its bounded per-request body chunks. No manual continuation button is required.
A continuation failure reports that the complete message could not be loaded, rather than displaying
a partial body as complete. Selecting a different message stops scheduling old continuations using
the existing reader request counter; the request already in flight is allowed to finish.

Reader JavaScript regression tests use Node's built-in runner (Node 18+), with no package dependencies:
`node --test tests/surfaces/message_reader.test.cjs`. They exercise the shipped frontend with a minimal
test DOM: complete continuation loading, long/short bodies, failures and overlapping selections.

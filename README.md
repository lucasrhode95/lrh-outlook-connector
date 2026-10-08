# Outlook connector

A local connector for one user's Exchange Online mailbox. It has two surfaces: an MCP server for agents and a small local web UI for exports. Both run as short-lived local processes started on demand (one per agent session; the UI until closed or idle), not as a hosted, long-running MCP or HTTP server. Reads go through Microsoft Graph. Writes (drafts, sending and mailbox changes) go through Outlook Web, because the tested deployment did not grant Graph writes to the selected Microsoft first-party clients. Other deployments should rerun the configurable [research probes](research/README.md) before choosing their authentication profiles and backends.

- **What it must do:** [docs/outlook-requirements-v4.md](docs/outlook-requirements-v4.md)
- **How it is built:** [docs/architecture.md](docs/architecture.md)
- **What Microsoft allows:** [docs/outlook-api-research.md](docs/outlook-api-research.md)
- **Work register:** [docs/outlook-roadmap.md](docs/outlook-roadmap.md)

**Terminology:** this project uses **conversation** consistently for the Exchange/Graph `conversationId` grouping. A conversation is what email users and many clients commonly call a **thread**; there is no separate thread id in this connector.

Research probes call live endpoints with explicit authorization. Browser captures and
downloaded bundles are temporary exploratory inputs kept outside the repository; the
published research contains endpoint documentation and findings, not capture files or
capture-processing tooling.

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

The client starts one `outlook-connector mcp` process per session. Read tools: list_folders, list_messages, search_messages, get_conversation, get_message,
list_attachments, download_attachment, save_message_mime, export_messages, auth_status,
list_signatures and get_signature. Reading never changes the mailbox, not even read state.

Attachment and message MIME downloads share a 150 MB limit. Download a larger attachment directly from Outlook.

Write tools (they need the write sign-in): native signatures use list_signatures, get_signature,
create_signature, update_signature, delete_signature and set_default_signature. These manage
Outlook's native roaming-signature settings; the organization's recipient-dependent add-in is not
handled. Names are exact and case-sensitive, cannot contain commas, and a rename is create-new then
delete-old. A deleted selected signature can leave a dangling default, so choose a valid default when
needed. Every settings write is sent once.

create_draft accepts exactly one of text_body or html_body, for new mail and replies. It freshly
resolves and inserts the corresponding Outlook native default unless include_signature=false; an
optional signature selects one exact name. A missing or unreadable selection raises rather than
falling back. Embedded signature images become inline attachments with CID references. It saves once
and requires Graph read-back, returning the Microsoft draft id, full server text/HTML, and verification
findings. Plain text is escaped without Markdown conversion; intentional HTML is passed through
except active web content. Replies report quoted-history checks. Drafts are composed once; there is
no edit tool. To change a draft, create a replacement with the full intended content (use the same
reply_to_message_id for a reply), verify its read-back, then move the old draft to Deleted Items
with delete_messages and use the new id. Never delete the old draft first. Tell the user the
previous version is in Deleted Items. Body edits and attachments added in Outlook are not carried
into the replacement. The replacement resolves the current default again; pass signature again to
keep a specifically selected signature. There is no version check against the old draft; read it
first if needed. Every change gets a new id, and a subject- or recipient-only change also recreates
the whole draft.

Inbox rules (write sign-in, including reads): `list_rules`, `create_rule`, `update_rule`,
`reorder_rules`, `delete_rule`. Conditions: From, Sent to, Subject contains, Subject-or-body contains.
Actions: Move to folder and Stop processing. Call a write without `user_confirmation` to propose it;
show the persistent change and returned RULE code, then repeat with that code after explicit human
confirmation. Every write is sent once and read back; changed server state invalidates confirmation.
Unsupported rules are read-only. Reordering is refused if any unsupported rule is present because
OWS resubmits the whole rule list. Enable/disable uses a separate update from field edits.

Mailbox changes (same sign-in): `set_read_state`, `set_flag`, `move_messages` and
`delete_messages` act on explicit message ids (up to 100 per call) and return per-message results
and counts: done, unchanged, not found, failed or unknown. `set_read_state` also takes conversation
ids: each expands to its messages in scope (all copies, up to the 1,000 messages the server lists per
conversation; `notes` says when one was cut there), and the 100 limit applies to explicit ids only.
When more than 100 messages are selected that way, `counts` covers them all, while `results` lists
only explicit ids and the messages that did not end done or unchanged.
`continue_on_error=true` (default) attempts later chunks after errors. With false, later unsent
messages return failed with detail `not sent`. Ambiguous writes are read back; failed read-back
remains unknown. Writes are never retried. Delete moves to
Deleted Items and never deletes permanently.

Tools that accept mailbox scope use one `scope` object with independent keys: `scope.sent_items` (default true) includes Sent Items, Drafts and Outbox; `scope.meeting_mail` (default true) includes invitations, RSVPs and cancellations; and `scope.deleted_items` (default false) includes Deleted Items and Junk Email. A folder you name is always included; subfolders count with their parent, so a folder you deleted in Outlook counts as Deleted Items. True shows more mail, false filters more. The UI's "Invites / RSVPs" switch starts off. Web GET requests use `sent_items`, `meeting_mail` and `deleted_items` query parameters; JSON requests such as export carry a nested `scope` object. Outside export, an operation rejects a non-default key it cannot apply. Export scope only narrows folder/date-window selections; selected conversations and explicit messages stay whole.

**Out of reach:** hidden folders, and items outside the mail folders (Teams meeting records,
settings and other non-mail items), are never listed, searched, counted, grouped into conversations
or picked up by a folder/date-window or conversation export, and `list_folders` does not show them. Search covers
mail only. A message id you name is the exception: `get_message` and `export_messages` read it
wherever it is. List and search results are compact by default (`detail="full"` for every field);
search hits carry the conversation's message count, and `list_messages(include_total=true)` returns the
server's count for the window (copies counted separately, and meeting mail included even when
`scope.meeting_mail=false` hides it).

`export_messages` takes conversations, message ids and/or a folder/date window (`since`, `until`,
`folder`), up to 2,000 messages (`limit` lowers that). Scope narrows only folder/date-window
selections. Conversation and explicit message selections stay whole; scope alone is not a selection.
`format="jsonl"` writes one JSON record per message for agents; `txt` is for people. Every exported message carries its message,
conversation and Internet ids, and a message that exists in several folders is exported once, with
`also_in` naming the other folders. Messages selected by id are exported whatever their folder (also
hidden folders and Sync Issues); folder/date windows follow the scope rules above; ids that
are copies of a selected message are merged with it. They are read from the server: if any of them
cannot be read (deleted or moved meanwhile, or still throttled), the export fails, writes nothing and
says which ones and what to do. Anything else that cannot be exported (a body, an attachment, an
attachment listing) is marked in place, and the result's `error_summary` (the file header's "Export
errors" line) says how many and why. For exports with attachments, inline images are included
when their content id is found in the rendered body. If the content id or rendered body cannot be read,
the image is included rather than silently dropped. A file attachment over the connector's 150 MB
limit gets a specific export error and should be downloaded directly from Outlook.

In TXT the mark is a block:

```text
[EXPORT ERROR] The body of this message could not be fetched.
  Step:   fetching message bodies
  Error:  HTTP 429 TooManyRequests, request-id <id>
  Likely: Microsoft throttled the mailbox (about 4 parallel requests or 10,000 per 10 minutes); the message itself is fine
  Fix:    export it again in a few minutes
```

In JSONL it is an `export_error` object (step, status, code, message, request id, likely cause,
`retry`, fix) on the message record or on the failed attachment record (`attachments_export_error`
when the attachments could not be listed). `get_conversation` sets `export_error` on each message whose body could not be
fetched; callers can count those markers without reading the body text.

**Throttling.** Microsoft Graph allows about 4 concurrent requests and 10,000 requests per 10 minutes
per mailbox; each item of a `$batch` (at most 20) counts. The connector keeps at most 4 requests and
2 batches in flight, re-sends throttled or temporarily failing batch items (429, 502, 503, 504) in
new batches of at most 20 after the advised delay, and reports items that still fail instead of
failing the whole call. Writes are never re-sent. Agents are told to avoid parallel tool calls and to
prefer one folder/date-window export over many small calls. Errors name the operation, the Graph error code and
message, and the request id; "not found" leaves the cause open (deleted, moved out of reach, or a
wrong id). A rejected token (401) is renewed
once before the connector asks you to sign in again; "access denied" (403) never asks for a sign-in.

**You (local UI).** Start it when you need it. It opens a browser tab; an open tab keeps it running, and
it stops 30 minutes after the last one is closed:

```bash
outlook-connector ui
```

The UI opens on the Inbox. Browse folders or recent mail grouped by conversation (a conversation with one
message is a plain row; a real conversation shows its message count across all folders, newest message on
top), search the mailbox, filter what is loaded, read messages (the whole body, however long, in one
request) and download their attachments, tick
conversations or single messages, and export them as one `.txt` or `.zip` (one file per conversation, one for
everything, or one per message; attachments optional). "export this view" exports the whole current
folder and date window. Deleted Items and Junk are left out unless you tick "Deleted / Junk"
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

The web reader's JavaScript tests use Node's built-in runner (Node 18+, no package dependencies) on
the shipped frontend with a minimal test DOM:

```bash
node --test tests/surfaces/message_reader.test.cjs
```

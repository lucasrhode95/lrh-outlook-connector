# Outlook for MCP clients

A local MCP server and browser UI for one user’s Exchange Online mailbox.

- Find messages by subject and contents, then read messages and full conversations.
- Export conversations, messages, or a folder/date window with attachments.
- Draft and send email, manage mailbox rules, and organize messages.
- Browse folders, read messages, and download attachments in the local UI.

![Local MCP clients and the Outlook connector, with a screenshot of the browser UI using synthetic mailbox data.](docs/connector-overview.svg)

## Requirements

- Install [uv](https://docs.astral.sh/uv/getting-started/installation/).
- Use an Exchange Online mailbox.

A separate Python installation is optional: uv selects an installed Python 3.12 or newer, or installs a managed version when needed.

## Setup

```bash
uv sync --locked
uv run outlook-connector auth
uv run outlook-connector status
```

Any Outlook, Teams, or SharePoint connector can start the shared sign-in bootstrap. It silently reuses
accepted profile credentials and asks Microsoft only for profiles that still need approval. The
shared catalog includes `outlook.graph` for Graph mail reads and `outlook.api` for Outlook Web
operations, plus the Teams profiles used by the other connectors. Microsoft may require separate
device-code and MFA approvals for different clients or resource audiences.

Core mail reads, search, exports and attachment downloads use `outlook.graph`. The `outlook.api`
profile is also required for inbox-rule and native-signature reads, as well as to draft, send, change
mailbox state or manage signature settings. `outlook.graph` requests Graph `Mail.Read` for mailbox
data and `User.Read` for the signed-in profile and photo.

The encrypted shared cache lives under `~/.lrh-m365-auth` by default; set `M365_AUTH_HOME` to use a
different shared location. It holds one active Microsoft account for all three connectors. To switch
accounts, sign out of the shared cache and sign in again:

```bash
uv run outlook-connector auth --sign-out
uv run outlook-connector auth
```

`uv run outlook-connector status` reads local profile state without network requests. It reports
cached client credentials and any matching resource access-token expiry. Profiles sharing one
client can have the same credential state even when only one resource token is cached.
`status --check` attempts token acquisition for every profile; it does not verify that each
Microsoft service or mailbox operation is authorized.

> Credentials are entered only on Microsoft's sign-in page. The application never sees them; tokens
> stay in the shared encrypted cache. If encryption is unavailable, authentication fails closed.

### Configure a local MCP client

Replace `/path/to/lrh-outlook-connector` with the absolute path to your clone.

**Claude Code** — add the server from a terminal:

```bash
claude mcp add --transport stdio lrh-outlook-connector -- uv run --directory /path/to/lrh-outlook-connector outlook-connector mcp
```

See the [Claude Code MCP guide](https://docs.anthropic.com/en/docs/claude-code/mcp).

**Codex CLI** — register the same stdio command:

```bash
codex mcp add lrh-outlook-connector -- uv run --directory /path/to/lrh-outlook-connector outlook-connector mcp
```

See the [Codex MCP guide](https://developers.openai.com/codex/mcp/).

**GitHub Copilot Chat in VS Code** — add this server to `.vscode/mcp.json` in your workspace:

```json
{
  "servers": {
    "lrh-outlook-connector": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/path/to/lrh-outlook-connector",
        "outlook-connector",
        "mcp"
      ]
    }
  }
}
```

Replace the sample path with your clone’s absolute path. In Windows JSON paths, escape backslashes (for example, `C:\\Users\\you\\lrh-outlook-connector`). This example requires Copilot Chat’s Agent mode with local MCP support; see [VS Code’s MCP guide](https://code.visualstudio.com/docs/copilot/chat/mcp-servers).

> **Client compatibility:** These examples require a locally running agentic client that can launch a stdio MCP server. Hosted ChatGPT Web, Claude web chat, Copilot web chat, and similar services cannot start this local process. This project does not provide a remote MCP endpoint; hosted clients require one to be deployed separately.

The UI opens in your browser and stops after 30 minutes without activity:

```bash
uv run outlook-connector ui
```

![150 percent screenshot of the local Outlook connector UI inside a macOS-style browser window. Mailbox content is synthetic.](docs/ui-preview.svg)

> The web UI is intended for manual message exports. The preview is plain text, not rendered HTML; it shows how extracted message content will appear in an export.

## MCP tools

### Find and read

- `auth_status` — Read cached client credentials, resource-token expiry and the sign-in command.
- `list_folders` — List reachable mail folders and their counts.
- `list_messages` — Read a page of recent messages, newest first; filter by folder, dates, and scope.
- `search_messages` — Search mail and group hits by conversation.
- `search_calendar_events` — Search Exchange calendar events online, with bounded paging and optional start-date bounds.
- `get_calendar_event` — Hydrate one event's agenda, attendees and attachment metadata from Outlook.
- `get_conversation` — Read a whole conversation across folders.
- `get_message` — Read one message body, including quoted history by default, and its attachment metadata.

### Attachments and exports

- `list_attachments` — List a message's attachment metadata, including inline image IDs.
- `download_attachment` — Save an attachment locally; downloads are limited to 150 MB.
- `save_message_mime` — Save the original message as an `.eml` file.
- `export_messages` — Export conversations, messages, or a folder/date window as TXT or JSONL. Quoted history and attachment files are included by default; use `body="unique"` to omit quoted history or `include_attachments=false` to omit attachment files. Conversations follow `scope.deleted_items`; sent/meeting filters only narrow folder/date windows. Explicit message ids are authoritative.

### Rules and signatures

- `list_rules` — Read inbox rules, including unsupported rules marked read-only.
- `create_rule` — Propose or confirm a supported inbox rule.
- `update_rule` — Propose or confirm edits to a supported inbox rule.
- `reorder_rules` — Propose or confirm the order of supported inbox rules.
- `delete_rule` — Propose or confirm deletion of a supported inbox rule.
- `list_signatures` — List native Outlook signatures and defaults.
- `get_signature` — Read one native signature's HTML and text.
- `create_signature` — Create a native Outlook signature from passive HTML; quote data-image `src` attributes.
- `update_signature` — Replace a native signature's contents.
- `delete_signature` — Delete a native Outlook signature.
- `set_default_signature` — Set or clear the new-message and reply defaults.

### Drafts and mailbox changes

- `create_draft` — Save a new message or reply as a draft.
- `send_draft` — Send an existing draft after an explicit user request.
- `set_read_state` — Mark selected messages or conversations read or unread.
- `set_flag` — Flag or unflag selected messages.
- `move_messages` — Move selected messages to a folder.
- `delete_messages` — Move selected messages to Deleted Items.

With no native signature settings or default, drafts are unsigned. If Outlook returns no account scope, configure a native signature in Outlook before using signature-write tools.

Rule writes first return a proposal; repeat the exact change only after a person approves it. Drafts are composed once; a replacement reply keeps the same `reply_to_message_id` and `reply_all` choice. Send or mailbox writes are sent once. Deletes move items to Deleted Items; nothing is permanently deleted.

## Product boundaries and calendar support

Outlook owns email and its attachments. It also owns Exchange calendar-event discovery and
hydration: schedule and recurrence, agenda, organizer, invitees, event attachment metadata, and a
Teams join URL when present in the event body. Event attachments are listed, but direct event-
attachment download is not currently exposed. Search result totals are omitted because recurring
occurrence totals were inconsistent in the tested Search v2 responses.

A calendar event linked to a Teams meeting has resources in both products. Outlook supplies the
event and verified meeting references; `lrh-teams-connector` owns the meeting chat, recordings,
transcripts, recap, and meeting notes. An agent can find the event in Outlook, then request its
related meeting resources from Teams. A join link for another meeting platform does not make the
event a Teams meeting.

The planned `lrh-sharepoint-connector` owns general OneDrive and SharePoint file search and selected
file retrieval. A file can also appear as an email attachment or a Teams message link: its storage
location does not transfer ownership of the email or message relationship to SharePoint.

## Development

```bash
uv run ruff check .
uv run ruff format --check .
uv run pyright
uv run pytest
node --test tests/surfaces/message_reader.test.cjs
```

Project references: [requirements](docs/outlook-requirements-v4.md) · [architecture](docs/architecture.md) · [API research](docs/outlook-api-research.md) · [open roadmap](docs/outlook-roadmap.md).

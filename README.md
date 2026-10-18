![Local MCP clients and the Outlook connector, with a screenshot of the browser UI using synthetic mailbox data.](https://raw.githubusercontent.com/lucasrhode95/lrh-outlook-connector/main/docs/connector-overview.svg)

A local MCP server and browser UI for one user’s Exchange Online mailbox.

- Find messages by subject and contents, then read messages and full conversations.
- Export conversations, messages, or a folder/date window with attachments.
- Draft and send email, manage mailbox rules, and organize messages.
- Browse folders, read messages, and download attachments in the local UI.

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

Run the `auth` command above and it logs you in. All `lrh-*` Microsoft 365 connectors share one login engine (`lrh-m365-auth`), so one login covers them all. Microsoft may ask for more than one approval, because the connectors use different Microsoft clients.

`status` shows the signed-in account and its cached profiles without touching the network. To switch accounts, run `auth --sign-out`, then `auth` again. Which profile serves which tool is in [architecture §6](docs/architecture.md#6-capability-routing-and-portability).

> Credentials are entered only on Microsoft's sign-in page. Tokens stay in the shared encrypted cache
> (`%USERPROFILE%/.lrh-m365-auth` on Windows, the XDG data folder elsewhere, or `M365_AUTH_HOME`).
> If encryption is unavailable, sign-in fails closed.

Outlook's data (the folder cache, exports, downloads and logs) is stored under
`%USERPROFILE%\.lrh-outlook-connector` on Windows, or `~/.local/share/lrh-outlook-connector` on other
systems. Set `OUTLOOK_CONNECTOR_HOME` to override it.

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

> **Client compatibility:** These examples require a locally running agentic client that can launch a stdio MCP server. Hosted chat apps cannot start this local process, and this project provides no remote MCP endpoint. ChatGPT can still reach it through OpenAI's [Secure MCP Tunnel](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels), where your plan and workspace allow custom MCP apps; files the server saves stay on the computer running the tunnel.

The UI opens in your browser and stops after 30 minutes without activity:

```bash
uv run outlook-connector ui
```

![150 percent screenshot of the local Outlook connector UI inside a macOS-style browser window. Mailbox content is synthetic.](https://raw.githubusercontent.com/lucasrhode95/lrh-outlook-connector/main/docs/ui-preview.svg)

> The web UI is intended for manual message exports; its export options are behind the cog. The preview is plain text, not rendered HTML; it shows how extracted message content will appear in an export.

## MCP tools

The MCP server exposes the tools below. Each one's inputs, results and caveats are in [architecture §8](docs/architecture.md#8-surfaces-tools-and-endpoints).

| MCP tool | Service module | Behavior |
|---|---|---|
| `auth_status` | shared auth provider | Read cached account/client credentials and resource-token expiry |
| `list_folders` | [`service/mailbox.py`](src/outlook_connector/service/mailbox.py) | List reachable folders and counts |
| `list_messages` | [`service/mailbox.py`](src/outlook_connector/service/mailbox.py) | List messages in a folder or mailbox, with folder, date-window and scope filters |
| `search_messages` | [`service/mailbox.py`](src/outlook_connector/service/mailbox.py) | Search messages with folder, date-window and scope filters |
| `search_calendar_events` | [`service/calendar.py`](src/outlook_connector/service/calendar.py) | Search Exchange events online across all dates, with bounded continuation |
| `list_calendar_events` | [`service/calendar_listing.py`](src/outlook_connector/service/calendar_listing.py) | List the events overlapping a window in time order, Graph first; `contains` and descriptions need Graph |
| `get_calendar_event` | [`service/calendar.py`](src/outlook_connector/service/calendar.py) | Hydrate an event's agenda, attendees, join URL and attachment metadata |
| `get_conversation` | [`service/conversations.py`](src/outlook_connector/service/conversations.py) | Read a conversation using the requested scope |
| `get_message` | [`service/mailbox.py`](src/outlook_connector/service/mailbox.py) | Read a message body |
| `list_attachments` | [`service/files.py`](src/outlook_connector/service/files.py) | List all attachment metadata: file names, sizes, types and inline image IDs |
| `download_attachment` | [`service/files.py`](src/outlook_connector/service/files.py) | Save a message's files locally by their listed names, also used in text markers |
| `save_message_mime` | [`service/files.py`](src/outlook_connector/service/files.py) | Save the original message as an .eml file |
| `export_messages` | [`service/export/`](src/outlook_connector/service/export/) | Export conversations, messages or a folder/date window |
| `list_rules`, `create_rule`, `update_rule`, `reorder_rules`, `delete_rule` | [`service/rules.py`](src/outlook_connector/service/rules.py) | Read rules or propose and confirm supported changes |
| `list_signatures`, `get_signature`, `create_signature`, `update_signature`, `delete_signature`, `set_default_signature` | [`service/signatures.py`](src/outlook_connector/service/signatures.py) | Read and manage native Outlook signatures |
| `create_draft`, `send_draft` | [`service/writes.py`](src/outlook_connector/service/writes.py) | Create drafts and send an existing draft |
| `set_read_state`, `set_flag`, `move_messages`, `delete_messages` | [`service/mutations.py`](src/outlook_connector/service/mutations.py) | Change read state, flags and message locations |

## Microsoft 365 connectors

This repository is part of the `lrh-*` **Microsoft 365** connectors: local tools that give MCP agents
access to a common corporate environment, its mail and calendar, chats and meetings, and files. Each
repository owns one product, and all of them share one sign-in.

<table>
  <tr><th width="270" nowrap>Repository</th><th>Responsible for</th></tr>
  <tr><td nowrap><picture><img src="https://raw.githubusercontent.com/lucasrhode95/lrh-m365-auth/main/docs/logos/entra-id.svg" height="29" align="absmiddle" alt="Entra ID"></picture>&nbsp;<a href="https://github.com/lucasrhode95/lrh-m365-auth">lrh-m365-auth</a></td><td>Shared <b>Microsoft</b> sign-in through <b>Entra ID</b>: one encrypted token cache and the profiles every connector uses</td></tr>
  <tr><td nowrap><picture><img src="https://raw.githubusercontent.com/lucasrhode95/lrh-m365-auth/main/docs/logos/outlook.svg" height="24" align="absmiddle" alt="Outlook"></picture>&nbsp;<a href="https://github.com/lucasrhode95/lrh-outlook-connector">lrh-outlook-connector</a></td><td><b>Exchange</b> mail, folders, drafts, rules, native signatures, and calendar events (schedule, agenda, attendees, <b>Teams</b> join link)</td></tr>
  <tr><td nowrap><picture><img src="https://raw.githubusercontent.com/lucasrhode95/lrh-m365-auth/main/docs/logos/teams.svg" height="24" align="absmiddle" alt="Teams"></picture>&nbsp;<a href="https://github.com/lucasrhode95/lrh-teams-connector">lrh-teams-connector</a></td><td><b>Teams</b> chats and channels, message attachments, and meeting chats, recordings, transcripts, recaps and notes</td></tr>
  <tr><td nowrap><picture><img src="https://raw.githubusercontent.com/lucasrhode95/lrh-m365-auth/main/docs/logos/sharepoint.svg" height="24" align="absmiddle" alt="SharePoint"></picture>&nbsp;<a href="https://github.com/lucasrhode95/lrh-sharepoint-connector">lrh-sharepoint-connector</a></td><td>General <b>OneDrive</b> and <b>SharePoint</b> file search and selected file retrieval (read-only)</td></tr>
</table>

A calendar event stays with **Outlook** even when it links to a **Teams** meeting; **Teams** owns that
meeting's chat and artifacts. A file linked from an email or a **Teams** message stays part of that
message, while the stored file itself is retrieved through **SharePoint**. Only verified identifiers
connect records across products.

## Development

```bash
uv run ruff check .
uv run ruff format --check .
uv run pyright
uv run pytest
node --test tests/surfaces/message_reader.test.cjs
```

`lrh-m365-auth` follows its `main` branch, and `uv.lock` records the exact commit. Run
`uv lock --upgrade-package lrh-m365-auth` to move to its latest commit.

Logs: each `mcp` and `ui` process writes a daily-rotated file under `<data directory>/logs/`, kept for seven days; the UI also logs to stdout. Set `OUTLOOK_CONNECTOR_LOG_LEVEL=DEBUG` for request timings ([architecture §7](docs/architecture.md#7-data-and-identity)).

Project references: [documentation index](docs/README.md) · [requirements](docs/requirements.md) · [architecture](docs/architecture.md) · [API research](docs/api-research.md) · [open roadmap](docs/roadmap.md).

## Disclaimer

> [!WARNING]
> Unofficial: these tools sign in as Microsoft's own apps and use undocumented services. Check your
> organization's AI and data policies before giving an agent your work data. Provided as is, without
> warranty. Microsoft, Outlook, Teams, SharePoint and Entra are Microsoft trademarks; this project is
> not affiliated with Microsoft.

## License

Copyright 2026 Lucas Rhode. This is source-available software, not open source: the code is public
so anyone can read and audit it. The [PolyForm Internal Use License 1.0.0](LICENSE.md) lets you and
your company use and adapt the software for your own internal operations, including at work. The
licensor also permits personal, noncommercial use. Serving other companies with it is not internal
use: you may not sell it, ship it in a product, or offer it as a service, and you may not distribute
the software or a modified copy of it.

# Outlook Connector — Requirements v4

**Evidence:** [Outlook API research](api-research.md) · **Contract per tool:** [architecture §8](architecture.md#8-surfaces-tools-and-endpoints) · **Open work:** [roadmap](roadmap.md)

This document records what the connector must do and the decisions behind it. How each tool behaves in detail is in architecture §8; open decisions are in the roadmap. Do not resolve an open decision by assumption.

---

## 1. Product

A local application for **one user's own Exchange Online mailbox**. It has two surfaces of equal importance:

- **Local UI:** you find mail and extract it as AI-friendly text bundles.
- **MCP server:** agents search, read, export and organize the same mail.

**Guiding principle: remote first.** Delegate as much as possible to Outlook's server APIs (listing, filtering, search, conversation grouping). Keep locally only what is genuinely required: account binding, folder cache, export assembly.

Both surfaces are thin adapters over one shared service, and neither may limit the other. Outlook, Teams and SharePoint are separately installed products that share one sign-in (`lrh-m365-auth`); product adapters and domain decisions stay in their owning connector.

**Ownership (2026-10-12).** Outlook owns Exchange mail, attachments and calendar events: schedule, recurrence, agenda, organizer, invitees, event attachment metadata and any join link. Teams owns a Teams meeting's chat, recordings, transcripts, recap and notes. Only verified identifiers link an event to those resources. A meeting on another platform does not become a Teams meeting. General OneDrive and SharePoint file search belongs to the SharePoint connector.

## 2. Use cases

### Agentic (MCP)

| # | Use case | Capabilities used |
|---|---|---|
| A1 | "Search my whole mailbox for information on topic X." The agent tries several keywords, then pulls the full conversation of each match. | `search_messages` (repeated) → `get_conversation` |
| A2 | "Summarize attention points from last week's email." | `list_messages(since=…)` across the mailbox → `get_conversation` per relevant message |
| A3 | "Explain what the Teams 'Example Project Chat' is talking about." The agent reads Teams, decides it needs more context, searches email and reads attachments. | Teams MCP + this MCP side by side in the client. **No cross-repo integration needed.** |
| A4 | "Delete all marketing email from last week." | `list_messages`/`search_messages` → `delete_messages` |
| A5 | "Move inbound items to their project folders (Project Alpha, Project Beta…). If unsure, don't move; list them for me." | `list_folders`, `list_messages`, `get_message`; the agent classifies, then calls `move_messages` per target |
| A6 | Agent marks messages read/unread or flags them as part of triage. | `set_read_state`, `set_flag` |
| A7 | Agent sends an email on explicit request. | `create_draft` → `send_draft` |
| A8 | "Find the Exchange event for a meeting, inspect its agenda and invitees, then decide whether to look up its related Teams resources." | `search_calendar_events` → `get_calendar_event`; then Teams MCP |
| A9 | "What do I have between these dates, and which of those meetings mention this topic in their description?" | `list_calendar_events` (`since`, `until`, `contains`, `include_body`) |

### Manual (local UI)

| # | Use case |
|---|---|
| U1 | Browse messages **grouped by conversation**. Select whole conversations and/or individual messages, choose export options, and get **one download**. |
| U2 | Search messages by subject and participants, and by content through online search. Instantly filter what's already loaded. |

## 3. Non-goals

- Shared/delegated mailboxes, even though `Mail.Read.Shared` is granted.
- Exchange Online / In-Place Archive: a **separate archive mailbox**. The account reports `HasArchive=false`, and Graph does not support it. This is *not* the normal **Archive folder** (the target of Outlook's Archive button), which is in the primary mailbox and fully in scope (Graph alias `archive`).
- **Hard delete or purge**, ever. "Delete" always means moving to Deleted Items.
- A permanent background sync daemon.
- Local full-text content search.
- Server-side text extraction from attachments. Agents receive the raw file.
- Branch-aware conversation reconstruction (parked in the roadmap).

## 4. Authentication (decided)

No new Entra registration exists or will be made. The shared `lrh-m365-auth` package owns
MSAL `PublicClientApplication` instances and one encrypted token cache used by Outlook, Teams and
SharePoint. Its profile catalog includes these Outlook client/resource pairs:

| Capability | Client / profile | Backend | Status |
|---|---|---|---|
| Core mail-data reads: folders, messages, conversations, search, attachment/MIME downloads; profile and photo | `outlook.graph` · Outlook Mobile `27922004-5251-4030-b22d-91ecd9a37ea4` | Microsoft Graph (`Mail.Read`; `User.Read` for profile/photo) | Live through shared silent auth; fresh-account consent not tested |
| Calendar event discovery and hydration | `m365.search` · One Outlook Web; `outlook.api` · One Outlook Web | Search v2 `Event` / `Exchange`; OWS `GetItem` | Live connector check 2026-10-12 (research §3.16); fresh-account consent not tested |
| Calendar windows | `teams.graph` (carries `Calendars.Read`) · One Teams client; `outlook.api` | Graph `/me/calendarView`; OWS `GetCalendarView` (fallback) | Live 2026-10-12 (research §3.18); fresh-account consent not tested |
| Mail changes: create/send drafts, read state, flags, move and soft delete | `outlook.api` · One Outlook Web `9199bf20-a13f-4107-85dc-02114787ef48` | Outlook Web Service (OWS) at `/owa/service.svc` | Send and every mutation proven (research §4.2) |
| Inbox rules | `outlook.api` | OWS `GetInboxRule` and rule actions | Proven (research §4.4) |
| Native roaming signatures | `outlook.api` | Outlook Cloud Settings | Proven (research §§4.7–4.8) |

Graph mail write/send scopes are denied (`AADSTS65002`) to the tested clients (research §2). The
`outlook.graph` profile also lacks `MailboxSettings.*`, so inbox-rule access uses OWS. Denied scope
pairs are never requested again.

- Never use browser tokens, cookies, storage, canaries or profiles.
- Encrypted persistence through `lrh-m365-auth` by default. If secure storage is unavailable, **fail closed**.
- `--unsecure` is an explicit development mode: a separate plaintext cache file outside the repository, with a warning on every use. It is never an implicit fallback.
- Interactive sign-in happens only through `outlook-connector auth`, which bootstraps every shared profile. UI and MCP calls use silent auth and return an actionable "sign-in required" error.
- Cross-process cache locking and account-generation checks protect concurrent processes and invalidate remembered credentials after sign-out or an account switch.
- Offline status separates cached client credentials from resource access-token expiry. `status --check` attempts token acquisition per profile but does not prove service authorization.

## 5. Account binding and identity

- The SQLite database is bound to a privacy-preserving fingerprint of stable token claims (`tid` + `oid`), never a display name or email address. If a different or unknown owner appears, keep the existing DB untouched and use a separate one.
- Item key: `(account fingerprint, mailbox id, Graph immutable ID)`. Every Graph call sends `Prefer: IdType="ImmutableId"`.
- Also store, as metadata only and never as the primary key: `internetMessageId`, `conversationId`, `parentFolderId`, `changeKey`.
- Writes address items by the same Graph immutable id, converted for OWS by one helper (research §4.2). Ids survive moves.

## 6. Backend

- **Prefer documented Graph when the endpoint and granted profile support the capability; use Outlook-specific services for gaps.** Here Graph handles core mail-data reads and post-write checks; OWS handles mail changes and inbox rules; Outlook Cloud Settings handles native signatures. This split is tenant-specific; [architecture §6](architecture.md#6-capability-routing-and-portability) describes the routing and how to re-route.
- One backend per capability. A failed or ambiguous write is never retried, and never retried through a different backend.

## 7. Data model and local cache

- **Folders** are cached for 10 minutes. An older cache waits for a refresh (about 23 folders in under a second).
- **Message metadata is not mirrored or cached.** Every list, search and conversation call reads the server. Decided by R1: 25.6k items, two thirds of them Junk, and a full mirror takes about 12 minutes (research §3.2).
- **Bodies and attachments** are fetched only on read or export. Attachment bytes are never cached.
- **No local retention** (decided 2026-10-02): mail deleted on the server is gone here too.
- Every message exposes one `received_at` timestamp: Graph's `receivedDateTime`, otherwise `sentDateTime`.
- Development DBs can be reset freely. No migrations.

## 8. Listing, conversations and search

- **One scope object** (`sent_items`, `meeting_mail`, `deleted_items`) is shared by list, search, conversations, exports and read-state changes. Defaults show sent and meeting mail and hide Deleted Items and Junk. Results count what was left out. The full rule is in [architecture §8.1](architecture.md#81-shared-rules).
- **Out of reach:** hidden folders, Sync Issues and non-mail items are never listed, searched or selected by folder, date window or conversation. **A message id named explicitly is authoritative** (decided 2026-10-13): reads, exports and the change tools act on it wherever Graph can; the change tools still refuse items outside the mail folders.
- **Copies** of one message (same Internet message id, e.g. mail sent to yourself) are shown once, naming the other folders.
- **Times without a zone are UTC**, normalized once in the service for every caller.
- **List:** folder-scoped or mailbox-wide, newest first, with an inclusive window, always from the server.
- **Conversation:** every message with one `conversationId` across all folders, copies once, oldest first, bounded with continuation. Messages you forgot to file still belong to the conversation. Graph rejects `$orderby` with the `conversationId` filter, so the connector sorts (research §3.4). A conversation is exactly Exchange's conversation.
- **Search is online only.** Graph `$search` has the same recall as Outlook's top-bar search, folds accents and supports field scoping (`subject:`, `from:`, `to:`) (research §3.3). Hits are grouped into conversations locally. No total is reported: Microsoft Search counts with a different engine and without the connector's folder rules. The UI adds an instant in-memory filter over what is loaded.
- **Calendar is a separate domain.** Window listing (`list_calendar_events`) reads Graph first and falls back to Outlook Web only when the Graph read fails; topic search across the calendar (`search_calendar_events`) uses Search v2; `get_calendar_event` hydrates one event. Neither searches Teams chats, recaps, recordings or transcripts.
- Every result reports coverage: whether more follow (a cursor), what the folder rules left out, and partial flags.

## 9. Reading content (MCP)

- `get_message` returns metadata plus a bounded full body with continuation, and attachment metadata, including inline-only attachments.
- The local UI reader shows the whole chosen body from one server read; a failure never shows part of it as complete.
- **Attachments are returned as raw files** for the agent's client to read. No server-side text extraction. Attachment and MIME downloads have a local 150 MB limit; larger files are downloaded from Outlook.
- **Attachments by name** (decided 2026-10-15, [roadmap](roadmap.md#enhancements)): the design shared with the Teams connector. One text extraction serves reads and exports; files and inline images appear in the text as markers named from the message; `download_attachment` fetches the named files of one message to local paths and fetches each attachment once; exports write each file under its marker name. Until it is built, the rules below and in [architecture §8.6–8.7](architecture.md#86-attachments-and-saved-files) describe the current behavior.
- MIME is delivered as a file, never inside JSON.
- **Read tools never change mailbox state, including read state.** They are annotated `readOnlyHint=true`.

## 10. Export

The selection is any mix of **whole conversations**, **individual messages** and a **folder and/or date window**, deduplicated by message and by copy. Scope never selects anything; scope alone is refused. At most 2,000 logical messages; a larger selection is refused with its count rather than cut. Options and defaults are in [architecture §8.7](architecture.md#87-exports): attachments on (non-inline, plus inline images the body references; no URL rewriting), one TXT per conversation, `txt` or `jsonl`, full bodies.

**Every export produces exactly one download:** a flat `.txt` only when the output is one TXT with no attachment files, otherwise one `.zip` (TXTs at the root, each TXT's attachments in a sibling folder named after it, no nested ZIPs). One shared orchestrator serves the UI and MCP.

**Export errors.** Messages selected by id are read first; if any cannot be read, the export fails and writes nothing. Anything that fails during the export (a body, an attachment download, an attachment listing) is marked in place, and the export completes. Each gap is one structured error, rendered the same way everywhere (TXT, JSONL, `get_conversation`):

```text
[EXPORT ERROR] The body of this message could not be fetched.
  Step:   fetching message bodies
  Error:  HTTP 429 TooManyRequests, request-id <id>
  Likely: Microsoft throttled the mailbox (about 4 parallel requests or 10,000 per 10 minutes); the message itself is fine
  Fix:    export it again in a few minutes
```

| Answer | Likely cause | Retry helps | Fix |
|---|---|---|---|
| 429 or 503 | Microsoft throttled the mailbox | yes | export it again in a few minutes |
| other 5xx, no response | Microsoft service or network problem | yes | retry later |
| 403 | access denied for this item (e.g. encrypted or protected) | no | retrying will not help |
| 404 | deleted or moved in Outlook during the export | no | refresh and select it again |
| connector download exceeds 150 MB | known local connector limit | no | download the attachment directly from Outlook |
| anything else | unexpected error | no | report it with the request id |

The output never contains tokens, signed URLs or authorization headers.

## 11. Writes

### 11.1 Drafts, sending and signatures

- **Agent-authored mail always enters Drafts first.** Drafts are composed once and never patched; a change creates a full replacement, verifies its read-back, then moves the old draft to Deleted Items. Never delete first. No attachment authoring or Send As.
- **Bodies:** exactly one of text (escaped, whitespace preserved) or HTML. HTML passes through except active content, markup left open at the end, and elements that change how the rest of the markup is read.
- **Read-back:** a create requires Graph read-back. Read-back retries are bounded; a failure returns the reliable saved id and never creates another draft.
- **Send:** only an existing draft, exactly as stored, after explicit human approval. The connector checks the bound account, cannot change the message while sending, and never retries an ambiguous send: it reads Sent Items when it can, otherwise reports unknown. Live tests use self-sends only.
- **Native signatures** are managed through Outlook Cloud Settings, account-bound; each write is sent once. `create_draft` inserts the native default (new mail or reply) after the body, unless a name is chosen or signatures are turned off. A missing selected signature fails without fallback. Add-in-generated, recipient-dependent signatures are out of scope.

### 11.2 Mailbox mutations

In priority order: **move to folder** (including Archive), **move to Deleted Items** (the only delete; no purge), **mark read/unread** (also per conversation), **flag**. Categories are parked (decided 2026-10-03). Creating folders is out of scope until requested.

- **Authorization:** the MCP client's allow/deny prompt is the safeguard. Write tools take explicit message ids and an explicit target, never a query, and are never auto-approved by annotation.
- **Per-item results:** `done`, `unchanged`, `not_found`, `failed` or `unknown`. Partial failure is reported, never hidden. Nothing is retried; an ambiguous result is read back.
- **Limits:** at most 100 explicit ids per call.

### 11.3 Inbox rules (W9)

List all rules in priority order, marking unsupported rules read-only. Conditions: From, Sent to, Subject contains and Subject-or-body contains. Actions: Move to folder and Stop processing more rules. Every create, update, toggle, reorder or delete requires a proposed write, explicit human confirmation bound to the account and the complete current rule state, one OWS write and a fresh read-back. Unsupported rules are never rewritten. Unknown outcomes must be checked in Outlook before repeating.

## 12. Operational rules

Logging, retries, bounds and errors are specified in [architecture §9](architecture.md#9-cross-cutting-concerns). The essentials:

- Never log mail content, addresses, query text or credentials.
- Never retry a write automatically. Reads, including each item of a Graph `$batch`, are retried on 429/502/503/504.
- "Not found" leaves the cause open: deleted, moved out of reach, or a wrong id.

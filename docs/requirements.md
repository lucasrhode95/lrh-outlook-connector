# Outlook Connector — Requirements v4

**Evidence:** [Outlook API research](api-research.md) · **Contract per tool:** [architecture §8](architecture.md#8-surfaces-tools-and-endpoints) · **Open work:** [roadmap](roadmap.md)

What the connector must do and the decisions behind it. How it does it is in the architecture; open decisions are in the roadmap. Do not resolve an open decision by assumption.

---

## 1. Product

A local application for **one user's own Exchange Online mailbox**, with two surfaces of equal importance over one shared service, neither limiting the other:

- **Local UI:** find mail and extract it as AI-friendly text bundles.
- **MCP server:** agents search, read, export and organize the same mail.

**Remote first.** Delegate listing, filtering, search and conversation grouping to Outlook's servers. Keep locally only the account binding, the folder cache and export assembly.

## 2. Ownership (2026-10-17)

The same in every `lrh-*` connector. Each owns one product, and all share one sign-in (`lrh-m365-auth`):

- **Outlook** owns Exchange mail and its attachments, and calendar events: schedule, recurrence, agenda, organizer, invitees, event attachments and any join link.
- **Teams** owns chats and channels, message attachments, and a Teams meeting's chat, recordings, transcripts, recap and notes.
- **SharePoint** owns discovery and retrieval of general OneDrive and SharePoint files.

A file linked from an email or a Teams message stays part of that message; the stored file is retrieved through SharePoint. Only verified identifiers connect records across products. A shared search endpoint or a similar name does not change ownership, and a meeting on another platform is not a Teams meeting.

## 3. Use cases

### Agentic (MCP)

| # | Use case | Capabilities used |
|---|---|---|
| A1 | "Search my whole mailbox for information on topic X." The agent tries several keywords, then pulls each match's conversation. | `search_messages` (repeated) → `get_conversation` |
| A2 | "Summarize attention points from last week's email." | `list_messages(since=…)` → `get_conversation` per relevant message |
| A3 | "Explain what the Teams 'Example Project Chat' is talking about." The agent reads Teams, then searches email and reads attachments. | Teams MCP and this MCP side by side. **No cross-repo integration needed.** |
| A4 | "Delete all marketing email from last week." | `list_messages`/`search_messages` → `delete_messages` |
| A5 | "Move inbound items to their project folders. If unsure, don't move; list them for me." | `list_folders`, `list_messages`, `get_message` → `move_messages` |
| A6 | Agent marks messages read/unread or flags them during triage. | `set_read_state`, `set_flag` |
| A7 | Agent sends an email on explicit request. | `create_draft` → `send_draft` |
| A8 | "Find the Exchange event for a meeting, inspect its agenda and invitees, then decide whether to look up its Teams resources." | `search_calendar_events` → `get_calendar_event`; then Teams MCP |
| A9 | "What do I have between these dates, and which meetings mention this topic?" | `list_calendars` → `list_calendar_events` |

### Manual (local UI)

| # | Use case |
|---|---|
| U1 | Browse messages **grouped by conversation**, select conversations and messages, choose export options, and get **one download**. |
| U2 | Search by subject, participants and content (online), and instantly filter what is already loaded. |

## 4. Non-goals

- Shared or delegated mailboxes, even though `Mail.Read.Shared` is granted.
- The **In-Place Archive** mailbox (`HasArchive=false`, and Graph does not support it). The normal **Archive folder** is in scope.
- **Hard delete or purge**, ever. "Delete" means moving to Deleted Items.
- A background sync daemon, local full-text search, and server-side text extraction from attachments.
- Branch-aware conversation reconstruction (parked in the roadmap).

## 5. Authentication (decided)

No new Entra registration exists or will be made. `lrh-m365-auth` owns the MSAL clients and one encrypted token cache shared by Outlook, Teams and SharePoint. Outlook uses these profiles:

| Capability | Profile · client | Backend | Status |
|---|---|---|---|
| Mail-data reads, downloads, profile and photo | `outlook.graph` · Outlook Mobile `27922004-5251-4030-b22d-91ecd9a37ea4` | Graph (`Mail.Read`, `User.Read`) | Live |
| Calendar event search | `m365.search` · One Outlook Web | Search v2 | Live 2026-10-12 (research §6.2) |
| Calendar windows, event hydration and event files | `outlook.calendar` · One Teams client | Graph (`Calendars.Read`) | Live 2026-10-12 and 2026-11-02 (research §3.8) |
| Mail changes, inbox rules, native signatures | `outlook.api` · One Outlook Web `9199bf20-a13f-4107-85dc-02114787ef48` | OWS and Outlook Cloud Settings | Proven (research §4.2, §4.4, §5.2) |

Fresh-account consent is untested for every profile (roadmap). Graph mail write and send scopes are denied to the tested clients (`AADSTS65002`, research §2), and `outlook.graph` lacks `MailboxSettings.*`; denied pairs are never requested again.

- Never use browser tokens, cookies, storage, canaries or profiles.
- Encrypted persistence by default; if secure storage is unavailable, **fail closed**. `--unsecure` is an explicit development mode with a separate plaintext cache and a warning on every use, never an implicit fallback.
- Only `outlook-connector auth` signs in interactively, preparing every shared profile. UI and MCP calls use silent auth and return an actionable "sign-in required" error.
- Offline status separates cached credentials from access-token expiry; `status --check` acquires tokens but does not prove authorization.

## 6. Account binding and identity

- The SQLite database is bound to a fingerprint of `tid` + `oid`, never a name or address. With a different or unknown owner, the existing database is left untouched and a separate one is used.
- Messages are keyed by their Graph immutable id, which survives moves; writes convert the same id for OWS (research §4.2).

## 7. Backend

- **Documented Graph first** where the endpoint and granted profile support it; Outlook-specific services fill the gaps. The current split is tenant-specific ([architecture §6](architecture.md#6-capability-routing-and-portability)).
- One backend per capability, with no fallback to another. A write is retried, bounded and honoring `Retry-After`, only when the failure proves it was not applied (never sent, or answered 429 or 503); an ambiguous write is never resent.

## 8. Data model and local cache

- **Folders** are cached for 10 minutes; an older cache waits for a refresh.
- **Message metadata is not mirrored.** Every list, search and conversation call reads the server (decided by R1: 25.6k items, two thirds of them Junk, and a full mirror takes about 12 minutes; research §3.2).
- **Bodies and attachments** are fetched only on read or export. Downloaded files are kept a week in the data directory and reused; nothing else of a message is stored.
- **No local retention** (decided 2026-10-02): mail deleted on the server is gone here too.
- Every message exposes one `received_at`: `receivedDateTime`, otherwise `sentDateTime`.

## 9. Listing, conversations and search

- **One scope object** (`sent_items`, `meeting_mail`, `deleted_items`) for list, search, conversations, exports and read-state changes. Defaults show sent and meeting mail and hide Deleted Items and Junk, except that **conversations include Deleted Items and Junk by default** (decided 2026-11-02): `get_conversation`, conversation sizes, search hits' `message_count` and an export's conversation selections. Results count what was left out. The full rule is [architecture §8.1](architecture.md#81-shared-rules).
- **Out of reach:** hidden folders, Sync Issues and non-mail items are never listed, searched or selected by folder, window or conversation. **An explicit message id is authoritative** (decided 2026-10-13): reads, exports and changes act on it wherever Graph can, though changes refuse items outside the mail folders.
- **Copies** of one message are shown once, naming the other folders.
- **Times without a zone are UTC**, normalized once for every caller.
- **List:** a folder or the whole mailbox, newest first, with an inclusive window.
- **Conversation:** exactly Exchange's conversation: every message with one `conversationId`, across all folders, copies once, oldest first, with continuation.
- **Search is online only**: Graph `$search`, with Outlook's top-bar recall, accent folding and `subject:`, `from:`, `to:` (research §3.4). Hits are grouped into conversations. No total is reported, because Microsoft Search counts with a different engine and without the folder rules.
- **Calendar is a separate domain**, read through Graph only, with no Outlook Web fallback (decided 2026-11-02): calendar list, windows on any of the user's calendars, event hydration and event files. Topic search across all dates uses Search v2. `get_calendar_event` returns the Teams chat id read from Graph's join link, labeled as an undocumented format (decided 2026-11-02). Nothing searches Teams chats, recaps, recordings or transcripts.
- Every result reports coverage: whether more follow, what the folder rules left out, and partial flags.

## 10. Reading content (MCP)

- `get_message` returns a bounded body with continuation and all attachment metadata, including inline-only files. The local UI reader shows the whole body; a failure never shows part of it as complete.
- **Attachments are raw files** for the agent's client to read, with a local 150 MiB limit; larger files are downloaded from Outlook. MIME is delivered as a file, never inside JSON.
- **Attachments by name** (decided 2026-10-15): reads, downloads and exports name files the same way, safe and unique within the message. Cloud links are listed but never downloaded.
- **Exchange's text for every message** (decided 2026-11-02): message text is Exchange's own conversion, placeholders unchanged, with inline images listed after it by name and no body position. The connector does not render HTML to text. The Teams connector still places `[Image: name]` where each image sits, because Teams serves bodies only as HTML and each image is a typed `<img>` tag; both connectors share the marker, the names and `download_attachment` by name.
- **Read tools never change mailbox state**, including read state, and are annotated `readOnlyHint=true`.

## 11. Export

The selection is any mix of **conversations**, **messages** and a **folder and/or date window**, deduplicated by message and copy. Scope alone selects nothing and is refused. At most 2,000 messages; a larger selection is refused with its count rather than cut.

**Every export is exactly one download:** a `.txt` only when the output is one text file with no attachments, otherwise one `.zip` with no nested zips. One orchestrator serves the UI and MCP. Exports include every inline file a message lists; URLs are not rewritten.

**Export errors.** Messages selected by id are read first; if any cannot be read, the export fails and writes nothing. Anything failing later (a body, an attachment download or listing) is marked in place and the export completes. Each gap renders the same way everywhere (TXT, JSONL, `get_conversation`):

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
| connector download exceeds 150 MiB | known local connector limit | no | download the attachment directly from Outlook |
| anything else | unexpected error | no | report it with the request id |

The output never contains tokens, signed URLs or authorization headers.

## 12. Writes

### 12.1 Drafts, sending and signatures

- **Agent-authored mail always enters Drafts first.** Drafts are never patched: a change creates a full replacement, verifies it, then moves the old draft to Deleted Items. Never delete first. No Send As.
- **Read-back:** every create is verified through Graph. A read-back failure returns the saved id and never creates another draft.
- **Local files on drafts** (decided 2026-10-15; every file in 14 MiB slices, decided 2026-11-02): any local path; unreadable files, or files over the send limit together, fail before a draft exists. A replacement carries only the files its call lists. Each file must be confirmed by Graph; a failed upload or confirmation moves the partial draft to Deleted Items and names the file.
- **Send limit:** 35 MiB of stored size (set 2026-11-03 from a delivered and a bounced test, research §4.6). A larger draft is refused with a readable error; a draft whose size is unknown is sent, since Outlook's own limit still bounces it.
- **Bodies:** exactly one of text (escaped, whitespace kept) or HTML. **The HTML check stays a deny-list** (decided 2026-11-02): Outlook stores and sends active content unchanged, and in testing recipients' clients still submitted forms and followed `<base>` ([research §4.5](api-research.md#45-what-outlook-keeps-of-saved-html)). Harmful elements are refused by name.
- **Send:** only an existing draft, exactly as stored, after explicit human approval, from the bound account. An ambiguous send is never retried: Sent Items is read when possible, otherwise the result is unknown. Live tests use self-sends only.
- **Native signatures** are Outlook Cloud Settings, account-bound. `create_draft` adds the native default after the body unless a name is chosen or signatures are off; a missing selected signature fails without fallback. Add-in-generated, recipient-dependent signatures are out of scope.

### 12.2 Mailbox mutations

In priority order: **move to folder** (including Archive), **move to Deleted Items** (the only delete), **mark read/unread** (also per conversation), **flag**. Categories are parked (decided 2026-10-03); creating folders is out of scope until requested.

- **Authorization:** the MCP client's prompt is the safeguard. Write tools take explicit ids (at most 100) and an explicit target, never a query, and are never auto-approved by annotation.
- **Per-item results**, with partial failure reported, never hidden. An ambiguous result is read back, never resent (§7).

### 12.3 Inbox rules (W9)

List all rules in priority order, marking unsupported ones read-only. Conditions: From, Sent to, Subject contains, Subject-or-body contains. Actions: Move to folder, Stop processing more rules. Every create, update, toggle, reorder or delete needs a proposal, explicit human confirmation bound to the account and the current rules, one write and a fresh read-back. Unsupported rules are never rewritten; unknown outcomes are checked in Outlook before repeating.

## 13. Operational rules

Logging, retries, bounds and errors are in [architecture §9](architecture.md#9-cross-cutting-concerns). Never log mail content, addresses, query text or credentials. "Not found" leaves the cause open: deleted, moved out of reach, or a wrong id.

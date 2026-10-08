# Outlook Connector — Requirements v4

**Status:** Decided scope  
**Date:** 2026-10-02  
**Evidence:** [Outlook API research](outlook-api-research.md) (single research record), private exploratory browser capture reviewed outside the repository
**Work register:** [Outlook roadmap](outlook-roadmap.md)

This document records the decisions made after the authentication investigation. Anything still undecided is listed in §14. Do not resolve it by assumption.

---

## 1. Product

A local application for **one user's own Exchange Online mailbox**. It has two surfaces of equal importance:

- **Local UI:** you find mail and extract it as AI-friendly text bundles.
- **MCP server:** agents search, read, export and eventually organize the same mail.

**Guiding principle: remote first.** Delegate as much as possible to Outlook's server APIs (listing, filtering, search, conversation grouping). Keep locally only what is genuinely required: account binding, folder cache, export assembly.

Both surfaces are thin adapters over one shared service, and neither may limit the other. The project is standalone: no shared package with the Teams exporter. How it is built is in [architecture.md](architecture.md).

## 2. Use cases

These are the use cases the design must serve. "Phase" refers to §12.

### Agentic (MCP)

| # | Use case | Capabilities used | Phase |
|---|---|---|---|
| A1 | "Search my whole mailbox for information on topic X." The agent tries several keywords, then pulls the full conversation of each match. | `search_messages` (repeated) → `get_conversation` | MVP |
| A2 | "Summarize attention points from last week's email." | `list_messages(since=…)` across the mailbox → `get_conversation` per relevant message | MVP |
| A3 | "Explain what the Teams 'Example Project Chat' is talking about." The agent reads Teams, decides it needs more context, searches email and reads attachments. | Teams MCP + this MCP side by side in the client. `search_messages`, `get_conversation`, attachment resources. **No cross-repo integration needed.** | MVP |
| A4 | "Delete all marketing email from last week." | `list_messages`/`search_messages` → `move_messages(target=deleteditems)` | Mutations |
| A5 | "Move inbound items to their project folders (Project Alpha, Project Beta, Project Gamma…). If unsure, don't move; list them for me." | `list_folders`, `list_messages`, `get_message`; the agent classifies, then calls `move_messages` per target and reports the unsure items in chat | Mutations |
| A6 | Agent marks messages read/unread or flags them as part of triage. | `set_read_state`, `set_flag` | Mutations |
| A7 | Agent sends an email on explicit request. | `send_draft` | Send |

### Manual (local UI)

| # | Use case | Phase |
|---|---|---|
| U1 | Browse messages **grouped by conversation**. Select whole conversations and/or individual messages, choose export options, and get **one download**. | MVP |
| U2 | Search messages by subject and participants, and by content through online search. Instantly filter what's already loaded. | MVP |

## 3. Non-goals

- Shared/delegated mailboxes, even though `Mail.Read.Shared` is granted.
- Exchange Online / In-Place Archive: a **separate archive mailbox**. The account reports `HasArchive=false`, and Graph does not support it. This is *not* the normal **Archive folder** (the target of Outlook's Archive button). That folder is in the primary mailbox and fully in scope (Graph alias `archive`, verified HTTP 200).
- **Hard delete or purge**, ever. "Delete" always means moving to Deleted Items.
- A permanent background sync daemon.
- Local full-text content search (§8).
- Server-side text extraction from attachments. Agents receive the raw file (§9).
- Branch-aware conversation reconstruction (parked: §10.3).

## 4. Authentication (decided)

No new Entra registration exists or will be made. The app runs its own MSAL `PublicClientApplication` against Microsoft first-party public clients that research proved usable:

| Role | Client | Resource | Status |
|---|---|---|---|
| **Reads**: folders, messages, conversations, MIME, attachments, delta, search | Outlook Mobile `27922004-5251-4030-b22d-91ecd9a37ea4` | `https://graph.microsoft.com/Mail.Read` | Proven: interactive + silent cache |
| **Writes**: send, move, delete, read state, flag, categories | One Outlook Web `9199bf20-a13f-4107-85dc-02114787ef48` | `https://outlook.office.com/.default` → OWS | Send and every mutation proven (research §4.2) |

Every Graph mail write/send scope is denied (`AADSTS65002`) to both clients (research §2). Denied pairs are never requested again.

Authentication requirements:

- Never use browser tokens, cookies, storage, canaries or profiles.
- Encrypted persistence via `msal-extensions` by default. If secure storage is unavailable, **fail closed**.
- `--unsecure` is an explicit development mode: a separate plaintext cache file outside the repository, a warning on every use, never implicit.
- Interactive sign-in happens only through `outlook-connector auth [read|write]`. UI and MCP calls use silent auth and return an actionable "sign-in required" error.
- Cross-process cache locking, because several processes may run at once.
- Sign-in is per client: the read client for the MVP, the write client when send ships. A token is used only if its account matches the bound account (§5).

## 5. Account binding and identity

- The SQLite database is bound to a privacy-preserving fingerprint of stable token claims (`tid` + `oid`), never a display name or email address. If a different or unknown owner appears, keep the existing DB untouched and use a separate one.
- Item key: `(account fingerprint, mailbox id, Graph immutable ID)`. Every Graph call sends `Prefer: IdType="ImmutableId"`.
- Also store, as metadata only and never as the primary key: `internetMessageId`, `conversationId`, `parentFolderId`, `changeKey`.
- Writes address items by the same Graph immutable id, converted for OWS by one helper (research §4.2). Ids survive moves.

## 6. Backend

- **Rule: documented Graph for every capability it can serve; OWS only for gaps.** For this tenant that means Graph for all reads and OWS for all writes (research §2). The split is tenant-specific; [architecture §6](architecture.md) describes how to re-route.
- One backend per capability. A failed or ambiguous write is never retried, and never retried through a different backend.

## 7. Data model and local cache

Lazy population:

- **Folders** are cached and served from the cache immediately; a cache older than 10 minutes is refreshed in the background. A full refresh is cheap: about 23 folders in under a second.
- **Message metadata is not mirrored or cached.** It is fetched by every list/search/conversation call. Decided by R1: 25.6k items, two thirds of them Junk, and a full mirror takes about 12 minutes (research §3.2). The summary cache for instant display was removed on 2026-10-04: listing is fast enough without it.
- **Bodies and attachments** are fetched only on read or export. Attachment bytes are never cached automatically.
- **No local retention** (decided 2026-10-02): mail deleted on the server is gone here too. Reading it gives "not found". A message selected by id that is gone when the export starts fails the export with a clear message; one deleted while the export runs is marked `[EXPORT ERROR]` in the file (§10.1).
- Development DBs can be reset freely. No migrations (see AGENTS.md).

## 8. Listing, conversations and search

**Scope, shared by list, search, conversation and export:** one `scope` object has independent keys: `scope.deleted_items` (default false) includes Deleted Items and Junk Email (O4); a folder named in the request is always included, and a subfolder counts with its parent. `scope.sent_items` (default true) includes Sent Items, Drafts and Outbox; `scope.meeting_mail` (default true) includes meeting mail (invitations, RSVPs, cancellations) where scope applies. A conversation stays whole. True shows more mail and false filters more; the service rejects a non-default key that an operation cannot apply. Results count what was left out. **Hidden folders, Sync Issues (classic Outlook's conflict copies) and non-mail items are out of reach** (never listed, searched or selected by folder/date-window/conversation export); search covers mail only. A message id named explicitly is authoritative: `get_message` and export read it wherever Graph can, hidden folders and Sync Issues included, and copies still merge. Web GET requests express these keys as `sent_items`, `meeting_mail` and `deleted_items` query parameters; JSON requests use a nested `scope` object. `since`/`until` without a time zone are UTC; the service normalizes both naive and aware dates once, for every caller. **Copies** of one message (same Internet message id, e.g. mail sent to yourself) are shown once, naming the other folders.

**List** (`list_messages`): folder-scoped or **mailbox-wide** (for A2), with inclusive `since`/`until` and a count limit. An optional server total (per-folder counts) helps plan large reads; it counts copies separately and includes meeting mail even when `scope.meeting_mail=false` hides it. A mailbox-wide listing read folder by folder reports the window's Deleted Items and Junk counts (subfolders included) on its first page, from the same count batch that chooses the folders; a count that fails is never replaced by a cached total in what is reported, and one failed folder never hides the others' counts.

- Always from the server; there is no local-only listing (`refresh=false` was removed with the summary cache on 2026-10-04).

**Conversation** (`get_conversation`): every message with one `conversationId` **across all folders**, deduplicated (copies shown once) and chronological.

- Messages you forgot to move into the right folder still belong to the conversation.
- Default scope: all folders except Deleted Items and Junk; pass `scope.deleted_items=true` to include them (O4).
- The result is bounded, with continuation for long conversations.
- Graph rejects `$orderby` combined with the `conversationId` filter, so **sort client-side** (research §3.4).
- In this phase, a "conversation" is exactly Exchange's conversation. Branches are not distinguished yet (§10.3).

**Search is online only.** No local FTS content search.

- UI search for subject/participants uses the same online search, with field-scoped queries where the backend supports them (e.g. Graph `subject:`, `from:`, `to:`).
- The UI also offers an instant in-memory filter over messages and conversations already loaded.
- Retained mail that was deleted remotely is reachable through list, conversation and filter, but not through search. Search results must say so.

**Search backend: Graph `$search`**, which has the same recall as Outlook's own top-bar search, folds accents and supports field scoping (research §3.3). Results are messages, grouped into conversations locally. No search total is reported: Microsoft Search (`/search/query`) counts with a different engine and without the connector's folder rules, so its number cannot be compared with the results.

Every search result reports coverage: whether more results follow (a cursor), what the folder rules left out, and partial flags.

## 9. Reading content (MCP)

- `get_message` returns metadata plus a bounded body with an `offset`/`max_chars` continuation. Text is the default. HTML and `uniqueBody` are available on request.
- The local UI reader shows the whole chosen body, however long, from one request (one server read); a failure says the complete message could not be loaded and never shows part of it as complete.
- **Attachments are returned as raw files** (an MCP resource or file artifact) for the agent's client to read. No server-side text extraction. Attachment types: file, inline, item and reference. Item and reference attachments are listed with their metadata. Fetching them is best effort.
- MIME (`$value`) is delivered as a file or resource, never inside JSON.
- Read tools are annotated `readOnlyHint=true`, `destructiveHint=false`. **Read tools never change mailbox state, including read state.**

## 10. Export

### 10.1 Selection and options (UI and MCP)

The selection is any mix of **whole conversations**, **individual messages** and a **folder** and/or **date window** (`since`/`until`), deduplicated by message and by copy. Scope narrows folder/date-window selections; it never selects anything and does not narrow conversations or explicit messages. Individual messages are exported wherever they are (§8); conversations and folder/date-window selections follow the scope rules. At most 2,000 messages; `limit` lowers that, and a larger selection is refused with its count rather than cut.

| Option | Default | Effect |
|---|---|---|
| `include_attachments` | off | **On:** attachment files are downloaded into the ZIP, with sanitized and deduplicated names. A failed download becomes an `[EXPORT ERROR]` block and does not fail the export. Only **non-inline** attachments by default. Inline images (signatures, quoted history: 74% of file attachments) are included only when the rendered body references their `cid:`. Forwarded-mail attachments (`itemAttachment`) are saved as `.eml`. **Off:** the TXT lists non-inline attachment file names (and sizes) only. **No URL rewriting either way**, and the TXT never contains Microsoft URLs. |
| `combine` | `per_conversation` | `per_conversation`: one TXT per conversation, chronological; a selected individual message goes into its conversation's TXT. `all`: one TXT for the whole selection, chronological, with per-conversation section headers. `none`: one TXT per message. |
| `format` | `txt` | `txt` for people. `jsonl` for agents: one JSON record per message (ids, dates, folder, people, body, attachments), always one file. |
| `scope` | `{"deleted_items": false, "sent_items": true, "meeting_mail": true}` | Independently narrow folder/date-window selection (see §8). |
| `body` | `unique` | `unique` strips quoted reply history (Graph `uniqueBody`). `full` keeps it. |

TXT content:

- Headers per message: From, To, CC, date, subject and folder, the message, conversation and Internet ids, the other folders of merged copies.
- The file header says what was left out by folder, and one "Export errors: …" line counts what could not be exported, by kind and likely cause; a merged copy is named on its message (`Also in:`).
- Then the body.

**Export errors.** Messages selected by id are read from the server first. If any cannot be read, the export fails and writes no file: "not found" when they are gone, "throttled" when any is still throttled after the retries, otherwise a service error. The message counts them, names the first few with their case, and says: they may have been deleted or moved in Outlook, or Microsoft is throttling requests; refresh the list and retry; nothing was exported. Conversations and folder/date-window selections are listed from the server at export time.

Anything that fails during the export (a body, an attachment download, an attachment listing) is marked in place, and the export completes. Callers never have to scan messages to learn about errors: the export result counts them (`export_errors`, `error_summary`), and `get_conversation` sets `export_error` on each affected message and counts them in `body_errors`. Every such gap is one structured error, rendered the same way everywhere (TXT, JSONL, `get_conversation`):

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
| anything else | unexpected error | no | report it with the request id |

JSONL carries the same error as an `export_error` object (on the message when its body is missing, on a failed attachment record, or as `attachments_export_error` when the attachments could not be listed). The export result has `export_errors` (per step) and `error_summary` (the header line), and the UI shows that line after the download.
- The output must never contain tokens, signed URLs or authorization headers.

### 10.2 Output packaging rule

**Every export produces exactly one download.**

- **A flat `.txt`** only when the output is exactly one TXT and no attachment files.
- **Otherwise a single `.zip`.** TXT files go at the root. When attachments are included, each TXT's attachments go into a sibling folder named after that TXT (`<stem>/`). No nested ZIPs.

One shared orchestrator serves both the UI (HTTP download) and MCP (resource or file reference).

### 10.3 Parked: branch-aware conversations

Build the reply tree from RFC 5322 `Message-ID` / `In-Reply-To` / `References` headers, with Exchange `ConversationId` and `Thread-Index` as supporting evidence. Detect branches, including forwards, and offer a merged chronological view where each message is labeled with its branch. Research the quality of the existing headers before committing to this (R4).

## 11. Writes

### 11.1 Send (first write phase)

Native Outlook roaming signatures are managed by list_signatures, get_signature,
create_signature, update_signature, delete_signature and set_default_signature. The settings are
account-bound through the write sign-in; each write is sent once. Names are exact and case-sensitive;
commas are refused because the name list uses commas as separators. Creation and content update take
passive HTML and derive the text format. Update changes contents without renaming. New-message and
reply/forward defaults are independent and may be cleared; deleting a selected signature may leave a
dangling default. Add-in-generated, recipient-dependent signatures are outside connector behavior.

create_draft freshly reads the native default appropriate to new mail or reply/forward, inserting it
after the supplied body and before Exchange's quoted reply history. An explicit signature name
overrides that default. include_signature=false suppresses insertion. A missing or unreadable
selected/configured signature fails without fallback. Native data-URI images become inline attachments
with matching CID references; text-only signatures are escaped as body text. A replacement draft
resolves the default again, so a caller must select a named signature again if it wants the same one.

Agent-authored outgoing mail always enters Outlook Drafts first. `create_draft` accepts exactly
one of `text_body` and `html_body`. Drafts are composed once and are never patched. To change a draft,
create a replacement with the full intended content and verify its Graph read-back before moving the
old draft to Deleted Items with `delete_messages`; use the replacement id from then on. Never delete
the old draft first, and tell the user the previous version remains in Deleted Items. For a reply,
use the same `reply_to_message_id` and `reply_all` choice. The replacement has a new id; Outlook body
edits and attachments are not carried over. There is no version check against the old draft; read
it first if needed. Recipient- or subject-only changes recreate the whole draft. No attachments
authoring or Send As is supported.

Text is escaped into minimal HTML, preserving line breaks, indentation, repeated spaces/tabs and
non-breaking spaces. HTML is intentional pass-through except scripts, embedded active content,
forms, JavaScript URLs and event handlers. Both routes share one private HTML write engine.

A successful create requires Graph read-back and returns the draft id, server text and HTML, and
simple findings (empty body, missing known attachments/inline images, reply history). Read-back
retries are bounded; failure returns a failed result with the reliable saved id, never creates
another draft, and stores no recovery registry.

The agent shows/uses the server read-back and obtains explicit human approval before calling
`send_draft(draft_id)`. The connector requires an existing Microsoft draft and checks the bound
write account; it cannot reconstruct or change the message while sending. No `propose_email`,
direct `send_email`, custom confirmation token or format guessing remains.
Ambiguous sends are never retried: read the exact id in Sent Items when possible, otherwise report
unknown and ask the user to check Sent Items and Outbox. MCP sending annotations remain destructive
and open-world. Live tests require an explicit user request and use self-sends only.

### 11.2 Mailbox mutations (second write phase)

Listed in priority order:

1. **Move to folder**, including Archive and project folders.
2. **Move to Deleted Items.** This is the only "delete". No purge.
3. **Mark read/unread.**
4. **Flag** (follow-up flag). Categories are parked: not needed (decided 2026-10-03).

Requirements:

- **Authorization:** the user's MCP client allow/deny prompt is the safeguard for mutations. There is no server-side plan or confirmation token. To keep that prompt meaningful:
  - Write tools take **explicit message IDs** and an explicit target. There is no "move everything matching a query" on the server.
  - They are never auto-approved by annotation: `readOnlyHint=false`. `destructiveHint=true` for move and delete, `false` for read state and flag.
- **Per-item results:** each tool returns results and `counts`: `done`, `unchanged` (already so; nothing sent), `not_found`, `failed` (with Outlook's code) or `unknown`. Partial failure is reported, never hidden. Nothing is retried. On an ambiguous result, the items are read back: `done` where the change is visible, `unknown` elsewhere (also when the read-back fails).
- **Errors in one chunk** (changes go 20 messages at a time): with `continue_on_error=true` (default) later chunks are still attempted; with false they are not sent and report `failed` with detail `not sent`.
- **Limits:** at most 100 explicit message ids per call. Read state per conversation expands each conversation to its messages in scope (all copies), up to the 1,000 the server lists per conversation, with a note when one is cut there. When such a selection has more than 100 messages, `counts` covers all of them and `results` keeps only explicit ids and messages that did not end `done` or `unchanged`.
- **Delete** moves to Deleted Items; messages already in Deleted Items are left alone, so nothing is ever deleted permanently. **Read state** also works per conversation (every message in scope).
- **Folder targets** are resolved through `list_folders`. Creating folders is out of scope until requested.

### 11.3 Inbox rules (W9)

List all rules in priority order, marking unsupported rules read-only. First-version conditions:
From, Sent to, Subject contains and Subject-or-body contains. Actions: Move to folder and Stop
processing more rules. Partial edits keep omitted fields; null conditions clear them, so a proposal
shows only the fields the request supplied.
Every create/update/toggle/reorder/delete requires a proposed write, explicit human confirmation,
a single OWS write and fresh read-back. Confirmation binds account and complete current rule state.
Enable/disable is a separate update. Reordering refuses collections containing unsupported rules,
which must never be rewritten. Unknown outcomes must be checked in Outlook before repeating.

## 12. Phases

| Phase | Content |
|---|---|
| **MVP (read)** | Auth (read client), folders, `list_messages`, `get_conversation`, `get_message`, attachment resources, online search, export (§10), local UI, MCP read surface |
| **Send** | Write-client sign-in, `send_draft` with safeguards |
| **Mutations** | move → delete → read state → flag |
| **Parked** | Branch-aware conversations (§10.3): rely on Exchange conversations for now |

## 13. Operational rules

Logging, retries, bounds and errors are specified in [architecture §9](architecture.md). The essentials:

- Never log mail content, addresses, query text or credentials.
- Never retry a write automatically. Reads, including each item of a Graph `$batch`, are retried on 429/502/503/504; an item that still fails keeps its last status.
- "Not found" leaves the cause open: deleted, moved out of reach, or a wrong id.

## 14. Open decisions

| # | Question | How it gets decided |
|---|---|---|
| O3 | UI layout for browsing. **Proposal:** a folder picker (plus an "all mail, recent" view) listing conversations grouped by `conversationId`, expandable to individual messages; a search box (online); an instant filter over the loaded list; checkboxes for conversations and messages; export options per §10.1 | Confirm while building U1 |
| O4 | `get_conversation` default excludes Deleted Items and Junk: confirm or change | Confirm during MVP |
| O5 | Conversation-header quality for branch detection | R4, before §10.3 |

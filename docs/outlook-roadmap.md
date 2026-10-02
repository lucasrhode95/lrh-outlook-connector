# Outlook Connector Roadmap

The work register for [Requirements v4](outlook-requirements-v4.md). The build order and the modules each item touches are in [architecture §11](architecture.md). Evidence is in [API research](outlook-api-research.md).

Snapshot **2026-10-02**: the read MVP is built, tested against a fake Graph mailbox, and verified live (read-only) against the real mailbox. A hardening pass after the 90-day review (H1–H3) is built and tested against the fake mailbox; its batched paths still need a live check (V1). Send and mailbox changes are next.

**Status terms:** **Done** (exists with tests or evidence) · **Partial** (specific gap remains) · **Pending** · **Parked** (plausible, but no current need).

## Research

| Item | Status | Result |
|---|---|---|
| A0 First-party auth | **Done** | Graph reads via Outlook Mobile; OWS writes via One Outlook Web. Graph write/send denied to both clients (research §2). |
| R1 Catalog sizing | **Done** | Folders only (research §3.2) |
| R2 Search bake-off | **Done** | Graph `$search` (research §3.3) |
| R3 Conversation retrieval | **Done** | Works across folders; sort locally (research §3.4) |
| R4 Thread-header quality | **Partial** | Received mail is fine; own messages need a fallback before E3 (research §3.4) |
| R5 OWS write contracts | **Done** | Send, read, flag, categories, conversation read, move, soft delete (research §4.2) |
| S2 Delta semantics | **Done** | `@removed` → GET by id; moves by id (research §3.6) |

## MVP (read)

| Item | Status | Scope |
|---|---|---|
| A1 Token provider | **Done** (live: encrypted DPAPI cache, silent refresh) | One centralized provider, named profiles from config, MSAL + encrypted cache, `--unsecure`, cross-process lock, account check (architecture §5.1) |
| A2 CLI | **Done** | `outlook-connector auth [read\|write] [--unsecure]` and `status` |
| B1 Graph reader | **Done** | `MailReader` over Graph: folders, list, get, conversation, `$search`, attachments, MIME, `$batch`. Folder delta (S2) is not used: the folder cache refreshes in full. |
| S1 Store | **Done** | Account-bound SQLite: folder cache, retained messages, tombstones |
| S3 Reconciliation | **Done** | Remove → GET by id → tombstone only on 404. Never erase known bodies. |
| L1 `list_messages` | **Done** | Folder or mailbox-wide, `since`/`until`, limit, `refresh`, `received_only` (no Sent/Drafts/Outbox/Deleted/Junk) |
| T1 `get_thread` | **Done** | Conversation across folders, local sort, retained-deleted merge, bounded; the cursor keeps the original selection; truncation past 1,000 messages is reported |
| L2 `search_messages` | **Done** | Graph `$search`, grouped by conversation, coverage |
| E1 Export | **Done** | Requirements v4 §10: threads + messages, attachment policy, combine options, one download. Also a range selection (`since`/`until`/`folder`/`received_only`), `limit` up to 2,000, source ids per message, counts of what was left out or unavailable. |
| M1 MCP surface | **Done** (read-only tools) | Read tools and resources (architecture §8) |
| U1 Local UI | **Done** | Thread-grouped list (opens on the Inbox; real conversation sizes, one-message conversations as plain rows; newest message on top), search, in-memory filter, selection, export (v4 O3) |

## Hardening (90-day review, 2026-10-02)

| Item | Status | Scope |
|---|---|---|
| H1 Safe `$batch` | **Done** | Numbered batch request ids (Graph compares ids case-insensitively). At most 2 batches and 4 requests in flight. Throttled items re-sent in batches of ≤20 after `Retry-After`. Per-item results. (OUTLOOK-02, OUTLOOK-03; likely cause of OUTLOOK-01) |
| H2 Diagnostics | **Done** | Errors name the operation, status, Graph code and message, request id, and failed batch-item count. Throttling limits stated to clients. (OUTLOOK-01) |
| H3 Partial exports | **Done** | Unfetchable bodies marked and listed instead of failing the export |
| V1 Live check | **Pending** | Re-run the 1,500-message export against the real mailbox; check conversation sizes and batched attachment listing live |

## Send

| Item | Status | Scope |
|---|---|---|
| W1 Send | **Pending** | `MailWriter.send` via OWS `CreateItem`. Plain text, `user_confirmation`, revalidation, no retry. Needs the write sign-in. |

## Mutations

| Item | Status | Scope |
|---|---|---|
| W2 Move to folder | **Pending** | Explicit ids + target. Per-item results. Store update. |
| W3 Delete (soft) | **Pending** | `DeleteItem` `MoveToDeletedItems`. Never purge. |
| W4 Read/unread | **Pending** | Per item, and per conversation (`ApplyConversationAction`) |
| W5 Flag / categories | **Pending** | Existing categories only, unless creation is requested later |

## Later and parked

| Item | Status | Note |
|---|---|---|
| E3 Branch-aware threads | **Pending** | After R4's fallback for the user's own messages |
| X1 Shared mailboxes | **Parked** | `Mail.Read.Shared` is granted. Needs a named mailbox and a need. |
| X2 Online Archive mailbox | **Parked** | Not in Graph, and the account has none. The normal Archive folder is in scope. |
| X3 OWS reader | **Parked** | Proven (research §4.3). Only needed in the "no Graph" scenario (architecture §6.2). |
| X4 Substrate search | **Parked** | Works (research §5). Graph `$search` has equal recall. |
| X5 Local full-text search | **Parked** | Online search was chosen |
| X6 Attachment text extraction | **Parked** | Agents receive raw files |
| X7 Resumable export with progress | **Parked** | Review suggestion. Not needed while exports finish in one call with per-item gaps (H1–H3); revisit if a real export still hits limits. |
| X8 Range export in the UI | **Parked** | The UI exports selected threads/messages; range export is MCP-only for now. |
| X9 Folder delta | **Parked** | Researched (S2); a full folder refresh takes under a second. |

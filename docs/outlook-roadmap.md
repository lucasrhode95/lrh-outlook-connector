# Outlook Connector Roadmap

The work register for [Requirements v4](outlook-requirements-v4.md). The build order and the modules each item touches are in [architecture §11](architecture.md). Evidence is in [API research](outlook-api-research.md).

Snapshot **2026-10-02**: research is complete for the MVP, send and mutations. There is no product code yet.

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
| A1 Token provider | **Pending** | One centralized provider, named profiles from config, MSAL + encrypted cache, `--unsecure`, cross-process lock, account check (architecture §5.1) |
| A2 CLI | **Pending** | `outlook-connector auth [read\|write] [--unsecure]` and `status` |
| B1 Graph reader | **Pending** | `MailReader` over Graph: folders (+delta), list, get, conversation, `$search`, attachments, MIME, `$batch` |
| S1 Store | **Pending** | Account-bound SQLite: folder cache, retained messages, tombstones |
| S3 Reconciliation | **Pending** | Remove → GET by id → tombstone only on 404. Never erase known bodies. |
| L1 `list_messages` | **Pending** | Folder or mailbox-wide, `since`/`until`, limit, `refresh` |
| T1 `get_thread` | **Pending** | Conversation across folders, local sort, retained-deleted merge, bounded |
| L2 `search_messages` | **Pending** | Graph `$search`, grouped by conversation, coverage |
| E1 Export | **Pending** | Requirements v4 §10: threads + messages, attachment policy, combine options, one download |
| M1 MCP surface | **Pending** | Read tools and resources (architecture §8) |
| U1 Local UI | **Pending** | Thread-grouped list, search, in-memory filter, selection, export (v4 O3) |

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

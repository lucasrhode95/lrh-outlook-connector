# Outlook Connector Roadmap

The work register for [Requirements v4](outlook-requirements-v4.md). The build order and the modules each item touches are in [architecture §11](architecture.md). Evidence is in [API research](outlook-api-research.md).

Snapshot **2026-10-02**: the read MVP is built, tested against a fake Graph mailbox, and verified live (read-only) against the real mailbox. Two hardening passes (the 90-day review, H1–H3, and the code review, H4–H6) are built, and the live check (V1) passed against the real mailbox. Of its findings, H8 and H9 are done; **next: H7 (count-guided mailbox-wide listing) and H10**, then a draft-first write path (W0) before send and the mutations, lowest risk first.

**Status terms:** **Done** (exists with tests or evidence) · **Partial** (specific gap remains) · **Pending** · **Parked** (plausible, but no current need).

## Research

| Item | Status | Result |
|---|---|---|
| A0 First-party auth | **Done** | Graph reads via Outlook Mobile; OWS writes via One Outlook Web. Graph write/send denied to both clients (research §2). |
| R1 Catalog sizing | **Done** | Folders only (research §3.2) |
| R2 Search bake-off | **Done** | Graph `$search` (research §3.3) |
| R3 Conversation retrieval | **Done** | Works across folders; sort locally (research §3.4) |
| R4 Thread-header quality | **Parked** | Received mail is fine; own messages would need a fallback. Only needed for E3, which is parked (research §3.4) |
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
| L1 `list_messages` | **Done** | Folder or mailbox-wide, `since`/`until`, limit, `refresh`; shared scope rules (`include_deleted_items`, `received_only`); `include_total`; compact by default for MCP |
| T1 `get_thread` | **Done** | Conversation across folders, local sort, retained-deleted merge, bounded; the cursor keeps the original selection; truncation past 1,000 messages is reported |
| L2 `search_messages` | **Done** | Graph `$search`, grouped by conversation with each conversation's message count, exact date bounds, coverage |
| E1 Export | **Done** | Requirements v4 §10: threads + messages, attachment policy, combine options, one download. Also a range selection (`since`/`until`/`folder`/`received_only`), `limit` up to 2,000, source ids per message, counts of what was left out, merged or unavailable. `format=jsonl` for agents. |
| M1 MCP surface | **Done** (read-only tools) | Read tools; files returned as local paths (architecture §8) |
| U1 Local UI | **Done** | Thread-grouped list (opens on the Inbox; real conversation sizes, one-message conversations as plain rows; newest message on top; merged copies), search, in-memory filter, selection, export and "export this view", attachment downloads, Deleted/Junk toggle (v4 O3) |

## Hardening (90-day review and live check, 2026-10-02)

| Item | Status | Scope |
|---|---|---|
| H1 Safe `$batch` | **Done** | Numbered batch request ids (Graph compares ids case-insensitively). At most 2 batches and 4 requests in flight. Throttled items re-sent in batches of ≤20 after `Retry-After`. Per-item results. (OUTLOOK-02, OUTLOOK-03; likely cause of OUTLOOK-01) |
| H2 Diagnostics | **Done** | Errors name the operation, status, Graph code and message, request id, and failed batch-item count. Throttling limits stated to clients. (OUTLOOK-01) |
| H3 Partial exports | **Done** | Unfetchable bodies marked and listed instead of failing the export |
| H4 Review bugs | **Done** | Retained deleted mail on every page (lists and range exports); `get_thread` coverage; summary column without bodies; 401 renew-then-sign-in (403 is access denied); exact search dates; Junk/Deleted folder views in the UI |
| H5 Consistency | **Done** | One scope rule set for every tool (`include_deleted_items`, `received_only`, stable `excluded` keys); copies of one message merged (`also_in`); JSONL export; compact results; totals; message counts on search hits |
| H6 Tooling | **Done** | Committed `uv.lock`; pyright (standard mode) clean and in the dev group |
| V1 Live check | **Done** (2026-10-02) | Real mailbox, read-only. **Exports:** the 240 newest days (1,471 messages, the 1,500-message case) completed as JSONL (260 s), TXT (206 s) and TXT with attachments (718 s, 173 MB, 761 files): no body, listing or download failures, identical files stored once. **`include_total`:** the Inbox count equals the folder total (144); mailbox-wide counts in under 1 s. **Conversation sizes:** 76 conversations in 2.3 s, 40 compared with `get_thread`, no mismatch; search-hit counts match too. **Merged copies:** 500 listed messages, no repeated id or Internet id across pages; self-sent mail shows `also_in: Sent Items`. Findings: H7–H10. |
| H7 Junk-heavy mailbox-wide listing | **Pending — next** (approach decided 2026-10-02) | Only scopes without a folder are affected (MCP `list_messages`/`export_messages` without `folder`, including `received_only`); folder views such as the UI's Inbox are not. Graph lists the whole mailbox newest first, Junk and Deleted Items included, and the connector drops those afterwards. With 16,846 junk messages, pages of 100 kept 11–40 messages, and the 1,471-message export read 13,821 summaries, most of its run time. Graph's `parentFolderId ne` filter excludes them server-side but takes 9–14 s per page (vs ~1 s). **Decided:** count first, then list only the folders that matter: one `$batch` of per-folder counts for the window (`count_messages`, already built and used by `include_total`) picks the in-scope folders with messages, which are then listed in parallel and merged newest first; the cursor carries each folder's continuation. That usually cuts a whole-mailbox page to a handful of folders. Emptying Junk only hides the symptom until it refills. |
| H8 Mail-only search; hidden items out of reach | **Done** (2026-10-02) | `$search` also returned Teams meeting items from the hidden `SkypeSpacesData/TeamsMeetings` folder (7 of 22 hits for `from:lucas`); `get_thread` said "not found" for them. Now hidden folders, and items outside the mail folders, are out of reach everywhere: never listed, searched, counted, threaded or exported (`coverage.excluded.hidden`), `list_folders` leaves them out and naming one is refused. Search is described as mail only in the MCP instructions, tool descriptions and the UI. An unknown folder id refreshes the folder list once, so a folder created meanwhile is still found. Meeting search stays X10. |
| H9 Sync Issues in scope | **Done** (2026-10-02) | `Sync Issues` and its subfolders (`Conflicts`, `Local Failures`, `Server Failures`) are created by classic Outlook for Windows, not by Microsoft mail or this app: when two versions of one item collide during sync, Outlook keeps one and files the other in `Conflicts`. The live export held 24 such copies (17 duplicated mail you later deleted, 7 had no other copy). Now treated like Deleted Items and Junk: left out by default (`excluded.sync_issues`), included with `include_deleted_items`, listed and nameable even when Graph marks the folder hidden. Also fixed on the way: subfolders count with their parent, so a folder deleted in Outlook (it moves into Deleted Items with its mail) is left out like Deleted Items. |
| H10 Merged-copy count | **Pending** | Reporting only: the export result's `duplicates_merged` (how many extra copies of one message were folded into one) said 0, while 4 exported messages carried `also_in`. The listing merged those copies before the export counted, so the export missed them. Count copies merged at either stage. |

## Send

Draft first: an agent prepares the message and you send it from Outlook. It proves the OWS write path with nothing leaving the mailbox, and needs no confirmation protocol.

| Item | Status | Scope |
|---|---|---|
| W0 Create draft | **Pending — first write** | `MailWriter.create_draft` via OWS `CreateItem` (`SaveOnly`) into Drafts, optionally as a reply. Returns the draft id; never sends. Needs the write sign-in. |
| W1 Send | **Pending** | After W0. `MailWriter.send` via OWS `CreateItem`. Plain text, `user_confirmation`, revalidation, no retry. |

## Mutations

Ordered by risk: reversible state changes first, then moves and deletes.

| Item | Status | Scope |
|---|---|---|
| W4 Read/unread | **Pending** | Per item, and per conversation (`ApplyConversationAction`) |
| W5 Flag / categories | **Pending** | Existing categories only, unless creation is requested later |
| W2 Move to folder | **Pending** | Explicit ids + target. Per-item results. Store update. |
| W3 Delete (soft) | **Pending** | `DeleteItem` `MoveToDeletedItems`. Never purge. |

## Later and parked

| Item | Status | Note |
|---|---|---|
| E3 Branch-aware threads | **Parked** | Decided 2026-10-02: rely on Exchange's conversations (and `uniqueBody` for quoted history) instead of rebuilding reply trees. Revisit only with a concrete need; it would start with R4. |
| X1 Shared mailboxes | **Parked** | `Mail.Read.Shared` is granted. Needs a named mailbox and a need. |
| X2 Online Archive mailbox | **Parked** | Not in Graph, and the account has none. The normal Archive folder is in scope. |
| X3 OWS reader | **Parked** | Proven (research §4.3). Only needed in the "no Graph" scenario (architecture §6.2). |
| X4 Substrate search | **Parked** | Works (research §5). Graph `$search` has equal recall. |
| X5 Local full-text search | **Parked** | Online search was chosen |
| X6 Attachment text extraction | **Parked** | Agents receive raw files |
| X7 Resumable export with progress | **Parked** | Review suggestion. Not needed while exports finish in one call with per-item gaps (H1–H3); revisit if a real export still hits limits. |
| X9 Folder delta | **Parked** | Researched (S2); a full folder refresh takes under a second. |
| X10 Meeting search | **Later — nice to have** | Decided 2026-10-02: belongs here, not in lrh-teams. Meetings are calendar events in the same Exchange mailbox (Graph `/me/calendarView`, `/me/events`: subject, time, organizer, attendees, agenda, Teams join link). lrh-teams keeps meeting chats, which it already reads; the join link carries the meeting chat id (`19:meeting_…`) so an agent can hand over to lrh-teams without duplicating either side. Not the `TeamsMeetings` folder items (undocumented Teams storage). First step: probe whether the read client's token grants `Calendars.Read`. |

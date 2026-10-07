# Outlook Connector — Architecture

**Status:** Accepted (2026-10-02)  
**Inputs:** [Requirements v4](outlook-requirements-v4.md) · [Roadmap](outlook-roadmap.md) · [API research](outlook-api-research.md)

![Outlook connector architecture](architecture.svg)

The classes and who calls whom, in more detail: [code map](code-map.svg).

This document describes how the application is built: processes, layers, modules, data, and the main flows. *What* it must do is in v4. *Why* each API choice was made is in the research doc.

---

## 1. Principles

1. **Remote first.** Outlook's servers do the work: listing, filtering, search, conversation grouping. The app keeps locally only what the server cannot give back.
2. **No always-on service.** Every entry point is a short-lived local process started on demand, for one user. Several may run at once, and weeks may pass between runs. It is not designed to be hosted as a long-running MCP or HTTP server (§2).
3. **One domain implementation, thin surfaces.** The UI and MCP call the same service. Neither reimplements domain decisions.
4. **Protocol knowledge stays at the edge.** Only the `remote/` package knows URLs, Graph/OWS JSON, ID formats and paging. The service works with domain models.
5. **Documented first, gaps filled.** Use Graph for every capability it can serve. OWS (Outlook Web's private JSON RPC) fills only the gaps. For this tenant those gaps are all writes, because every Graph mail write/send scope is denied to the usable clients (research §2). The split is tenant-specific and swappable (§6).
6. **Small modules.** Each responsibility gets its own module from the start. Do not grow a 2,000-line connector or service.
7. **Current format only.** Follow AGENTS.md: no migrations or compatibility shims. Development databases are reset, not migrated.

## 2. Process model

There is no daemon. Three entry points, all short-lived:

| Command | Lifetime | Started by | Loads |
|---|---|---|---|
| `outlook-connector mcp` | One process per MCP client session (stdio). Exits when the session ends. | Claude Code / Codex, from its MCP config | core + MCP |
| `outlook-connector ui [--port]` | Starts a localhost web server and opens the browser. Exits on Ctrl+C or after an idle timeout. | You, occasionally | core + web stack |
| `outlook-connector auth [read\|write] [--unsecure]` | One-shot device-code sign-in | You, rarely | core + auth |

`outlook-connector status` (offline) shows the signed-in account, which clients have tokens, and the store location.

Consequences of short-lived processes (decided; see AGENTS.md):

- **Nothing in memory is relied on across sessions.** A continuation travels in the result: cursors carry the remote link or per-folder positions and the original options.
- **Local caches are used only while fresh.** A process often starts after days or weeks idle, so the cache it finds may be very old. The folder cache is used while younger than 10 minutes and otherwise refreshed before the call continues; a stale-while-revalidate design was dropped on 2026-10-04 because the refresh only helped the next process while the current call used a weeks-old folder tree (H7).
- **No background work outlives a call:** no schedulers, sync loops or warm-up tasks. The UI's idle timer only stops the process.
- **Concurrent processes share the local files:** the store uses WAL and short transactions, the token cache a cross-process lock, output files an exclusive create.
- **Not a hosted service:** one user, localhost only (the UI binds 127.0.0.1 with a per-run session token), no multi-user auth or remote access.

**Consequences:**

- **Concurrency is cross-process.** Two agent sessions and the UI may run at the same time. The shared resources are the **token cache** (MSAL file cache with `msal-extensions` cross-process locking) and the **SQLite store** (WAL mode, `busy_timeout`, short write transactions, a connection per operation).
- **Startup cost matters**, because every agent session pays it.
  - Nothing runs at startup: no sync, no folder walk, no token refresh before the first call.
  - The web stack (Starlette, uvicorn) is imported only by `ui`, never by `mcp`.
  - Folders come from the cache while it is younger than 10 minutes; an older cache waits for a refresh (under a second). Messages are always listed from the server (no local message cache; removed 2026-10-04).
  - Each process builds its MSAL clients once and keeps access tokens in memory until shortly before expiry (rebuilding the client costs a network round trip).
- **No in-memory state outlives a call.** MCP continuation cursors are self-contained (they encode the remote `nextLink` or offset plus the original selection). They survive a client restart.
- **Exports run in the process that asked.** A UI export is one request that returns the file; an MCP export completes inside the tool call and returns a local file path. There is no background job queue and no progress reporting. An export holds at most 2,000 messages; it fetches bodies in bounded `$batch` rounds and finishes with per-message gaps marked rather than failing as a whole (§5.8).

## 3. Layers

The diagram at the top shows the layers, top to bottom:

1. **Surfaces** (`surfaces/mcp_main.py`, `surfaces/web/`) are thin adapters: parse, call, shape output.
2. The **service** (`service/`) makes all domain decisions, using the models in `domain/`.
3. The **remote** layer (`remote/`) and the **store** (`store/`) sit underneath. Remote adapters implement the `MailReader` / `MailWriter` ports. The store is SQLite.
4. **Transport and tokens** (`remote/transport.py`, `auth/tokens.py`) are at the bottom.

**Rules:**

- Surfaces import `service` and `domain` only.
- `service` imports `remote/ports.py` (not concrete adapters), `store` and `domain`. It never sees raw Graph or OWS JSON. `bootstrap.py` wires the adapters.
- `remote` returns `domain` models. `remote/graph_mapping.py` is the only file that knows Graph field names.
- `domain` imports nothing from the app.

## 4. Repository layout

```
lrh-outlook-connector/
├─ pyproject.toml                  # Python ≥3.12; msal, msal-extensions[portalocker], mcp, httpx, pydantic,
│                                  # starlette, uvicorn (uv- and pip-compatible; dev group: pytest, ruff)
├─ README.md · AGENTS.md
├─ docs/                           # requirements v4, roadmap, research, this file
├─ research/                       # stdlib probes + README (independent of the package)
├─ src/outlook_connector/
│  ├─ __main__.py                  # CLI dispatch: auth · status · mcp · ui (lazy imports per command)
│  ├─ bootstrap.py                 # lazy composition: tokens → transport → reader → store → services
│  ├─ config.py                    # client profiles, denied pairs, data directory, cache paths
│  │
│  ├─ auth/
│  │  └─ tokens.py                 # one centralized TokenProvider; named profiles are config
│  │
│  ├─ remote/                      # all Microsoft protocol knowledge (async)
│  │  ├─ ports.py                  # MailReader / MailWriter protocols the service depends on
│  │  ├─ transport.py              # shared httpx.AsyncClient and request policy
│  │  ├─ graph.py                  # Graph plumbing: paging, $batch, ImmutableId preference, downloads
│  │  ├─ graph_mail.py             # MailReader over Graph
│  │  └─ graph_mapping.py          # Graph JSON → domain models
│  │  ├─ ows.py                    # OWS plumbing: envelope, headers, item results, the bare-request style
│  │  ├─ ows_mail.py               # MailWriter over Outlook Web (OWS): drafts, send, mutations
│  │  ├─ ows_mapping.py            # domain values → OWS request JSON
│  │  ├─ ids.py                    # Graph ↔ OWS ids
│  │
│  ├─ domain/
│  │  ├─ models.py                 # pydantic models (the single schema source)
│  │  └─ errors.py                 # domain errors
│  │
│  ├─ store/
│  │  └─ db.py                     # account-bound SQLite: account binding, folder cache
│  │
│  ├─ service/
│  │  ├─ mailbox.py                # folders, list, get, search
│  │  ├─ conversations.py                # conversation retrieval (+ branch labelling later)
│  │  ├─ failures.py               # one classification and rendering of export errors
│  │  ├─ cursors.py                # self-contained continuation cursors
│  │  ├─ files.py                  # attachment and .eml downloads (MCP and UI)
│  │  ├─ localfiles.py             # exports/ and downloads/ folders: 7-day cleanup, exclusive file names
│  │  └─ export/
│  │     ├─ orchestrator.py        # ExportRequest → ExportArtifact (the only export path)
│  │     ├─ formatter.py           # TXT rendering
│  │     ├─ attachments.py         # attachment selection policy + safe filenames
│  │     └─ packaging.py           # flat TXT vs single ZIP
│  │  ├─ writes.py                 # drafts, confirmed send
│  │  ├─ mutations.py              # read state, flag, move, delete
│  │
│  └─ surfaces/
│     ├─ mcp_main.py               # FastMCP (stdio) tools
│     └─ web/
│        ├─ web_main.py            # uvicorn launch + idle shutdown (imported only by `ui`)
│        ├─ routes.py              # Starlette JSON API, session-token/Host guard, export download
│        └─ static/                # index.html, app.js (ES module, no build step), app.css
└─ tests/
   ├─ fakes/                       # fake MSAL; in-memory Graph mailbox over httpx.MockTransport
   ├─ unit/                        # config, tokens, CLI, Graph reader, store
   ├─ service/                     # mailbox, conversations, exports (against the fake Graph)
   └─ surfaces/                    # MCP contract, web routes
```

## 5. Module responsibilities

### 5.1 `auth/tokens.py`

- **One centralized provider** serving any number of named profiles (`client_id` + resource scope), all defined in `config.py`. Cache, locking, fail-closed handling and the account check are shared. Profiles today (research §2):
  - `read`: Outlook Mobile `27922004-…` → `https://graph.microsoft.com/Mail.Read`.
  - `write`: One Outlook Web `9199bf20-…` → `https://outlook.office.com/.default`.
- One MSAL `PublicClientApplication` per profile per process; the access token is kept in memory until 5 minutes before expiry. Silent acquisition first. Device code only from the `auth` command, so surfaces never start an interactive sign-in. They raise `AuthenticationRequired` with the exact command to run.
- Encrypted cache via `msal-extensions` by default, **fail-closed** when unavailable. `--unsecure` selects a separate, clearly named plaintext cache file in the same data directory, with a warning on every use. Every process finds it at the same path, wherever it was started.
- Cross-process lock around cache reads and writes.
- The account fingerprint (`tid`+`oid`) must match the store owner (§7).
- The `write` profile is optional. Read-only use works without it, and write tools report "write sign-in required".
- The recorded `AADSTS65002` client/scope pairs (`config.DENIED_PAIRS`) are never requested; a unit test checks the profiles against them.

### 5.2 `remote/transport.py`

- One `httpx.AsyncClient` per process. It keeps connections alive within a call.
- Host allowlist: `graph.microsoft.com`, `outlook.office.com`, `outlook.cloud.microsoft`. No redirects. (Sign-in traffic to `login.microsoftonline.com` goes through MSAL, not this client.)
- Response size caps. Downloads (attachments, MIME) stream.
- Retries **only for idempotent requests**: GETs, and read-style POSTs the caller marks idempotent (`$batch` of GETs). They honor `429` / `Retry-After`. Writes (`write=True`) are sent once: an answer that never completes, or a 5xx, raises `WriteOutcomeUnknown` (the write may have happened); a 4xx or 429 is a definite failure. A 401 still renews the token once, since nothing was processed.
- At most **4 requests in flight** per process: Exchange Online allows about 4 concurrent requests per app and mailbox (and 10,000 per 10 minutes).
- **401** (token rejected: revoked, or a continuous-access-evaluation challenge): the token is renewed once (`force_refresh`, or the `claims` challenge from `WWW-Authenticate`) and the request retried; a second 401 raises `AuthenticationRequired` with the sign-in command. **403** is "access denied" for that item and never asks for a new sign-in.
- Maps HTTP and Graph/OWS errors to domain errors. An error names the operation in progress (`operation()` context, e.g. "While fetching message bodies"), the HTTP status, the service error code and a shortened message, and the `request-id`. Throttling errors state the limits. Logs metadata only.

### 5.3 `remote/ids.py`

- Graph immutable REST id → OWS `ItemId`: swap the base64 alphabet (`-`→`/`, `_`→`+`). The same applies to `conversationId`. Verified live (research §4.2).
- This is the only place that converts IDs.

### 5.4 `remote/graph.py`, `graph_mail.py`, `graph_mapping.py`

**`graph.py`:**
- Sends `Prefer: IdType="ImmutableId"` on every call.
- Follows `nextLink` safely, keeping it on the Graph host.
- Batches with **`$batch`**, up to 20 sub-requests per call. Used to hydrate conversations and exports, count conversation sizes and list attachments, instead of N sequential GETs. Rules:
  - batch request ids are numbers assigned per batch and mapped back. Graph compares them case-insensitively, and immutable ids can differ only by case;
  - at most 2 batches in flight per process, shared by all concurrent callers (each sub-request counts against the mailbox's concurrency limit);
  - throttled or temporarily failing sub-requests (429/502/503/504, the statuses single reads retry) are re-sent in new batches of at most 20 after the advised `Retry-After`, for up to 4 rounds; an item that still fails keeps its last status;
  - results are per item: what is still throttled or failed is returned to the caller, which reports it per message instead of failing the whole call.

**`graph_mail.py`** implements `MailReader`. Operations:
- `list_folders()` with hidden folders included (the service needs them to tell what is out of reach). Folder delta is researched (S2) but not used: the folder cache is refreshed in full.
- `list_messages(folder | mailbox, since, until, page_size, page, skip)`: `skip` starts a folder listing past its newest messages (per-folder listing, below).
- `get_message(id, body_format)` and `get_messages(ids)` (batched; per-item failures returned).
- `conversation(conversation_id)`, which returns all folders, leaves sorting to the caller (`$orderby` is rejected with this filter) and reports truncation past 1,000 messages.
- `conversation_folders(conversation_ids)`: batched (folder, Internet message id) per message of each conversation, for counts.
- `count_messages(folder_ids, window)`: the server's count per folder for a window (`$count`, `ConsistencyLevel: eventual`), in `$batch`. A folder whose sub-request fails is left out of the answer, so one failure (e.g. a folder deleted since the folder list was read) never hides the other counts. The server total needs every in-scope folder counted; the per-folder listing chooses an uncounted folder by its cached total and claims the excluded count only when every excluded folder was counted.
- `search(query)`: `$search`, field-scoped queries passed through. `$search` returns regular ids, so each page's ids are converted with one `POST /me/translateExchangeIds` call into the immutable ids every other call uses (if that call fails, the page keeps its search ids rather than failing).
- `list_attachments(id)`, `list_attachments_many(ids)` (batched) and `attachment_content_ids` (`contentId` via typed `$select`, for the inline images of every message of an export at once: one `$batch` item per image, 20 per batch across messages).
- `download_attachment(id, att_id)` and `download_mime(id)` stream to a file. They are retried like other GETs (429/503 after `Retry-After`, gateway errors, a connection that fails or drops mid-download), each time from scratch; what still fails is a domain error (`Upstream` with no status for a lost connection), so an export marks that one attachment instead of failing.
- Every operation is named for error messages ("While listing attachments: …").

**`graph_mapping.py`:** maps every Graph shape to `domain.models`. Unknown fields are ignored. Missing optional fields become `None`.

### 5.5 `remote/ows.py`, `ows_mail.py`, `ows_mapping.py`

- Split like the Graph side (§5.4): `ows.py` is the client (`Ows`), `ows_mapping.py` builds request bodies (pure functions), `ows_mail.py` holds `OwsMailWriter`.
- `OwsMailWriter` implements `MailWriter`. This is a gap fill (§6), replaceable by a Graph writer where Graph mail write scopes are available.
- The bearer-only OWS envelope and write contracts proven in research §4.1–4.2. Payloads ≤ 2,048 characters go in the `X-OWA-UrlPostData` header. Anchor mailbox, correlation headers.
- `Ows.call(action, body)` sends one action and returns its item results; an item whose `ResponseClass` is not `Success`/`Warning` raises an error naming its `ResponseCode`. The anchor mailbox is the write token's `upn`.
- `Ows.call_request(action, fields)` sends the inbox-rule actions, which use a second style (research §4.4): the request object itself, no `JsonRequest` wrapper and no `Body`; the answer's `WasSuccessful` / `ErrorCode` decide success, and an answer without them is an unknown outcome. Same URL and headers, sent once. Used by `OwsRules` through the `RuleWriter` port (W9).
- Actions:
  - `create_draft` (`CreateItem` with `SaveOnly`, into Drafts; returns the draft id, mapped to Graph's alphabet);
  - `edit_draft` (`UpdateItem` / `SaveOnly`, partial field updates). Replies use EWS's `ReplyToItem` / `ReplyAllToItem` with explicit recipients and subject and an HTML body, so the quoted original keeps its formatting and inline images;
  - `send_draft` (`UpdateItem` with `SendAndSaveCopy` and no field updates on an existing draft, its `ItemId` carrying the change key the draft was read at, `NeverOverwrite`; `SendItem` is not supported over OWS);
  - `set_read` / `set_flag` (`UpdateItem`, one `SetItemField` per message, read receipts suppressed);
  - `move` (`MoveItem`; a well-known target by `DistinguishedFolderId` as proven, any other folder by `FolderId`, pending a live check, V3);
  - `delete` (`DeleteItem` with `MoveToDeletedItems`; there is **no hard delete**).
  Conversation read state is done per message with `UpdateItem` (all copies in scope), so `ApplyConversationAction` is not used.
- Mutations return an outcome per message (`None`, or Outlook's `ResponseCode`). Never retries.

### 5.6 `domain/models.py` and `errors.py`

- Pydantic models: `Folder`, `Recipient`, `MessageSummary` (with `also_in` for merged copies, and `meeting` on meeting mail: kind, start, end, location, out of date; read from Graph's `eventMessage` fields in the same listing, so no extra requests), `Message`, `Attachment`, `Coverage` (with `excluded` counts per `ExclusionReason`: `deleted_or_junk`, `outgoing`), `MessagePage` (with cursor), `ConversationHit` (with `message_count`) + `SearchResult`, `MessageContent`, `ConversationMessage` + `Conversation`, `ConversationSize`, `ExportRequest`, `ExportArtifact`. Drafts add `OutgoingMessage` (explicit text/HTML and reply context), `DraftEdit`, private `DraftMessage`, `DraftResult` and `SendResult`; mutations add `ItemResult` and `MutationResult`.
- Output models serialize optional fields only when set (no nulls, no empty lists): MCP results stay small, and a missing field means its default. Required fields are always present.
- These models are the schema source for MCP (FastMCP derives tool input/output schemas from them) and for the web JSON API. No hand-written schemas.
- `ExportError` (step, status, code, message, request id, likely cause, `retry`, fix) describes a gap in an export or conversation body.
- Errors: `AuthenticationRequired`, `AccountMismatch`, `NotFound`, `InvalidRequest`, `Throttled`, `Upstream`, `WriteOutcomeUnknown`. Each surface maps them to its own protocol. A transport error carries Microsoft's answer as a `Failure` (status, code, shortened message, request id), which batch results also report per item.

### 5.7 `store/`

- One SQLite database per account fingerprint under the user data directory. WAL mode, `busy_timeout`, a connection per operation, short transactions.
- Holds only the account binding (owner fingerprint) and the **folder cache** (id, parent, alias, counts, and a TTL timestamp).
- No message data at all: no summaries, bodies, attachment bytes or mailbox mirror (research §3.2). Mail deleted on the server is gone here too (decided 2026-10-02: no local retention). The summary cache was removed on 2026-10-04 with its only uses (MCP `list_messages(refresh=false)`, the UI's cached first draw, exports by id and the `.eml` file name).
- When the owner fingerprint does not match, the existing database is left untouched and a separate one is used.

### 5.8 `service/`

**`mailbox.py`:**
- `list_folders` and every scope decision use the folder cache only while it is younger than 10 minutes (whichever process saved it; within a long process the map is reloaded at the same age). An older or empty cache, or `refresh=true`, waits for Graph, whose folder levels are fetched in parallel (under a second). Decided 2026-10-04: processes are short-lived and often start after days idle, so the former stale-while-revalidate served a weeks-old tree to the call that needed it, which missed folders created meanwhile (fatal for the per-folder listing below). One refresh at a time per process.
- **Scope rules, shared by list, search, conversations, sizes and export:** each folder gets a category from itself and its parents (`folder_categories`): `hidden` (also Sync Issues) > `deleted_or_junk` > `outgoing`. Deleted Items and Junk Email (with their subfolders; a folder deleted in Outlook sits inside Deleted Items) are left out unless `include_deleted_items` (a folder named in the request is always included); Sent Items, Drafts and Outbox are left out when `include_sent_items` is false; list, search and range exports also leave out meeting mail (invitations, RSVPs, cancellations; `excluded.meeting_mail`) when `include_meeting_mail` is false, per message, so a conversation that is only meeting traffic disappears and one with real replies shows through them, while conversations stay whole (every flag: true shows more mail, false filters more); `coverage.excluded` counts what was left out, per reason.
- **Out of reach:** hidden folders, and items whose folder is not among the mail folders (e.g. Teams meeting records in `SkypeSpacesData/TeamsMeetings`), are always dropped (`excluded.hidden`): from lists, search, conversations, conversation sizes, totals and range/conversation exports. `list_folders` omits hidden folders and naming one is refused. An unknown folder id first refreshes the folder list once (a folder created meanwhile is found); ids still unknown are remembered as outside and never refresh again. Sync Issues and its subfolders (classic Outlook's conflict and failure copies) are out of reach too, recognized by their well-known names since Graph does not always mark them hidden (decided 2026-10-04). **Copies** of one message (same Internet message id: mail sent to yourself or to a list you are on) are shown once, keeping a received copy; `also_in` names the other folders.
- `list_messages(selection, include_sent_items, include_deleted_items, include_total, detail)` fetches from remote and returns `MessagePage` + `Coverage`. A mailbox-wide scope (no `folder`) is listed one of two ways, chosen on every first page (H7, decided 2026-10-04): when the folders the scope leaves out (Junk Email, Deleted Items and their subfolders, hidden folders, and Sent Items, Drafts and Outbox with `include_sent_items=false`) hold at least two thirds of the mailbox's messages (`PER_FOLDER_SHARE`, from the cached folder counts, no request), it lists **folder by folder**: one `$count` batch picks the in-scope folders with mail in the window, each is read newest first from its position in small chunks that grow as the merge takes from it (folders that need more are read in parallel, within the transport's limit of 4), and the merge returns the newest `limit`. The same count batch also counts the window's Deleted Items and Junk (with subfolders), reported in `coverage.excluded` on the first page only (continuation pages do not repeat the full-window count); an excluded count is reported only when every excluded folder was counted. The cursor carries each folder's position (`offsets`, folder id → messages returned), and `coverage.notes` says the listing was per folder and why. Otherwise it lists `/me/messages` and the folder filters apply after paging, so a filtered page can hold fewer than `limit` messages; a page that the filters empty entirely is skipped (bounded). Both return the same messages in the same order; a cursor keeps the method its first page chose. Folder views (a named `folder`) are always one listing. `include_total` adds `server_total` (copies counted separately; meeting mail included even when `include_meeting_mail=false` hides it, and a coverage note says so). `detail=compact` (the MCP default) drops recipients, categories and Internet ids.
- `get_message(id, offset, max_chars, body)` returns a bounded body with continuation; `max_chars=None` returns the whole body (the local web reader: one server read, where chunks would each re-read the message). An unreadable id is `NotFound`, worded without guessing the cause (deleted, moved out of reach, or a wrong id).
- `search(query, since?, until?, folder?, include_sent_items, include_deleted_items, detail)` runs Graph `$search` and groups hits by `conversationId`, each hit with the conversation's `message_count`. `since`/`until` are normalized once here for every caller (naive dates are UTC, aware dates converted to UTC) before the window and the cursor are built. KQL only takes dates, so the query asks for a day more on each side and results are then filtered to the exact `since`/`until`.
- `conversation_sizes(conversation_ids)` counts each conversation's messages the way `get_conversation` lists them (copies once), for the UI and search hits.

**`conversations.py`:**
- `get_conversation(conversation_id, include_deleted_items=False)`:
  1. fetch the conversation from remote;
  2. **sort locally**, oldest first;
  3. hydrate bodies via `$batch` when requested (a body that cannot be fetched is marked in the text).
  Its cursor carries the original selection (conversation, body options, `include_deleted_items`, `max_chars`). Coverage is incomplete when the server listing was truncated (over 1,000 messages) or a body could not be fetched for a reason a retry could fix. Every body that could not be fetched, retryable or not, sets `export_error` on its `ConversationMessage`, is counted in `Conversation.body_errors` and is named in a coverage note.
- `bodies()` returns a body or an `ExportError` for every message (still throttled, access denied, deleted meanwhile); the text shows the `[EXPORT ERROR]` block (`failures.py`) in place of the body.

**`failures.py`:**
- One classification of failed Microsoft requests for exports and conversations: `export_error(step, failure)` turns the remote layer's `Failure` (status, code, shortened message, request id; status `None` = no response) into an `ExportError` with the likely cause, `retry` and the fix: 429/503 throttled (retry), other 5xx or no response service or network (retry), 403 access denied, 404 deleted or moved during the export, anything else unexpected (report it with the request id).
- `error_block` renders the TXT block used by exports and `get_conversation` (which also returns the `ExportError` itself); `error_summary` writes the export header's one "Export errors: …" line.
- Later (E3): build the reply tree from `Message-ID` / `In-Reply-To` / `References`, label branches, with a fallback for the user's own messages that lack headers.

**`writes.py`:**
- `create_draft` validates explicit text/HTML, recipients, subject and reply context; text becomes
  escaped minimal HTML preserving whitespace and NBSP. Intentional HTML passes through, with only
  active web content refused. The private writer takes `DraftMessage`, always HTML.
- `edit_draft` validates partial changes against a current Graph draft. Omitted fields and attachments
  survive. `UpdateItem` saves only changed fields; no local draft registry.
- Both require bounded Graph read-back and return `DraftResult`: id, saved/failed, full server text/HTML,
  message metadata and simple findings. Reply creation retains full quoted-history verification.
- `send_draft(id)` requires an existing draft and the bound write account, and sends once with
  `UpdateItem` / `SendAndSaveCopy` and no field updates, bound to the draft's change key as read
  (`NeverOverwrite`): a draft changed since then is refused with nothing sent (proven live, research
  §4.2). It accepts no content arguments and never
  reconstructs mail. An ambiguous answer reads the exact immutable id: a Sent Items copy proves sent,
  otherwise unknown. Human approval is the host/agent interaction, not a connector token.

**`rules.py` / `remote/ows_rules.py` (W9):**
- Service depends on `RuleWriter`; only `OwsRules` knows rule wire fields and identity envelopes.
- Supported From/Sent to and subject/subject-or-body conditions, Move to folder and Stop processing.
  Non-neutral unsupported fields (including exceptions) mark a rule read-only; its complete server
  revision is retained as a hash for confirmation binding. No unsupported rule is updated or deleted.
  Neutral: description metadata (`DescriptionTimeFormat`/`TimeZone`) and each inactive condition's
  exact "not set" value (`NullImportance`, `NullSensitivity`, `NullInboxRuleMessageFlag`,
  `NullInboxRuleMessageType`, per field), which OWS reports on every rule; any other value is active.
- `NewInboxRule` can report an identity that fresh `GetInboxRule` does not, so creation takes the id
  of the one new rule in read-back that matches the request.
- Each write first returns a stateless proposal with a RULE code bound to account, normalized request
  and every current rule revision. Confirmation re-reads state and refuses changed proposals.
- One OWS write follows confirmation and is always read back. No automatic write retry. Failed
  read-back or ambiguous unverified outcomes return unknown; a mismatched successful write returns failed.
- Toggles are separate from field edits to keep each confirmation to one proven OWS call. Reordering
  includes every supported rule in the requested order, preserves enabled flags and is refused when
  unsupported rules exist. No local rule cache/proposal registry.
- OWS reads target folder names/references rather than Graph ids; folder write targets resolve through
  Mailbox. Target read-back compares the reported folder name; duplicated names cannot prove exact identity.

**`mutations.py`:**
- `set_read` (also per conversation: every message in scope, all copies), `set_flag`, `move(folder)` and `delete` act on **explicit ids only**, at most 100 per call, checked at each entry point before anything is read. The write sign-in must be the bound account.
- Conversations given to `set_read` expand without that limit, up to the 1,000 messages the server lists per conversation; `notes` names a conversation cut there. When the selection has more than 100 messages, `counts` covers all of it and `results` keeps only explicit ids and messages that did not end `done` or `unchanged`.
- Flow: read every message's state through Graph (`get_summaries`, one `$batch`): unknown ids → `not_found`, hidden or outside the mail folders → `failed`, already as wanted → `unchanged` (nothing sent). Then send in chunks of 20, with a status per message (`done`, `not_found`, `failed` with the code). On `WriteOutcomeUnknown`, read the chunk back: `done` where the change is visible, `unknown` elsewhere.
- Categories are not written (parked hard, 2026-10-03: never used). They are still read and returned with each message.
- `move` resolves the target like `list_folders` (hidden folders refused) and refuses Deleted Items. `delete` moves to Deleted Items and leaves messages already in Deleted Items (or its subfolders) alone, since deleting there again would take them out of the folder view.
- All four accept `continue_on_error=true` by default. Clear chunk failures become per-message
  failed results; ambiguous outcomes are read back and remain unknown if that read fails. With false,
  any unresolved/error result in a sent chunk stops subsequent chunks: failed, `not sent`. Earlier
  done/unchanged results and counts always survive; no write is retried.

**`export/`:**
- `orchestrator.py` resolves a selection: conversations, individual messages and/or a range (`since`, `until`, `folder`, `include_sent_items`, paged through `list_messages` with the same scope rules), then merges copies across the whole selection. The range is paged with `skip_returned_copies=False`, so copies on different pages reach that merge and `also_in` names every folder. The selection is refused above `limit` (at most 2,000) with its count. Messages selected by id are authoritative, like `get_message(id)`: they are exported wherever Graph can read them (hidden folders and Sync Issues included) and only labeled and merged with the rest, while conversations and ranges keep the scope rules. They are read from the server in `$batch`; if any cannot be read, the export fails before writing anything: `NotFound` when they are gone, `Throttled` when any is still throttled after the batch retries, `Upstream` otherwise, with their count, the first few ids and their case, and what to do. It then hydrates through `$batch` (reusing bodies fetched while selecting), formats, attaches and packages, and returns an `ExportArtifact` with `messages_excluded`, `unavailable_message_ids`, `export_errors` (failures per step) and `error_summary` (the header's "Export errors" line). A body, attachment download or attachment listing that fails during the export becomes an `ExportError` marked in the file; the export still completes.
- `formatter.py` renders **TXT** (for people): a header (counts, date span, what was left out, and the "Export errors" summary), then per-message headers with the message, conversation and internet ids, `Also in:` for merged copies, `uniqueBody` by default and `full` optional, plus attachment lines and `[EXPORT ERROR]` blocks for what failed; people are separated by `; ` (display names are often "Last, First"). Or **JSONL** (`format=jsonl`, for agents): one record per message with ids, dates, folder, `also_in`, people, the body (or `body: null` and an `export_error` object), attachment records (with the file path inside the ZIP when attachments are included, or their own `export_error`), and `attachments_export_error` when the attachments could not be listed.
- `attachments.py` applies the attachment policy:
  - non-inline attachments by default;
  - inline images only when the rendered body references their `cid:`;
  - `itemAttachment` → `.eml`;
  - sanitized, deduplicated names; identical files (same bytes, e.g. a signature logo on every message) are stored once per output file, and every message points to that file;
  - a failed download becomes an `[EXPORT ERROR]` block ("The attachment <name> could not be downloaded.") and an `export_error` on its JSONL record.
- `packaging.py` decides the output: one flat `.txt` only when the result is a single TXT with no attachment files, otherwise one `.zip` (TXTs at the root, `<stem>/` folders for attachments). The file is created exclusively in the exports directory (a numbered suffix on a name clash, so concurrent exports never overwrite each other).

### 5.9 `surfaces/`

**`mcp_main.py`:**
- FastMCP over stdio. Each tool is a few lines: validate, call the service, return a model.
- Bounded responses with self-contained cursors.
- Attachments, MIME and export artifacts are returned as local file paths, never inline base64.
- Server instructions explain the scope rules, what is out of reach (hidden folders, non-mail items; search is mail only), merged copies, coverage and cursors, `include_sent_items=false` for "latest mail", `include_total`, the export options (`format=jsonl` for analysis), Graph's throttling limits (no parallel tool calls; prefer one range export), and that "access denied" is not a sign-in problem. With send, they will also repeat the send-authorization rule.
- List and search results are compact by default (`detail="full"` for every field).

**`web/`:**
- Starlette JSON API over the same service calls, plus `POST /api/export` (one download; an `X-Export-Errors` header carries the "Export errors" line when something could not be exported, and the UI shows it) and `POST /api/heartbeat` (keeps the idle timer alive while a tab is open).
- `index.html` + `app.js` provide:
  - Outlook-style rows: sender, subject and preview; unread rows with a blue bar and blue subject; Outlook dates ("Fri 9:31 AM"); flag and paperclip icons; file chips under messages with attachments (names fetched after the list loads, one batched `POST /api/attachments` per 200 messages; a chip downloads its file); meeting mail labelled (Invite, Updated, Canceled, Accepted, Tentative, Declined) with its time and place, a conversation showing its current invitation. The whole row is clickable; colors follow Outlook's light and dark themes;
  - a conversation-grouped list that opens on the Inbox. After a list loads, the UI asks for each conversation's real size: one-message conversations are plain rows, conversations show an accurate count. An expanded conversation spans all folders and shows the newest message on top; merged copies carry an "also in" badge, and search matches are marked;
  - an "Invites / RSVPs" switch in the list header, off by default (meeting mail hidden), applied to the list, search and "export this view";
  - a "Deleted / Junk" switch in the list header, applied to the list, search, counts, expansion and exports (always on inside those folders and their subfolders), and "export this view" (the current folder and date range). The folder list shows only reachable folders;
  - a reader that shows the whole chosen body in one request (no size cap; an answer for a message no longer selected is ignored), with attachment download buttons;
  - your display name and profile photo in the header (Graph `/me` and `/me/photos/48x48/$value`, read once per process and kept in memory; initials when no photo is set);
  - a folder picker and a "recent, all mail" view;
  - an online search box;
  - an in-memory filter;
  - selection checkboxes;
  - Ctrl+click (Cmd+click on a Mac) on a row toggles its selection instead of opening it;
  - export options: switches for attachments and quoted history (one quoted-history setting for the reader and exports), and a segmented control for how files are split (per conversation, all in one, per message).
- Binds to localhost only.

## 6. Capability routing and portability

The backend split is a **tenant-specific outcome**, not a design preference. The rule is: **use documented Graph for every capability it can serve; fill only the remaining gaps with OWS.** For Landis+Gyr on 2026-10-02 the probes established (research §2):

| Capability | Graph available? | Backend used | Token profile |
|---|---|---|---|
| Folders, list, get, conversations, search, attachments, MIME, delta | Yes (`Mail.Read` via Outlook Mobile) | **Graph** | `read` |
| Send | No: `Mail.Send` denied for every usable client | **OWS** `CreateItem` | `write` |
| Read state, flag, categories, move, delete | No: `Mail.ReadWrite` denied for every usable client | **OWS** `UpdateItem` / `MoveItem` / `DeleteItem` / `ApplyConversationAction` | `write` |

### 6.1 How the code stays swappable

- **Ports.** `remote/ports.py` defines two protocols: `MailReader` (folders, list, get, conversation, search, attachments, MIME) and `MailWriter` (create_draft, edit_draft, send_draft, set_read, set_flag, move, delete). The service depends **only on these ports**, never on a concrete backend.
- **Adapters.** `remote/graph_mail.py` implements `MailReader`. `remote/ows_mail.py` implements `MailWriter`. Each adapter maps its protocol to the same `domain` models, so swapping an adapter never changes the service, surfaces, store or tests above it.
- **Wiring.** `bootstrap.py` picks one adapter per port, and one token profile per adapter, from `config.py`. There is exactly one implementation per port at runtime. No dual backends and no automatic cross-backend fallback (a write must never be retried through a second backend).

### 6.2 Adapting to another tenant or a policy change

Re-run the probe suite first ([`research/README.md`](../research/README.md): `auth.py`, `graph_scopes.py`, then the capability probes), and update the research record. Then:

| Situation | Remote layer change | Token layer change |
|---|---|---|
| **Full Graph** (a client — first-party or a registered app — can obtain `Mail.ReadWrite` + `Mail.Send`) | Add `remote/graph_mail_writer.py` implementing `MailWriter` (`PATCH /messages/{id}` for read/flag/categories, `POST /messages/{id}/move`, move to `deleteditems`, `POST /me/sendMail`). Wire it in `bootstrap.py`. **Delete** `remote/ows.py`, `ows_mail.py`, `ows_mapping.py` and the OWS id mapping in `ids.py`. | Point the `write` profile at the Graph client/scopes, or merge it into `read` if one client covers both. Remove the One Outlook Web profile. |
| **Partial change** (e.g. Graph `Mail.ReadWrite` but no `Mail.Send`) | Split `MailWriter` wiring per capability group only if needed: Graph for mutations, OWS for send. Keep the rule "one backend per capability". | `write` profile for Graph, plus a `send` profile for OWS. |
| **No Graph mail at all** (no client can get `Mail.Read`) | Add `remote/ows_reader.py` implementing `MailReader`. The OWS read actions are already proven (research §4.3). Search moves to Substrate `searchservice/api/v2/query`, which works with client A (research §5). Change detection would need OWS sync actions (to be researched). | The `read` profile points at One Outlook Web (`outlook.office.com/.default`), and a `search` profile is added (`outlook.office.com/search/.default`). |
| **Different first-party clients work** | No change if the scopes are equivalent. | Change the profile's `client_id` in config. Keep the AADSTS65002 list (`DENIED_PAIRS`) per tenant; its test guards the profiles. |

What never changes: `domain/`, `service/`, `store/`, `surfaces/`, and their tests. If a backend change forces an edit there, the port boundary has leaked and should be fixed instead.

## 7. Data and identity

- **Account fingerprint:** a hash of `tid` + `oid`. It selects the database and must match both token profiles.
- **Item key:** the Graph immutable id, within the account's own store. Immutable ids survive moves within the mailbox (verified).
- **Conversation key:** Graph `conversationId`.
- **Locations:**
  - data directory: `%USERPROFILE%/.lrh-outlook-connector` on Windows, `$XDG_DATA_HOME/lrh-outlook-connector` (or `~/.local/share/...`) elsewhere. Not under AppData: Windows redirects files that packaged apps (the Claude desktop app and the MCP servers it starts) create there into a private per-app copy that terminals and Explorer never see, which split tokens, store and exports in two (found 2026-10-03);
  - token cache: `<data directory>/token-cache.bin` (encrypted), or `token-cache.plaintext-dev.json` with `--unsecure`;
  - `OUTLOOK_CONNECTOR_HOME` overrides the data directory (tests, portability);
  - store: `<data directory>/accounts/<fingerprint>/mail.sqlite3` (account binding and folder cache only);
  - exports and downloaded attachments: `<data directory>/exports/`, removed after 7 days. Attachment downloads during an export use an OS temp directory, removed when the export is packaged.

## 8. Surfaces: tools and endpoints

| MCP tool | Service call | Annotations |
|---|---|---|
| `list_folders` (reachable folders only) | `mailbox.folders` | read-only |
| `list_messages(folder?, since?, until?, limit?, cursor?, include_sent_items=True, include_deleted_items, include_total, detail=compact)` | `mailbox.list_messages` | read-only |
| `search_messages(query, since?, until?, folder?, limit?, cursor?, include_sent_items=True, include_deleted_items, detail=compact)` | `mailbox.search` | read-only |
| `get_conversation(conversation_id, include_bodies=True, body=unique\|full, include_deleted_items=False, max_chars, cursor?)` | `conversations.get_conversation` | read-only |
| `get_message(id, offset=0, max_chars, body=unique\|full\|html)` | `mailbox.get_message` | read-only |
| `list_attachments(id)` | `mailbox.attachments` | read-only |
| `download_attachment(id, attachment_id)` · `save_message_mime(id)` | `files` | read-only (local file) |
| `auth_status()` | `tokens.status` (offline) | read-only |
| `export_messages(conversation_ids?, message_ids?, since?, until?, folder?, include_sent_items=True, limit<=2000, format=txt\|jsonl, include_attachments, combine, body, include_deleted_items)` | `export.orchestrator` | read-only (local file) |
| `create_draft(..., text_body?, html_body?)` | `writes.create_draft` | not read-only, not destructive, closed world |
| `edit_draft(draft_id, ...)` | `writes.edit_draft` | not read-only, not destructive, closed world |
| `send_draft(draft_id)` | `writes.send_draft` | destructive, open-world |
| `list_rules()` | `rules.list_rules` | read-only, write sign-in |
| `create_rule(changes, user_confirmation?)` / `update_rule(id, changes, user_confirmation?)` / `reorder_rules(ids, user_confirmation?)` / `delete_rule(id, user_confirmation?)` | `rules` | destructive, non-idempotent, proposed/confirmed, single write |
| `move_messages(message_ids, folder)` · `delete_messages(message_ids)` | `mutations.move` / `mutations.delete` | destructive, idempotent |
| `set_read_state(read, message_ids?, conversation_ids?, include_deleted_items)` · `set_flag(message_ids, flagged)` | `mutations` | not read-only, not destructive, idempotent |

Web endpoints mirror the read tools (`GET /api/folders`, `/api/messages`, `/api/search`, `/api/conversations/{id}`, `/api/messages/{id}`, all with `include_deleted_items`) and add `GET /api/messages/{id}/attachments/{attachment_id}` (download), `POST /api/conversation-sizes` (per-conversation message counts, one Graph `$batch` per 20 conversations), `POST /api/export`, `GET /api/status`, `GET /api/me` and `/api/me/photo` (the header's name and photo), and `POST /api/heartbeat`. Every `/api` call needs the per-run session token embedded in the page and a localhost Host header. Write tools are MCP-first. UI write actions are optional later.

## 9. Cross-cutting concerns

| Concern | Rule |
|---|---|
| Logging | Operation, status, counts, durations, backend. Never subjects, bodies, addresses, query text, tokens, cookies or ids in clear (logged URL paths show `{id}` instead of ids). |
| Retries | Idempotent requests only, bounded, honoring `Retry-After`. Throttled `$batch` items are re-sent in batches of at most 20. Writes never retry. An ambiguous write → re-read state. |
| Throttling | Exchange Online: about 4 concurrent requests and 10,000 requests per 10 minutes per app and mailbox; each `$batch` item counts. At most 4 requests and 2 batches in flight per process. Clients are told the limits (MCP instructions, error messages). |
| Bounds | Every list/search/conversation/body response is size-bounded with a cursor. Binary content goes through resources or files. |
| Errors | Domain errors from the service, naming the operation, service code/message and request id. 401 → renew once, then `AuthenticationRequired`; 403 → access denied (no new sign-in). The MCP surface returns tool errors, the web surface returns HTTP status + JSON. Bulk operations report per-item gaps instead of failing whole. |
| Security | Localhost-only web binding. Host allowlist for outbound calls. No arbitrary URL fetching. No browser credentials, ever. |

## 10. Tooling

- **uv** for the environment, with a committed `uv.lock` (`uv sync`). pip works too (`pip install -e . --group dev`). The `outlook-connector` console script comes from `pyproject.toml`.
- **ruff** for lint and format, **pyright** (standard mode) for types: `uv run ruff check .`, `uv run ruff format --check .`, `uv run pyright`, `uv run pytest`.
- **pytest** with **`httpx.MockTransport`** fakes serving synthetic Graph/OWS payloads. Fixtures are never derived from real captures.
- The research probes in `research/` stay stdlib and separate, and are excluded from pytest and pyright. They are for discovering APIs, not for testing the app.

## 11. Build order

| Step | Modules | Roadmap |
|---|---|---|
| 1. Skeleton | `config`, `auth/tokens`, `remote/transport`, `ports`, `graph`, `domain/models`, `errors`, `__main__ auth/status` | A1, A2 |
| 2. Reads | `graph_mail`, `graph_mapping`, `store/*`, `service/mailbox`, `conversations` | B1, S1, S2, L1, T1, L2 |
| 3. MCP | `surfaces/mcp_main` | M1 |
| 4. Export | `service/export/*` | E1 |
| 5. UI | `surfaces/web/*` | U1 |
| 6. Send | `remote/ows` (draft, send), `remote/ids` (Graph ↔ OWS ids), `service/writes` (propose, draft, send) | W0, W1 |
| 7. Mutations | `remote/ows` (update/move/delete), `service/mutations` | W2–W5 |
| Parked | branch-aware conversations in `service/conversations` | R4, E3 |

## 12. Decisions

| # | Question | Decision (2026-10-02) |
|---|---|---|
| A1 | Auth | **MSAL + encrypted cache** (`msal-extensions`, fail-closed). An explicit `--unsecure` plaintext cache (a separate file in the data directory) is available during development. |
| A2 | Package / CLI name | **`outlook_connector` / `outlook-connector`** |

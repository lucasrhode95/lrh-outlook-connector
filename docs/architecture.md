# Outlook Connector — Architecture

**Status:** Accepted (2026-10-02)  
**Inputs:** [Requirements v4](outlook-requirements-v4.md) · [Roadmap](outlook-roadmap.md) · [API research](outlook-api-research.md)

![Outlook connector architecture](architecture.svg)

This document describes how the application is built: processes, layers, modules, data, and the main flows. *What* it must do is in v4. *Why* each API choice was made is in the research doc.

---

## 1. Principles

1. **Remote first.** Outlook's servers do the work: listing, filtering, search, conversation grouping. The app keeps locally only what the server cannot give back.
2. **No always-on service.** Every entry point is a short-lived process started on demand. Several may run at once.
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

**Consequences:**

- **Concurrency is cross-process.** Two agent sessions and the UI may run at the same time. The shared resources are the **token cache** (MSAL file cache with `msal-extensions` cross-process locking) and the **SQLite store** (WAL mode, `busy_timeout`, short write transactions, a connection per operation).
- **Startup cost matters**, because every agent session pays it.
  - Nothing runs at startup: no sync, no folder walk, no token refresh before the first call.
  - The web stack (Starlette, uvicorn) is imported only by `ui`, never by `mcp`.
  - The folder cache refreshes lazily with a short TTL, using cheap folder delta.
- **No in-memory state outlives a call.** MCP continuation cursors are self-contained (they encode the remote `nextLink` or offset plus the original selection). They survive a client restart.
- **Exports run in the process that asked.** The UI shows progress in-process. An MCP export completes inside the tool call and returns a resource or file path. There is no background job queue.

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
├─ pyproject.toml · uv.lock        # Python ≥3.12; deps: mcp, msal, msal-extensions
│                                  # (mcp already brings httpx, pydantic, starlette, uvicorn, anyio)
├─ README.md · AGENTS.md
├─ docs/                           # requirements v4, roadmap, research, this file
├─ research/                       # stdlib probes + README (independent of the package)
├─ src/outlook_connector/
│  ├─ __main__.py                  # CLI dispatch: auth · mcp · ui · status (lazy imports per command)
│  ├─ bootstrap.py                 # lazy factory: config → auth → transport → clients → store → service
│  ├─ config.py                    # frozen settings: client ids, scopes, allowed hosts, paths, limits
│  ├─ logging.py                   # stdlib logging + redaction filter
│  │
│  ├─ auth/
│  │  └─ tokens.py                 # one centralized TokenProvider; named profiles are config
│  │
│  ├─ remote/                      # all Microsoft protocol knowledge (async)
│  │  ├─ ports.py                  # MailReader / MailWriter protocols the service depends on
│  │  ├─ transport.py              # shared httpx.AsyncClient and request policy
│  │  ├─ ids.py                    # Graph immutable id ↔ OWS id; conversation ids
│  │  ├─ graph.py                  # Graph plumbing: paging, $batch, ImmutableId preference
│  │  ├─ graph_mail.py             # mail reads (folders, list, get, conversation, search, attachments, MIME)
│  │  ├─ graph_mapping.py          # Graph JSON → domain models
│  │  └─ ows.py                    # OWS envelope + write actions
│  │
│  ├─ domain/
│  │  ├─ models.py                 # pydantic models (the single schema source)
│  │  └─ errors.py                 # domain errors
│  │
│  ├─ store/
│  │  ├─ db.py                     # sqlite3 connection policy, owner fingerprint check
│  │  ├─ schema.sql                # current schema
│  │  └─ retained.py               # folder cache, retained messages/bodies, tombstones
│  │
│  ├─ service/
│  │  ├─ mailbox.py                # folders, list, get, search
│  │  ├─ threads.py                # conversation retrieval (+ branch labelling later)
│  │  ├─ reconcile.py              # deleted/moved detection and retention merge
│  │  ├─ writes.py                 # send + mailbox mutations
│  │  └─ export/
│  │     ├─ orchestrator.py        # ExportRequest → ExportArtifact (the only export path)
│  │     ├─ formatter.py           # TXT rendering
│  │     ├─ attachments.py         # attachment selection policy + safe filenames
│  │     └─ packaging.py           # flat TXT vs single ZIP
│  │
│  └─ surfaces/
│     ├─ mcp_main.py               # FastMCP (stdio) tools + resources
│     └─ web/
│        ├─ main.py                # Starlette app + uvicorn launch (imported only by `ui`)
│        ├─ routes.py              # JSON API, export download, progress
│        └─ static/                # index.html, app.js (ES modules, no build step), app.css
└─ tests/
   ├─ fakes/                       # httpx.MockTransport handlers with synthetic Graph/OWS payloads
   ├─ unit/                        # ids, mapping, formatter, packaging, attachment policy, reconcile
   ├─ service/                     # threads, search grouping, retention merge, writes (against fakes)
   └─ surfaces/                    # MCP contract (schemas, annotations), web routes
```

## 5. Module responsibilities

### 5.1 `auth/tokens.py`

- **One centralized provider** serving any number of named profiles (`client_id` + resource scope), all defined in `config.py`. Cache, locking, fail-closed handling and the account check are shared. Profiles today (research §2):
  - `read`: Outlook Mobile `27922004-…` → `https://graph.microsoft.com/Mail.Read`.
  - `write`: One Outlook Web `9199bf20-…` → `https://outlook.office.com/.default`.
- MSAL `PublicClientApplication` per profile. Silent acquisition first. Device code only from the `auth` command, so surfaces never start an interactive sign-in. They raise `AuthenticationRequired` with the exact command to run.
- Encrypted cache via `msal-extensions` by default, **fail-closed** when unavailable. `--unsecure` selects a separate, clearly named plaintext cache file in the same data directory, with a warning on every use. Every process finds it at the same path, wherever it was started.
- Cross-process lock around cache reads and writes.
- The account fingerprint (`tid`+`oid`) must match the store owner (§7).
- The `write` profile is optional. Read-only use works without it, and write tools report "write sign-in required".
- A guard refuses the recorded `AADSTS65002` client/scope pairs.

### 5.2 `remote/transport.py`

- One `httpx.AsyncClient` per process. It keeps connections alive within a call.
- Host allowlist: `graph.microsoft.com`, `outlook.office.com`, `outlook.cloud.microsoft`, `login.microsoftonline.com`. No redirects.
- Response size caps. Downloads (attachments, MIME, export) stream.
- Retries **only for idempotent GETs**, honoring `429` / `Retry-After`. POSTs are never retried.
- A concurrency limiter (e.g. 4–8 in flight) keeps parallel fetches polite.
- Maps HTTP and Graph/OWS errors to domain errors. Logs metadata only.

### 5.3 `remote/ids.py`

- Graph immutable REST id → OWS `ItemId`: swap the base64 alphabet (`-`→`/`, `_`→`+`). The same applies to `conversationId`. Verified live (research §4.2).
- This is the only place that converts IDs.

### 5.4 `remote/graph.py`, `graph_mail.py`, `graph_mapping.py`

**`graph.py`:**
- Sends `Prefer: IdType="ImmutableId"` on every call.
- Follows `nextLink` safely, keeping it on the Graph host.
- Batches with **`$batch`**, up to 20 sub-requests per call. Used to hydrate threads and exports instead of N sequential GETs.

**`graph_mail.py`** implements `MailReader`. Operations:
- `folders()` with hidden folders included, and folder delta.
- `list_messages(folder | mailbox, since, until, top)`.
- `get_message(id, body=text|html|unique, headers?)`.
- `conversation(conversation_id)`, which returns all folders and leaves sorting to the caller (`$orderby` is rejected with this filter).
- `search(query)`: `$search`, field-scoped queries passed through.
- `attachments(id)` with `contentId` via typed `$select`.
- `stream_attachment(id, att_id)` and `stream_mime(id)`.

**`graph_mapping.py`:** maps every Graph shape to `domain.models`. Unknown fields are ignored. Missing optional fields become `None`.

### 5.5 `remote/ows.py`

- Implements `MailWriter`. This is a gap fill (§6), replaceable by a Graph writer where Graph mail write scopes are available.
- The bearer-only OWS envelope and write contracts proven in research §4.1–4.2. Payloads ≤ 2,048 characters go in the `X-OWA-UrlPostData` header. Anchor mailbox, correlation headers.
- Actions:
  - `send` (`CreateItem` with `SendAndSaveCopy`);
  - `set_read` / `set_flag` / `set_categories` (`UpdateItem`);
  - `move` (`MoveItem`);
  - `delete` (`DeleteItem` with `MoveToDeletedItems`; there is **no hard delete**);
  - `set_conversation_read` (`ApplyConversationAction`).
- Returns per-item `WriteResult`s. Never retries.

### 5.6 `domain/models.py` and `errors.py`

- Pydantic models: `Folder`, `Recipient`, `MessageSummary`, `Message`, `Attachment`, `Thread`, `SearchHit`, `Coverage`, `Page[T]` (with cursor), `HistorySelection`, `ExportRequest`, `ExportArtifact`, `OutgoingMessage`, `WriteResult`.
- These models are the schema source for MCP (FastMCP derives tool input/output schemas from them) and for the web JSON API. No hand-written schemas.
- Errors: `AuthenticationRequired`, `NotFound`, `InvalidRequest`, `Throttled`, `Upstream`, `WriteOutcomeUnknown`. Each surface maps them to its own protocol.

### 5.7 `store/`

- One SQLite database per account fingerprint under the user data directory. WAL mode, `busy_timeout`, a connection per operation, short transactions.
- Holds only:
  - the **folder cache** (id, parent, alias, counts, and a TTL timestamp);
  - **retained messages**: metadata and body for messages the app has read or exported, so content survives a later server-side deletion;
  - **tombstones** (`is_deleted`, `deleted_at`);
  - attachment metadata for retained messages.
- Attachment bytes and a full mailbox mirror are not stored (research §3.2).
- When the owner fingerprint does not match, the existing database is left untouched and a separate one is used.

### 5.8 `service/`

**`mailbox.py`:**
- `list_folders` uses the cache when fresh and folder delta otherwise.
- `list_messages(selection, refresh=True)` fetches from remote, merges retained rows that are deleted remotely (labelled), and returns `Page` + `Coverage`.
- `get_message(id, offset, max_chars, body)` returns a bounded body with continuation and retains what it fetched.
- `search(query, since?, until?, folder?)` runs Graph `$search` and groups hits by `conversationId`. Coverage reports "server search; retained-deleted mail not included".

**`threads.py`:**
- `get_thread(conversation_id, include_deleted_items=False)`:
  1. fetch the conversation from remote;
  2. **sort locally**;
  3. hydrate bodies via `$batch` when requested;
  4. merge retained-deleted messages.
- Later (E3): build the reply tree from `Message-ID` / `In-Reply-To` / `References`, label branches, with a fallback for the user's own messages that lack headers.

**`reconcile.py`:**
- When a message disappears remotely, it is GET-checked by immutable id. Only a **404** marks it deleted. Graph reports `reason: "deleted"` even for soft deletes (research §3.6).
- Moves update the folder by id.
- Known bodies are never erased.

**`writes.py`:**
- `send` checks `user_confirmation`, revalidates every material field against the confirmed proposal, never retries, and on an ambiguous result checks Sent Items.
- `move` / `delete` / `set_read` / `set_flag` / `set_categories` act on **explicit ids only**. Folder targets are resolved via `list_folders`.
- Every write returns per-item results and updates the store afterwards.

**`export/`:**
- `orchestrator.py` resolves a selection (threads + individual messages, deduplicated). It hydrates through `$batch` with bounded concurrency, formats, attaches and packages. It returns an `ExportArtifact` and reports progress through a callback.
- `formatter.py` renders the TXT: per-message headers, a deleted marker, `uniqueBody` by default and `full` optional, plus attachment lines.
- `attachments.py` applies the attachment policy:
  - non-inline attachments by default;
  - inline images only when the rendered body references their `cid:`;
  - `itemAttachment` → `.eml`;
  - sanitized, deduplicated names;
  - a failure becomes an `[Attachment unavailable: name]` line.
- `packaging.py` decides the output: one flat `.txt` only when the result is a single TXT with no attachment files, otherwise one `.zip` (TXTs at the root, `<stem>/` folders for attachments). It streams to a temp file.

### 5.9 `surfaces/`

**`mcp_main.py`:**
- FastMCP over stdio. Each tool is a few lines: validate, call the service, return a model.
- Bounded responses with self-contained cursors.
- Attachments, MIME and export artifacts are exposed as resources or file paths, never inline base64.
- Server instructions repeat the send-authorization rule.

**`web/`:**
- Starlette JSON API over the same service calls, plus `POST /api/export` (one download) and export progress.
- `index.html` + `app.js` provide:
  - a thread-grouped list;
  - a folder picker and a "recent, all mail" view;
  - an online search box;
  - an in-memory filter;
  - selection checkboxes;
  - export options.
- Binds to localhost only.

## 6. Capability routing and portability

The backend split is a **tenant-specific outcome**, not a design preference. The rule is: **use documented Graph for every capability it can serve; fill only the remaining gaps with OWS.** For Landis+Gyr on 2026-10-02 the probes established (research §2):

| Capability | Graph available? | Backend used | Token profile |
|---|---|---|---|
| Folders, list, get, conversations, search, attachments, MIME, delta | Yes (`Mail.Read` via Outlook Mobile) | **Graph** | `read` |
| Send | No: `Mail.Send` denied for every usable client | **OWS** `CreateItem` | `write` |
| Read state, flag, categories, move, delete | No: `Mail.ReadWrite` denied for every usable client | **OWS** `UpdateItem` / `MoveItem` / `DeleteItem` / `ApplyConversationAction` | `write` |

### 6.1 How the code stays swappable

- **Ports.** `remote/ports.py` defines two protocols: `MailReader` (folders, list, get, conversation, search, attachments, MIME, delta) and `MailWriter` (send, set_read, set_flag, set_categories, move, delete, set_conversation_read). The service depends **only on these ports**, never on a concrete backend.
- **Adapters.** `remote/graph_mail.py` implements `MailReader`. `remote/ows.py` implements `MailWriter`. Each adapter maps its protocol to the same `domain` models, so swapping an adapter never changes the service, surfaces, store or tests above it.
- **Wiring.** `bootstrap.py` picks one adapter per port, and one token profile per adapter, from `config.py`. There is exactly one implementation per port at runtime. No dual backends and no automatic cross-backend fallback (a write must never be retried through a second backend).

### 6.2 Adapting to another tenant or a policy change

Re-run the probe suite first ([`research/README.md`](../research/README.md): `auth.py`, `graph_scopes.py`, then the capability probes), and update the research record. Then:

| Situation | Remote layer change | Token layer change |
|---|---|---|
| **Full Graph** (a client — first-party or a registered app — can obtain `Mail.ReadWrite` + `Mail.Send`) | Add `remote/graph_mail_writer.py` implementing `MailWriter` (`PATCH /messages/{id}` for read/flag/categories, `POST /messages/{id}/move`, move to `deleteditems`, `POST /me/sendMail`). Wire it in `bootstrap.py`. **Delete** `remote/ows.py` and the OWS id mapping in `ids.py`. | Point the `write` profile at the Graph client/scopes, or merge it into `read` if one client covers both. Remove the One Outlook Web profile. |
| **Partial change** (e.g. Graph `Mail.ReadWrite` but no `Mail.Send`) | Split `MailWriter` wiring per capability group only if needed: Graph for mutations, OWS for send. Keep the rule "one backend per capability". | `write` profile for Graph, plus a `send` profile for OWS. |
| **No Graph mail at all** (no client can get `Mail.Read`) | Add `remote/ows_reader.py` implementing `MailReader`. The OWS read actions are already proven (research §4.3). Search moves to Substrate `searchservice/api/v2/query`, which works with client A (research §5). Change detection would need OWS sync actions (to be researched). | The `read` profile points at One Outlook Web (`outlook.office.com/.default`), and a `search` profile is added (`outlook.office.com/search/.default`). |
| **Different first-party clients work** | No change if the scopes are equivalent. | Change the profile's `client_id` in config. Keep the AADSTS65002 guard list per tenant. |

What never changes: `domain/`, `service/`, `store/`, `surfaces/`, and their tests. If a backend change forces an edit there, the port boundary has leaked and should be fixed instead.

## 7. Data and identity

- **Account fingerprint:** a hash of `tid` + `oid`. It selects the database and must match both token profiles.
- **Item key:** `(fingerprint, mailbox, Graph immutable id)`. Immutable ids survive moves within the mailbox (verified).
- **Conversation key:** Graph `conversationId`.
- **Locations:**
  - token cache: `%LOCALAPPDATA%/lrh-outlook-connector/token-cache.bin` (encrypted), or `token-cache.plaintext-dev.json` with `--unsecure`;
  - `OUTLOOK_CONNECTOR_HOME` overrides the data directory (tests, portability);
  - store: `%LOCALAPPDATA%/lrh-outlook-connector/<fingerprint>/mail.sqlite3`;
  - export temp files: OS temp directory, cleaned after delivery.

## 8. Surfaces: tools and endpoints

| MCP tool | Service call | Annotations |
|---|---|---|
| `list_folders` | `mailbox.list_folders` | read-only |
| `list_messages(folder?, since?, until?, limit?, refresh=True, cursor?)` | `mailbox.list_messages` | read-only |
| `search_messages(query, since?, until?, folder?, cursor?)` | `mailbox.search` | read-only |
| `get_thread(conversation_id, include_deleted_items=False, cursor?)` | `threads.get_thread` | read-only |
| `get_message(id, offset=0, max_chars, body=unique\|full\|html)` | `mailbox.get_message` | read-only |
| `list_attachments(id)` + resource `attachment://{message}/{attachment}` | `export.attachments` | read-only |
| `export_messages(threads?, messages?, include_attachments, combine_per_thread, combine_all, body)` | `export.orchestrator` | read-only (local file) |
| `send_email(message, user_confirmation)` | `writes.send` | destructive, open-world |
| `move_messages(ids, folder)` · `delete_messages(ids)` | `writes.move` / `writes.delete` | destructive |
| `set_read_state(ids, read)` · `set_flag(ids, flagged)` · `set_categories(ids, categories)` | `writes.update` | not read-only, not destructive |

Web endpoints mirror the read tools (`GET /api/folders`, `/api/messages`, `/api/search`, `/api/threads/{id}`, `/api/messages/{id}`) and add `POST /api/export` plus `GET /api/export/{id}/progress`. Write tools are MCP-first. UI write actions are optional later.

## 9. Cross-cutting concerns

| Concern | Rule |
|---|---|
| Logging | Operation, status, counts, durations, backend. Never subjects, bodies, addresses, query text, tokens, cookies or ids in clear. |
| Retries | Idempotent GETs only, bounded, honoring `Retry-After`. Writes never retry. An ambiguous write → re-read state. |
| Bounds | Every list/search/thread/body response is size-bounded with a cursor. Binary content goes through resources or files. |
| Errors | Domain errors from the service. The MCP surface returns tool errors, the web surface returns HTTP status + JSON. |
| Security | Localhost-only web binding. Host allowlist for outbound calls. No arbitrary URL fetching. No browser credentials, ever. |

## 10. Tooling

- **uv** for the environment and lockfile. The `outlook-connector` console script comes from `pyproject.toml`.
- **ruff** for lint and format, **pyright** for types (strict on `domain`, `remote`, `service`).
- **pytest** with **`httpx.MockTransport`** fakes serving synthetic Graph/OWS payloads. Fixtures are never derived from real captures.
- The research probes in `research/` stay stdlib and separate, and are excluded from pytest and pyright. They are for discovering APIs, not for testing the app.

## 11. Build order

| Step | Modules | Roadmap |
|---|---|---|
| 1. Skeleton | `config`, `auth/tokens`, `remote/transport`, `ports`, `ids`, `graph`, `domain/models`, `errors`, `__main__ auth/status` | A1, A2 |
| 2. Reads | `graph_mail`, `graph_mapping`, `store/*`, `service/mailbox`, `threads`, `reconcile` | B1, S1, S2, L1, T1, L2 |
| 3. MCP | `surfaces/mcp_main` | M1 |
| 4. Export | `service/export/*` | E1 |
| 5. UI | `surfaces/web/*` | U1 |
| 6. Send | `remote/ows` (send), `service/writes` (send) | W1 |
| 7. Mutations | `remote/ows` (update/move/delete), `service/writes` | W2–W5 |
| Later | branch-aware threads in `service/threads` | R4, E3 |

## 12. Decisions

| # | Question | Decision (2026-10-02) |
|---|---|---|
| A1 | Auth | **MSAL + encrypted cache** (`msal-extensions`, fail-closed). An explicit `--unsecure` plaintext cache (a separate file in the data directory) is available during development. |
| A2 | Package / CLI name | **`outlook_connector` / `outlook-connector`** |

# Outlook Connector Roadmap

Outstanding decisions, implementation, fixes and validation only. Completed behavior belongs in
the requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-07**.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §11](architecture.md)
- API evidence: [API research](outlook-api-research.md)

## Current priority

1. **W11:** group the three scope flags into one `scope` object across every tool (takes
   precedence: later items are written against the new parameter shape).
2. **W10 → W8:** remove `edit_draft` (drafts are composed once), then native signatures: CRUD,
   default selection, and the default inserted into new drafts unless told otherwise.
3. **Correctness:** H19, H22 and H26.
4. **Performance and cleanup:** H30–H32 and H34–H38.
5. **W12:** README rewrite and architecture doc cleanup, last, once everything above is merged.

Status wording:

- **Decided, not built** — behavior is settled; implementation remains.
- **Pending** — a concrete problem and next fix are known.
- **Decision needed** — an owner choice remains.
- **Later** — useful work outside the current build sequence.

# Send and mailbox features

## W11 — One `scope` object instead of three scope flags

**Status:** Decided, not built (2026-10-08). First in priority.

**Why:** every read/export tool takes three loose booleans (`include_sent_items`,
`include_meeting_mail`, `include_deleted_items`). They are one concept, the scope, and should read as
one. Clarity change only: behavior must not change.

**Design:**

- One model, `Scope(extra="forbid")`: `sent_items: bool = True` (Sent Items, Drafts, Outbox),
  `meeting_mail: bool = True` (invitations, RSVPs, cancellations), `deleted_items: bool = False`
  (Deleted Items and Junk; a folder you name is always included). True always shows more mail.
  Keys are independent: an omitted key keeps its default (no restating, unlike a list of enums).
- MCP tools that take any of the flags take `scope: Scope = Scope()` instead: list_messages,
  search_messages, get_conversation, export_messages, set_read_state and any other found.
- One model everywhere. If a caller explicitly sets a key that has no effect for that tool (e.g.
  `meeting_mail` on get_conversation, whose conversations stay whole), the service entry point raises
  InvalidRequest naming the key; defaults merely present are fine.
- Service methods take one `scope: Scope` instead of three keyword arguments; `ExportRequest` gets a
  `scope` field. Cursor state uses the new keys (fresh format, no old-cursor reader).
- Web: GET endpoints take `sent_items`, `meeting_mail`, `deleted_items` query params; JSON bodies
  (export) take a `scope` object; the UI's request builders change with them, UI behavior identical.
- Exclusion texts and error messages name the new keys (e.g. `scope.deleted_items=false`).
- Docs: MCP server instructions and tool descriptions, README, architecture, requirements, and H22's
  text in this roadmap. No compatibility aliases.

## W10 — Drafts are composed once: remove `edit_draft` (enabler of W8)

**Status:** Decided, not built.

**Decision (2026-10-07):** remove the `edit_draft` tool. A draft is composed once by `create_draft`
and never patched. When the user wants to change a draft (text, recipients, subject or, with W8,
the signature), the agent creates a new draft with the full intended content and then deletes the
old one with `delete_messages` (moved to Deleted Items, so it stays recoverable).

**Why:** `edit_draft` replaces the whole body. With W8 placing a signature block (`id="Signature"`,
research §4.6) and replies carrying Exchange's quoted history, an in-place body edit must either
make the agent resubmit signature and quote HTML, or make the connector parse the body into message,
signature and quote regions on every draft read-back, with text/HTML ambiguity and refusals for
drafts edited in Outlook. Today a body edit on a reply already drops the quote silently. Composing
from parts avoids all of it: the connector writes `message + signature` and, for replies, Exchange
appends the quote (`NewBodyContent`), so no draft body is ever parsed or rewritten.

**Implementation:**

- Remove `edit_draft` end to end: MCP tool, service method, `MailWriter.edit_draft`, the OWS
  `UpdateItem`/`SaveOnly` field updates, `DraftEdit`, fakes, tests and docs. No compatibility shim.
- Orient the agent where it looks when a change is requested (descriptions and instructions are
  loaded once per session at the MCP handshake; no per-result note field):
  - `create_draft` description: drafts are never edited; to change one, create the replacement
    with the full content (for a reply, the same `reply_to_message_id`), check its read-back, then
    delete the old draft with `delete_messages`, and use the new id from then on (e.g. `send_draft`).
  - Drafts bullet of the MCP server instructions: the same rule, replacing the `edit_draft`
    sentence. Never delete first. Tell the user the previous version is in Deleted Items.
  - `delete_messages` description: also used to remove a draft after creating its replacement.
- Update requirements §11.1, architecture (`writes.py`, `ows_mail.py`, tool table) and README.

**Accepted costs:** the draft id changes on every change; body edits the user made to the draft in
Outlook, and attachments added there, are not carried into the replacement (the old draft remains
in Deleted Items). No version check guards the replacement: a user who edits a draft by hand knows
what they are doing, and an agent may check the draft with the read tools first if it chooses.
Every change, including recipient- or subject-only ones, is delete-and-recreate: one simple,
predictable rule.

## W8 — Native signatures

**Status:** Decided, not built (scope confirmed 2026-10-07). Depends on W10.

**Scope: native Outlook signatures only.** The connector reads and writes the signatures stored in
the user's Outlook settings (roaming signatures, Cloud Settings adapter; contracts and live evidence
in [research §4.7](outlook-api-research.md#47-native-roaming-signature-discovery-and-standalone-reads-2026-10-07)
and [§4.8](outlook-api-research.md#48-native-signature-crud-and-default-lifecycle-2026-10-07)).
Add-in-generated signatures are ignored in code: see *Corporate signature add-in* below. A user who
wants agent drafts to carry the corporate signature makes a native signature equal to it.

**MCP tools:**

- `list_signatures()`: names, which one is the new-message default and which the reply/forward
  default (either may be none), and whether each listed name has readable contents.
- `get_signature(name)`: HTML and text contents.
- `create_signature(name, html)`, `update_signature(name, html)`, `delete_signature(name)`:
  native signature CRUD (proven contracts, research §4.8). One write each, never retried; a
  rejected write raises. The agent may read back to confirm if it wants.
- `set_default_signature(name | none, for)`: `for` is **required**: `new`, `reply` or `both`.
  Applies to Outlook itself, nothing stored locally. One write, never retried; a rejection raises.

**Draft creation (`create_draft`):**

- By default, insert the right native default, read fresh: the new-message default for new mail, the
  reply/forward default for replies. An explicitly empty default inserts nothing.
- Optional `signature` (name), described as rarely needed ("omit to use your Outlook default"),
  and `include_signature: bool = True`. With `include_signature=false` no signature is inserted;
  a name supplied with it is ignored and noted in the draft's existing findings (if that note would
  need new reporting machinery, ignore it silently).
- A named signature, or a configured default, that does not exist or has no readable contents
  raises. No fallback to another signature.
- Placement once, at creation (W10): the message body, then one `<div id="Signature"
  data-signature-name="{name}">` block (the Outlook Web convention, research §4.6); for replies,
  Exchange appends the quoted history after it.
- Native signature HTML carries images as `data:` URIs: convert them into inline attachments with
  matching `cid:` references. Text-only contents are escaped like `text_body` (drafts are HTML).
- "Editing" a draft is delete-and-recreate (W10); the replacement resolves the default again, so a
  default changed in between applies, and a specifically named signature must be passed again.

**Intended workflow — copying a signature from a draft into a native signature:** the user creates
a draft in Outlook (the corporate add-in inserts its signature there), then asks the agent to read
it and save that signature as a native one. The agent reads the draft's HTML (`get_message`,
`body=html`), extracts the `#Signature` block and calls `create_signature`/`update_signature`.
The draft's signature images are `cid:` references to inline attachments, while native signatures
embed images as `data:` URIs: converting them (download the inline attachment, embed it) is part
of this feature, either in `create_signature` or as a documented agent step.

**Implementation constraints from live findings:** a raw name-list entry may have no readable
contents, and deleting a selected signature can leave a dangling default. Cloud Settings normalizes
HTML wrappers, line breaks and id prefixes, so verify visible text and decoded image bytes rather
than byte-identical HTML. Names are the identifiers (no stable ids): check them against a fresh
list.

**Research first (live, before building the tools):** with throwaway signatures only, never the
user's real ones, and recorded in research §4.8:

1. **Rename:** how Outlook renames a signature (same entry under a new name, or delete and
   re-create), and whether the new-message/reply defaults follow the new name or dangle.
2. **Case:** whether names are case-sensitive, and whether two names differing only in case can
   coexist.
3. **Commas and special characters:** whether a name containing a comma is accepted (the name list
   is comma-joined), and how names with spaces, accents and other characters round-trip through
   the list, contents and default settings.
4. **Images:** that a signature created with a `data:` URI image (converted from a draft's `cid:`
   inline attachment) reads back intact and renders when inserted into a new draft.

Let the results shape name validation in `create_signature`/`update_signature` and whether
`update_signature` supports renaming at all.

**Corporate signature add-in (documented, not handled in code):** in this tenant the official
signature is generated per message by the organization-deployed, mandatory officeatwork "Mail
Signature" add-in (rollout July 2026), from the user's profile and add-in settings, and differs for
internal and external recipients. It inserts into the message being composed; it does not write
native signature settings, and it does not touch drafts created by the connector. Details in
research §4.6.

**Remaining validation:** synthetic fixtures for missing references, no default, missing default
contents, name encoding, configuration changes and account ownership; then public-tool behavior,
draft image attachments, new-message/reply placement and recipient rendering.

# Correctness and reliability

## H19 — Inline images can disappear from exports without an error

**Status:** Pending.

**Problem:** inline images are exported only when the HTML body references their `cid:`. If the body or content-id lookup fails, an image can currently be treated as unreferenced and silently skipped.

**Current evidence:** Graph can return each file attachment's `contentId` directly in the attachment-list request, so the separate per-inline-image content-id lookup is unnecessary.

**Check first (live):** the evidence above is not yet confirmed live in the research record. Before
changing the code, verify that the normal attachment listing returns each file attachment's
`contentId` (for example via `$select` with the `microsoft.graph.fileAttachment/contentId` cast) for
messages with inline images, including in a `$batch`. Record the result in the research doc. If it
does not work, keep the existing per-image lookup and apply only the fail-open rule.

**Next:** return `contentId` from the normal attachment listing, remove the extra lookup, and use a fail-open rule: if the connector cannot determine whether an inline image is referenced, include it rather than silently dropping it.

## H22 — A scope setting alone can turn an export into a whole-mailbox export

**Status:** Pending.

**Terminology (decided 2026-10-08).** "Range" mixed different things; no umbrella term is used.

- **Selections**, which choose messages and add up: **conversations** (`conversation_ids`),
  **messages** (`message_ids`), a **folder** (`folder`) and a **date window** (`since`/`until`). A
  folder and a date window combine: the messages `list_messages` would return for them.
- **Scope** (W11: `scope.sent_items`, `scope.meeting_mail`, `scope.deleted_items`), which never selects
  anything. It narrows what a folder or date window selects; it does not create a selection and does
  not narrow conversations or explicit messages (conversations stay whole).

**Root cause:** `ExportRequest.by_range` decides whether the folder/date selection is active, but it
returns true not only for `since`, `until` or `folder`, but also when `include_sent_items=false` or
`include_meeting_mail=false`. `Exports._select()` adds conversation messages, then calls
`_select_range()` whenever `by_range` is true, then adds explicit messages. With no real folder or date
selector, `_select_range()` calls `list_messages(folder=None, since=None, until=None, ...)`, an
unbounded reachable-mailbox listing.

**Example:** `export_messages(conversation_ids=["budget-thread"], include_meeting_mail=false)` should
export that conversation. Today it also lists the whole reachable mailbox without meeting mail and
merges it in; on a large mailbox it can hit the 2,000-message cap before anything is exported.

**Next:**

- Only `folder`, `since` and `until` select. A request with only scope settings selects nothing and is
  refused: an export with only `scope.sent_items=false`, which today exports the newest received mail
  of the whole mailbox, is no longer a selection.
- Rename with the fix (no aliases, per the compatibility policy):

  | Today | New |
  |---|---|
  | `ExportRequest.by_range` | `ExportRequest.selects_folder_or_dates` |
  | `Exports._select_range()` | `Exports._select_folder_or_dates()` |
  | "range export" (MCP instructions, README, requirements, architecture) | "folder or date-window export" |
  | Error "…or a range (since/until/folder/include_sent_items=false)" | "Select conversations, messages, a folder, or a date window (since/until)." |
  | Requirements §10.1 "a **range** (`since`/`until`, optional `folder`, `include_sent_items`)" | "a **folder** and/or a **date window** (`since`/`until`), narrowed by the scope" |

- The web UI is unaffected: "export this view" already requires a folder or a date range.

## H26 — Oversized attachments are classified as unexpected failures

**Status:** Pending.

**Problem:** the connector's 150 MB download guard is a known local limit, but exports currently classify it like an unexpected error.

**Why 150 MB:** keep the existing 150 MB limit as a connector policy. The code currently defines `MAX_DOWNLOAD_BYTES = 150 * 1024 * 1024` and streams attachment downloads to disk through that guard. This is not a Microsoft Graph download requirement; 150 MB is instead a conservative local ceiling that also lines up with upper Outlook/Exchange attachment/message limits in some Microsoft contexts. Most mail providers impose much tighter practical limits, so a normal file attachment reaching this guard should be extremely rare. Retaining it protects against unexpectedly huge downloads/disk usage without materially constraining ordinary email export.

**Next:** classify hitting the guard explicitly as a known connector limit rather than an unexpected failure: the attachment is larger than the connector's 150 MB download limit and should be downloaded directly from Outlook.

# Performance

## H30 — Per-folder listing reads busy folders in several rounds

**Status:** Pending.

**Problem:** the per-folder merge starts every folder with a small chunk and doubles it as needed. A mailbox dominated by one busy folder can therefore read that folder two or three times sequentially for one page.

**Current evidence:** the count batch can return both each folder's message count and its newest message date in the same request.

**Next:** use the counts to size each folder's first read according to its expected share of the page, and use the newest date when useful to avoid reading folders that cannot contribute to the current merge. Aim for one read per contributing folder in the common case.

## H31 — `include_total` can send the same count batch twice

**Status:** Pending.

**Problem:** the first page of a per-folder listing counts folders to choose/list them, then `include_total` asks `_count` for effectively the same data again.

**Next:** carry the first count result through the call and reuse it for `server_total` and H27's excluded counts.

## H32 — Conversation selections are expanded sequentially

**Status:** Pending.

**Problem:** exports that select many conversations, and conversation-based read-state changes, list each conversation one after another.

**Next:** expand conversations concurrently under the transport's existing four-request mailbox limit, or use a Graph batch where the response shape remains simple.

# Service and code cleanup

## H34 — Work is repeated inside a single call

**Status:** Pending.

**Current duplication worth removing:**

- mailbox scope/folder categories are recomputed several times during one listing or related operation;
- conversation-based `set_read_state` lists summaries and then fetches them again;
- reply draft/send verification rereads messages already fetched earlier in the same call;
- export selections can run already-finished messages through scope/merge work again;
- `save_message_mime` checks the message before downloading MIME even though the MIME headers can supply the subject.

**Next:** compute/validate once at the public service entry point and pass the resulting scope/data down to helpers that trust it.

## H35 — Limits are validated in more than one layer

**Status:** Pending.

**Problem:** MCP signatures and service methods both enforce several numeric bounds, sometimes with different minimums.

**Next:** make the service authoritative for validation because both MCP and web call it. Surface schemas/descriptions should document limits but not enforce a second, different rule.

## H36 — Derived counts are stored beside the data they count

**Status:** Pending.

**Problem:** several response models store both a list and a count that is always derivable from that list, creating two values that must stay synchronized.

**Next:** remove redundant derived counts where callers can count the data directly. Keep `MutationResult.counts`, because tallying a large per-message mutation result is useful to an agent.

## H37 — Repeated logic still has multiple sources of truth

**Status:** Pending.

**Main cleanup targets:**

- one canonical message timestamp (`received_at` or `sent_at`);
- one failure/detail formatter and one fetched-failure shape;
- one shared authentication/retry loop for normal requests and downloads;
- one authority for URL/host validation;
- remove duplicate recipient de-duplication;
- one definition of the UI's default port/idle timeout;
- use the production id helpers in the fake mailbox;
- return attachment `contentId` from the normal listing instead of a separate lookup (also simplifies H19).

Recent constructor/protocol cleanup reduced unrelated duplication, but these review targets remain.

## H38 — Small dead-code cleanup

**Status:** Pending, low priority.

**Current state:** the recent constructor cleanup removed some unused test seams, but the review still has a handful of small leftovers to remove or retype after H37 so the same code is not churned twice. Known candidates include unused attachment/content-id representation, an exclusion-reason type that is not actually used as a type, never-used optional parameters on Graph helpers, an impossible HTML fallback in export code, and stale comments.

**Next:** do one dead-code pass after the structural cleanup and delete only what the current test suite proves unused. Keep the account-ownership checks intentionally.

# Documentation

## W12 — Landing README rewrite and architecture doc cleanup

**Status:** Decided, not built (2026-10-08). Last: do it after every item above is finished, so it
describes the final tool set.

**README, in this order:**

1. **What it is, in a nutshell.** One short sentence ("a connector that lets you connect to your
   Outlook account and…") and a short, direct bullet list that makes the reader want to know more, e.g.:
   search messages by subject and contents; export individual messages or entire conversations or
   folders, *attachments included*; draft, send or reply to emails; move messages between folders;
   manage the rules that file or delete mail automatically; … (match the final tool set).
2. **A diagram (SVG), marketing in spirit,** like `docs/architecture.svg` but for readers, not
   developers: MCP clients on the left (Claude, Codex/ChatGPT, "any other MCP client") and the local
   web UI (a small screenshot, captioned e.g. "custom UI for browsing and downloading conversations"),
   all with arrows into the connector, which connects to Outlook. Product names; use logos only where
   their usage is permitted. Note: ChatGPT (web/desktop) only connects to remote MCP servers over
   HTTP, so it cannot use this local stdio server today; show what actually works (Claude Code,
   Claude Desktop, Codex, other local MCP clients) or mark ChatGPT accordingly.
3. **Setup:** `uv sync`, then the sign-in commands. A `>` note: credentials are entered only on
   Microsoft's sign-in page, never seen by the application; tokens are kept in the encrypted local
   cache. Then the UI command, with a larger version of the UI screenshot.
4. **MCP tools:** every tool by name with a one-line description; no arguments. A short hint where it
   helps (e.g. "filters such as folder, dates and scope"), nothing longer.
5. **Development:** how to run the checks and tests (as today), plus the links to roadmap,
   requirements, architecture and API research, moved here from the top.

Keep the command cheatsheet style that already works; cut the verbose "Use it" prose. Details that
leave the README belong in the requirements/architecture docs if they are not already there.

**docs/architecture.md:**

- Under "4. Repository layout", embed `docs/code-map.svg`.
- Describe the current state only: drop history such as "removed 2026-10-04", "decided 2026-10-04"
  and "(the summary cache was removed…)". Decisions and dates stay in the research record.
- Every file or module mentioned (e.g. "5.1 `auth/tokens.py`") links to that file with a relative link.
- Remove "11. Build order" (renumber the following section) and repoint the roadmap's
  "Build/module map" link, which targets it, to "4. Repository layout".

# Later

## X11 — SharePoint file search and download

**Status:** Candidate; read-only feasibility proven 2026-10-07, no product tool implemented.

**Evidence:** the existing read sign-in can search for and begin downloading a known SharePoint
file; unrestricted file search and known-person search also worked. Limits and unproven cases
are recorded once in [research §3.7](outlook-api-research.md#37-file-and-cross-category-search-feasibility-2026-10-07).

**Next:** a separate Graph files adapter and service port, exposing bounded file search with
continuation and read-only download by URL or a returned file reference. Validate input once,
preserve account ownership, bound downloads, use exclusive output creation and handle
preauthenticated redirects without leaking credentials. Keep mail
attachment downloads distinct. Broader sharing-link support and a capability reusable by the
Teams connector remain design decisions; the Teams sign-in was not repeated in these probes.

Keep each category's coverage and permissions explicit if a future broad discovery surface
orchestrates these providers. Calendar authentication remains a separate blocker (X10).

## X10 — Meeting/calendar search

**Status:** Later; blocked on authentication scope.

**Goal:** search Exchange calendar events (subject, time, organizer, attendees, agenda and Teams join link) and hand a Teams meeting's chat id to `lrh-teams` when useful.

**Current blocker:** the read profile has no `Calendars.*` scope, so Graph `/me/calendarView` is not available through the current read sign-in.

**Next when revisited:** find or prove a sign-in/client route that grants `Calendars.Read`; the existing write token is a candidate to test. Keep calendar/event data in this connector and meeting chat in `lrh-teams`.

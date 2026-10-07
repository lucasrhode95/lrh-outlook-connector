# Outlook Connector Roadmap

Outstanding decisions, implementation, fixes and validation only. Completed behavior belongs in
the requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-07**.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §11](architecture.md)
- API evidence: [API research](outlook-api-research.md)

## Current priority

1. **W9:** fix inactive rule values/description metadata and create-response identity reconciliation,
   then repeat the public-tool lifecycle checks.
2. **W8:** implement signature listing/content retrieval and fresh default resolution, including
   missing signatures and dangling defaults.
3. **Correctness:** H19, H22 and H26.
4. **Performance and cleanup:** H30–H32 and H34–H38.

Status wording:

- **Decided, not built** — behavior is settled; implementation remains.
- **Pending** — a concrete problem and next fix are known.
- **Decision needed** — an owner choice remains.
- **Later** — useful work outside the current build sequence.

# Send and mailbox features

## W8 — Signatures

**Status:** Decided, not built. Public tools and automatic draft integration remain.

**Goal:** retrieve the user's current native Outlook default automatically when composing a new
message. Resolve the reply/forward default separately, and list/retrieve configured signatures
reliably. No user-supplied file or frozen imported copy is required.

**Backend:** Graph first. Microsoft's [roaming-signature documentation](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-organizationconfig?view=exchange-ps#-postponeroamingsignaturesuntillater)
identifies a Graph capability gap; use the proven Outlook Cloud Settings adapter for native
configuration and Graph for draft/attachment reads. API contracts and live evidence are in
[research §4.7](outlook-api-research.md#47-native-roaming-signature-discovery-and-standalone-reads-2026-10-07)
and [§4.8](outlook-api-research.md#48-native-signature-crud-and-default-lifecycle-2026-10-07).

**Remaining implementation:**

- Read-only MCP tools `list_signatures` and `get_signature`, with default selections, content
  availability and missing-reference reporting. Signature-management writes are a separate scope
  choice from this initial read/default-composition feature; no editor is needed.
- A signature-reader port and Cloud Settings adapter; resolution and placement belong in the service.
- Fresh default/content queries during composition, without a persisted signature/default cache.
  Check account ownership and the relevant revision after retrieval. Respect an explicit no-default
  setting; report missing configured content or a failed/changed read rather than selecting another
  signature or using a stale copy.
- Placement: new body, then signature; for replies, retain Outlook's quoted history after the
  signature. Body edits must replace the connector-managed signature rather than duplicate it.
  Select the signature before saving the draft; sending preserves the already saved draft.
- Convert embedded image data into actual inline attachments with matching CID references.

**Implementation constraints from live findings:** a raw name-list entry may have no readable
contents, and deleting a selected signature can leave a dangling default. Resolve and verify the
selected contents. Cloud Settings normalizes HTML wrappers, line breaks and id prefixes, so verify
visible text and decoded image bytes rather than byte-identical HTML.

**Corporate add-in constraint:** native resolution does not execute officeatwork's sender/recipient,
language or template policies. A rendered-draft snapshot does not establish fresh corporate policy
evaluation; see [research §4.6](outlook-api-research.md#46-outlook-web-signature-capture-2026-10-07).

**Remaining validation:** synthetic fixtures for missing references, no default, missing default
contents, name encoding, configuration changes and account ownership; then public-tool behavior,
draft image attachments, new-message/reply placement and recipient rendering.

## W9 — Inbox-rule fixes and public-tool validation

**Status:** Pending; two adapter fixes and the public-tool lifecycle checks remain.

**Current blockers:**

- Description metadata and inactive OWS enum values are classified as unsupported behavior,
  making supported rules read-only. Recognize the proven per-field values while retaining protection
  for active unsupported conditions/actions.
- `NewInboxRule` can return an identity that differs from fresh `GetInboxRule` read-back, so
  creation reports failure even when the rule was created. Reconcile against the fresh collection.

**Remaining work:** add synthetic fixtures for those wire shapes, fix both cases, then validate
edit/condition clearing, disable/enable, delete and reorder through the public tools. Preserve
state-bound confirmation, a single write and fresh read-back; never rewrite unsupported rules.
The specific fields and live evidence are in [research §4.5](outlook-api-research.md#45-connector-live-validation-2026-10-07).

# Correctness and reliability

## H19 — Inline images can disappear from exports without an error

**Status:** Pending.

**Problem:** inline images are exported only when the HTML body references their `cid:`. If the body or content-id lookup fails, an image can currently be treated as unreferenced and silently skipped.

**Current evidence:** Graph can return each file attachment's `contentId` directly in the attachment-list request, so the separate per-inline-image content-id lookup is unnecessary.

**Next:** return `contentId` from the normal attachment listing, remove the extra lookup, and use a fail-open rule: if the connector cannot determine whether an inline image is referenced, include it rather than silently dropping it.

## H22 — A filter flag alone can turn an export into a whole-mailbox range

**Status:** Pending.

**Root cause:** export has three additive selection sources: conversation ids, an optional mailbox range, and explicit message ids. `ExportRequest.by_range` is supposed to say whether the mailbox-range source is active, but it currently returns true not only for `since`, `until` or `folder`, but also when `include_sent_items=false` or `include_meeting_mail=false`.

`Exports._select()` first adds messages selected by conversation id, then calls `_select_range()` whenever `by_range` is true, then adds explicit message ids. These sources are merged; the explicit selection is not being ignored. The bug is that a filter flag accidentally activates an additional range selection. If no real range selector was supplied, `_select_range()` calls `list_messages(folder=None, since=None, until=None, ...)`, which is effectively an unbounded reachable-mailbox listing subject to the include filters.

**Example:** `export_messages(conversation_ids=["budget-thread"], include_meeting_mail=false)` should export that conversation. Today, `include_meeting_mail=false` also makes `by_range=true`, so the exporter additionally lists the whole reachable mailbox with meeting mail excluded and merges those messages with the requested conversation. On a large mailbox this can even hit the 2,000-message export cap before anything is exported.

**Next:** only `since`, `until` and `folder` activate the mailbox-range selection source. `include_meeting_mail` and `include_sent_items` remain filters for a range when one is actually requested; they must never create a range on their own or broaden an explicit conversation/message selection.

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

# Later

## X10 — Meeting/calendar search

**Status:** Later; blocked on authentication scope.

**Goal:** search Exchange calendar events (subject, time, organizer, attendees, agenda and Teams join link) and hand a Teams meeting's chat id to `lrh-teams` when useful.

**Current blocker:** the read profile has no `Calendars.*` scope, so Graph `/me/calendarView` is not available through the current read sign-in.

**Next when revisited:** find or prove a sign-in/client route that grants `Calendars.Read`; the existing write token is a candidate to test. Keep calendar/event data in this connector and meeting chat in `lrh-teams`.

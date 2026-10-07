# Outlook Connector Roadmap

Current work only. This is a decision and implementation register, not a changelog: completed work and intentionally parked ideas are omitted.

Snapshot **2026-10-07**, against `main`.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §11](architecture.md)
- API evidence: [API research](outlook-api-research.md)

## Implementation baseline

The implementation batch and review fixes were merged on 2026-10-06 in PRs #20–#27 and
[#29](https://github.com/lucasrhode95/lrh-outlook-connector/pull/29). Synthetic checks passed.
Live evidence and remaining gaps are recorded below and in [API research §4.5](outlook-api-research.md#45-connector-live-validation-2026-10-07).

**Live-test cleanup (2026-10-07):** all 105 temporary drafts were moved to Deleted Items and verified
there; none remain in Drafts. Initial read states were restored, and the eight original inbox rules
were restored exactly. Four sent test messages and their four received self-copies remain as
evidence. No permanent deletion was performed.

## Current priority

1. **W9 live blockers:** distinguish inactive OWS enum values/description metadata from unsupported
   rule behavior, and reconcile create-response identities with fresh rule-list identities; then
   repeat the tool lifecycle checks.
2. **W8:** signatures; Web capture reveals client insertion, corporate templates and inline uploads.
   Prove an app-owned extraction/reuse route before implementing it.
3. **Correctness:** H19, H22 and H26.
4. **Performance and cleanup:** H30–H38 (H33 is already implemented).
5. **Deferred live research:** HTML/client rendering; calendar authentication remains Later (X10).

Status wording used below:

- **Decided, not built** — behavior is settled; implementation remains.
- **Pending** — a concrete problem and next fix are known.
- **Decision needed** — there is still an owner choice to make.
- **Later** — useful, but not part of the current build sequence.

# Send and mailbox features

## W7 — Draft-first text and HTML sending

**Status:** Implemented (merged 2026-10-06); text drafts, edits and version-bound sending validated
live on 2026-10-06. HTML drafts and replies exercised on 2026-10-07; the owner confirmed correct
HTML rendering of the received test emails. Broader recipient-client research remains deferred.

**Goal:** make Outlook drafts the only entry point for agent-authored outgoing mail. The connector must never guess whether a body is plain text or HTML, and it must never reconstruct a message at send time.

**Public write surface:**

- `create_draft(..., text_body=..., html_body=...)`
- `edit_draft(draft_id, ..., text_body=..., html_body=...)`
- `send_draft(draft_id)`

`create_draft` accepts exactly one of `text_body` or `html_body`. `edit_draft` uses the same explicit body fields when the body is changed; neither is required for recipient/subject-only edits. New messages and replies use the same tools, with the existing reply context (`reply_to_message_id`, `reply_all`) selecting reply behavior.

**Body handling:**

- Plain text is escaped into minimal HTML before entering the shared private write path. Preserve line breaks, indentation, repeated spaces/tabs and especially non-breaking-space semantics; do not parse Markdown, auto-link, inject fonts or make formatting decisions.
- HTML is intentional pass-through for now. Do not distinguish fragments from full HTML documents, normalize/repair markup, restrict CSS, or reject remote images solely for being remote.
- Reject only clearly active/web-application content such as scripts, iframes/objects/embeds, forms, JavaScript URLs and event-handler attributes.
- The connector has one private HTML write engine after the public text/HTML boundary.

**Draft-first flow:**

1. Validate recipients, subject/body and reply arguments locally.
2. Create or update the Outlook draft through OWS.
3. Read that server draft back through Graph.
4. Return the Microsoft draft id plus both server text and server HTML, along with simple verification findings.
5. The agent shows/uses that read-back in the conversation. If the user later explicitly asks to send it, `send_draft(draft_id)` sends that existing Microsoft draft exactly as stored.

There is no `propose_email`, direct `send_email`, format flag, format guessing, connector-side draft registry or custom confirmation token. Human approval remains the agent/host interaction before `send_draft`; the connector's safety property is that the send operation accepts only an existing Microsoft draft id and cannot alter/reconstruct the message while sending.

**Read-back and checks:** successful create/edit requires Graph read-back. A few bounded retries for normal propagation/transient failures are fine; if read-back ultimately fails, report the operation as failed and return any reliable server id/details already obtained. Do not create another draft automatically or maintain local recovery state.

For the first version, automatic verification stays deliberately simple and high-confidence: flag an empty/missing body, missing expected attachments/inline images where cheaply knowable, and preserve the existing reply-history checks. Do not compare submitted HTML with Microsoft's returned HTML or implement a generalized body diff.

**Deferred live-test research, not blockers:** after the basic flow works, test how Outlook/Graph and recipient clients treat fragments vs complete HTML documents, malformed-but-accepted HTML, CSS, remote images and representative formatting. Revisit body-change detection only with real examples; a future approach may compare normalized visible text (entities decoded, whitespace/non-breaking spaces normalized) rather than HTML structure.

**Live checks (2026-10-06, self-sends and probe drafts):** text draft creation, Cc/Bcc clearing with an empty list, subject edit, and existing-draft send with no field updates bound to the draft's change key (sent when current; refused with nothing sent when the draft changed since it was read). Two probe drafts were also moved to Deleted Items through `delete_messages` (`done: 2`). See research §4.2.

**Live checks (2026-10-07, owner-authorized self-sends and personal recipients):** HTML creation,
body edit and existing-draft send passed. Graph read-back of the saved draft, sent copy and received
self-copy preserved tables, lists, links, line breaks and non-ASCII text. Text replies and HTML
reply-all drafts passed the connector's history checks, including quoted structure and inline-image
bytes; default recipients matched the expected reply behavior. Recipient-only edits retained the
inline attachment, and both replies were sent once. See research §4.5 for received-copy checks and
the limits of this validation.

**Owner confirmation (2026-10-07):** the received test emails looked correct and HTML rendered
properly. The individual recipient/client combinations were not specified. Automated recipient-side
inspection was unavailable (Gmail connector unauthenticated, browser automation initialization failed).

**Next:** broader HTML/client research remains: full documents, malformed accepted HTML, CSS,
remote images and a client-specific comparison matrix.

## W8 — Signatures

**Status:** Unblocked; server-draft probes and Web capture analysis completed on 2026-10-07.
Client insertion and a corporate template source are identified; app-owned reuse and implementation remain.

**Retrieval goal:** reuse the user's configured signature HTML and inline images without requiring
an Outlook `.htm` export. The server-created minimal-draft route was tested below; Web capture now
provides the next discovery routes.

**Live finding (2026-10-07):** minimal text, repeated minimal HTML and empty HTML server-created
drafts returned only submitted content (or an empty body), with no extra signature text or images.
The empty draft correctly returned the empty-body verification finding. This draft-creation route
did not expose a reusable signature; the run did not inspect the user's configured signature in the
Outlook UI or exhaust other settings APIs.

**Capture finding (2026-10-07):** the supplied Web HAR contains client-submitted signature HTML,
four successful inline uploads and a draft save. The officeatwork Mail Signature add-in retrieves a
SharePoint `.ofawmsig` package (Nunjucks template, metadata and images), profile/group data and extra
image assets. Its loaded source renders those inputs and uses Office.js `setSignatureAsync` plus
inline attachment insertion. A selected template pointer is persisted in add-in extension settings.
This is sufficient evidence of client-supplied signatures for the captured draft; the server-only
probe cannot trigger this compose flow. See [research §4.6](outlook-api-research.md#46-outlook-web-signature-capture-2026-10-07)
for endpoint shapes, source/traffic distinctions and authentication limits.

**Extraction finding:** the saved signature block and four matching PNGs were recovered privately
from the capture. Final OWS `NormalizedBody` substitutes four 42-byte GIF placeholders; use actual
attachments and CID HTML for extraction. This is one static compose snapshot, not automatic reuse.

**Native settings gap:** the owner reports adding, editing and deleting a signature. The capture's
106-second window does not expose identifiable native signature CRUD requests; the two configuration
PATCH writes contain Sales add-in preferences. Do not count them as signature CRUD validation.

**Next:** test read-only discovery with the connector's own authentication: the selected template
reference and assets, or a user-named Web-created draft and its real inline attachments. Prefer
extracting a rendered draft if that avoids implementing corporate template/policy behavior. Prove
inline upload and new-message/reply placement with synthetic fixtures, then implement the chosen
route. A settings-page load plus add/edit/delete capture is still needed for native signature API
discovery. Keep import as fallback; no app-owned automatic extraction/reuse has been demonstrated.

**Fallback scope:** exactly one signature for the connector: import it, remove it, enable it or disable it. No multiple-signature selection, routing rules or signature editor.

**Fallback build:**

- MCP tools: `import_signature`, `delete_signature`, `set_signature_active(true|false)`.
- Import an Outlook-style `.htm` signature and optional `<name>_files/` image folder.
- Store the one signature under the connector data directory; active by default after import.
- Send signature images as inline attachments referenced by `cid:`.
- Placement matches Outlook: new message = body then signature; reply = new text, signature, then Outlook's quoted history untouched.
- MCP only for now; no UI work required.

## W9 — Inbox rules

**Status:** Implemented (merged 2026-10-06); synthetic tool validation complete, live tool validation
on 2026-10-07 exposed blockers.

**API evidence (2026-10-04):** the Outlook Web rule contracts were captured and replayed successfully through `Ows.call_request`. Reading, creating, editing (including clearing a condition), enabling/disabling, reordering and deleting a throwaway rule all worked live at that layer. Graph folder ids are accepted by the write request. The public-tool blockers found on 2026-10-07 are recorded below.

**First-version scope:**

- `list_rules`
- `create_rule`
- `update_rule`, including enable/disable
- `reorder_rules`
- `delete_rule`

Support only the conditions/actions already proven and used by the mailbox: From, Sent to, Subject contains, Subject-or-body contains, Move to folder, and Stop processing more rules. Unsupported rules are listed read-only and are never rewritten.

**Safeguard:** every rule write is proposed first, requires user confirmation, is sent once, and is read back. Rules persist and affect future mail, so writes must not be retried automatically.

**Implementation notes:** stateless account/state-bound proposals, one confirmed write and fresh read-back. Enable/disable is a separate update; reordering refuses collections containing unsupported rules to avoid rewriting them. Folder read-back uses the reported name (OWS returns a mailbox path rather than a Graph id).

**Live findings (2026-10-07):** all eight existing rules and a newly created throwaway rule were
classified read-only because description metadata and inactive enum strings such as `NullImportance`
and `NullInboxRuleMessageType` were treated as unsupported behavior. Creation wrote once, but
reported `failed`: its returned identity differed from the new rule's identity in `GetInboxRule`.
The fresh list confirmed the intended conditions and Archive destination. Update was refused as
read-only; stale create confirmation and reordering the unsupported collection were rejected.
The throwaway rule was removed once through the underlying adapter for cleanup; all eight original
rule ids, order, priorities, enabled states and revisions were restored.

**Next:** recognize only the proven per-field inactive enum values and description metadata (keep
active unsupported settings protected); resolve create identities against fresh read-back. Add
synthetic fixtures for these shapes, then repeat edit/clear, disable/enable, delete and reorder
through the public tools. Their complete live lifecycle is not yet validated. See research §4.5.

# Correctness and reliability

## H17 — Partial mutation failures are reported incorrectly

**Status:** Implemented (merged 2026-10-06); synthetic validation complete, controlled live
read-state failure checks passed on 2026-10-07.

**Problem:** mutations are sent in chunks. If a later chunk fails, earlier chunks may already have changed the mailbox while the tool call raises only the later error.

**Decision:** add `continue_on_error` to `set_read_state`, `set_flag`, `move_messages` and `delete_messages`, defaulting to `true`.

- A clearly failed chunk returns `failed` results for its messages.
- By default, later chunks are still attempted.
- With `continue_on_error=false`, later chunks are not sent and are returned as `failed: not sent`.
- A failed read-back after an unclear write outcome returns `unknown`.
- The call always returns per-message results and `counts`, so already-completed work is visible.

**Controlled live checks (2026-10-07):** MCP calls on 60 dedicated drafts used real OWS writes and
local failure injection before the second 20-item chunk. Default continuation returned `done: 40,
failed: 20`; `continue_on_error=false` returned `done: 20, failed: 40`, including 20 explicitly
`not sent`. Independent Graph reads matched both outcomes. A real 20-item write followed by injected
response loss and read-back failure returned `unknown: 20`; an independent read then confirmed the
actual changes. Initial states were restored. These prove the reporting flow against a real mailbox;
they do not represent naturally occurring Microsoft outages, and move/delete/flag failure variants
remain synthetic coverage.

## H19 — Inline images can disappear from exports without an error

**Status:** Pending.

**Problem:** inline images are exported only when the HTML body references their `cid:`. If the body or content-id lookup fails, an image can currently be treated as unreferenced and silently skipped.

**Current evidence:** Graph can return each file attachment's `contentId` directly in the attachment-list request, so the separate per-inline-image content-id lookup is unnecessary.

**Next:** return `contentId` from the normal attachment listing, remove the extra lookup, and use a fail-open rule: if the connector cannot determine whether an inline image is referenced, include it rather than silently dropping it.

## H20 — Explicit message ids can be filtered out during export

**Status:** Recommended decision implemented (merged 2026-10-06); synthetic validation complete.

**Problem:** `export_messages(message_ids=[...])` reads the requested messages, then applies reach/scope filtering at the end. An explicitly supplied id from a hidden/out-of-reach folder can therefore disappear from the export without the selection behaving like `get_message(id)`.

**Recommended decision:** treat explicit message ids as authoritative, like `get_message` does. Do not apply the hidden-folder filter to that selection source; only merge/label it with the other selected messages.

**Alternative:** keep the filter but preserve and report its exclusion counts. This is stricter, but adds complexity for ids the connector itself never normally exposes.

## H21 — Long conversations cannot be marked read/unread cleanly

**Status:** Implemented (merged 2026-10-06); synthetic validation complete, bulk read-state behavior
validated live on 2026-10-07.

**Problem:** `set_read_state(conversation_ids=[...])` expands a conversation to messages and then applies the 100-message limit intended for explicit `message_ids`. A long conversation can therefore be rejected even though the caller supplied only one conversation id. Conversations beyond the 1,000-message listing cap also need an explicit truncation note.

**Implemented:** apply the 100 limit only to explicit `message_ids`. Expand conversation ids up to the existing conversation-listing limit, send changes in normal 20-item write chunks, and report when a conversation was truncated. Keep the response compact for very large conversations by relying on `counts` for ordinary `done`/`unchanged` results if needed.

**Live check (2026-10-07):** 101 dedicated unsent reply drafts plus the original's sent and received
copies formed a 103-message conversation. The MCP `set_read_state` tool marked all 103 unread and
then read (`done: 103` each); fresh Graph summaries independently confirmed every state. Both
responses had zero ordinary detailed results and one compact-results note. Every initial read state
was restored. The 1,000-message truncation boundary remains covered synthetically, not live.

## H22 — A filter flag alone can turn an export into a whole-mailbox range

**Status:** Pending.

**Root cause:** export has three additive selection sources: conversation ids, an optional mailbox range, and explicit message ids. `ExportRequest.by_range` is supposed to say whether the mailbox-range source is active, but it currently returns true not only for `since`, `until` or `folder`, but also when `include_sent_items=false` or `include_meeting_mail=false`.

`Exports._select()` first adds messages selected by conversation id, then calls `_select_range()` whenever `by_range` is true, then adds explicit message ids. These sources are merged; the explicit selection is not being ignored. The bug is that a filter flag accidentally activates an additional range selection. If no real range selector was supplied, `_select_range()` calls `list_messages(folder=None, since=None, until=None, ...)`, which is effectively an unbounded reachable-mailbox listing subject to the include filters.

**Example:** `export_messages(conversation_ids=["budget-thread"], include_meeting_mail=false)` should export that conversation. Today, `include_meeting_mail=false` also makes `by_range=true`, so the exporter additionally lists the whole reachable mailbox with meeting mail excluded and merges those messages with the requested conversation. On a large mailbox this can even hit the 2,000-message export cap before anything is exported.

**Next:** only `since`, `until` and `folder` activate the mailbox-range selection source. `include_meeting_mail` and `include_sent_items` remain filters for a range when one is actually requested; they must never create a range on their own or broaden an explicit conversation/message selection.

## H23 — Search dates without a timezone can crash

**Status:** Implemented (merged 2026-10-06); synthetic validation complete.

**Problem:** the web API can pass a naive `since`/`until` datetime into service code, which is then compared with timezone-aware Graph dates.

**Next:** normalize dates to UTC once at the service entry point and remove duplicate surface-specific normalization.

## H25 — The UI silently truncates very long message bodies

**Status:** Implemented (merged 2026-10-06); automatic full-body loading validated.

**Problem:** the UI reader asks for a bounded body and ignores `next_offset`, so a message over the current limit can stop mid-body without telling the user.

**Decision for this batch:** load the complete chosen body without a new total length cap. The first version followed `next_offset` in 200,000-character chunks, but each chunk re-read the whole message (and its attachment list) from Graph, so an L-character body cost about L/200,000 full downloads (review of #26). The local web reader now asks for the whole body in one request; MCP keeps its bounded chunks.

## H26 — Oversized attachments are classified as unexpected failures

**Status:** Pending.

**Problem:** the connector's 150 MB download guard is a known local limit, but exports currently classify it like an unexpected error.

**Why 150 MB:** keep the existing 150 MB limit as a connector policy. The code currently defines `MAX_DOWNLOAD_BYTES = 150 * 1024 * 1024` and streams attachment downloads to disk through that guard. This is not a Microsoft Graph download requirement; 150 MB is instead a conservative local ceiling that also lines up with upper Outlook/Exchange attachment/message limits in some Microsoft contexts. Most mail providers impose much tighter practical limits, so a normal file attachment reaching this guard should be extremely rare. Retaining it protects against unexpectedly huge downloads/disk usage without materially constraining ordinary email export.

**Next:** classify hitting the guard explicitly as a known connector limit rather than an unexpected failure: the attachment is larger than the connector's 150 MB download limit and should be downloaded directly from Outlook.

## H27 — Per-folder listing does not report excluded-folder counts

**Status:** Implemented (merged 2026-10-06, #27); synthetic validation complete.

**Problem:** when the mailbox uses the per-folder listing strategy, Junk and Deleted Items are never read. That is efficient, but `coverage.excluded` therefore lacks the `deleted_or_junk` count even though those messages are outside the result.

**Decision:** include the left-out folders in the existing count batch and report their counts for the requested window, with no extra round trip. A failed count drops only that folder: the excluded count is then not claimed, and the in-scope folders keep their fresh counts (review fix, 2026-10-06).

## H28 — `server_total` includes meeting mail that the listing may hide

**Status:** Implemented (merged 2026-10-06, #27); synthetic validation complete.

**Problem:** with `include_meeting_mail=false`, `server_total` still includes invitations, cancellations and RSVPs, so the total can be larger than the messages the listing can return.

**Constraint:** Graph's folder count endpoint cannot filter out meeting-message types with the available query shapes.

**Next:** make the coverage note explicit that `server_total` includes meeting mail even when the listing hides it. Do not add a second expensive counting path.

## H29 — NotFound wording assumes deletion

**Status:** Implemented (merged 2026-10-06, #27); synthetic validation complete.

**Problem:** a bad or inaccessible message id currently produces wording equivalent to "not on the server (deleted on the server)", even though the id may simply be wrong or the item may have moved out of reach.

**Next:** use neutral wording: the message was not found; it may have been deleted, moved out of reach, or the id may be wrong. Remove the mailbox-level `GONE` wording that implies deletion.

## H33 — `$batch` retries fewer transient statuses than single requests

**Status:** Implemented (merged 2026-10-06, #27); synthetic validation complete.

**Problem:** individual requests retry transient 429/502/503/504 responses, but a failing item inside a Graph `$batch` is retried only for throttling. A 503/504 batch item therefore becomes an export gap immediately even though the equivalent single request would retry.

**Next:** use the same transient status set for batch-item retries as for single requests.

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

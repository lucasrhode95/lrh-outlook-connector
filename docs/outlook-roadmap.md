# Outlook Connector Roadmap

Current work only. This is a decision and implementation register, not a changelog: completed work and intentionally parked ideas are omitted.

Snapshot **2026-10-06**, against `main`.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §11](architecture.md)
- API evidence: [API research](outlook-api-research.md)

## Completed in this implementation batch

Implementation branches start independently from `main`; all PRs target `main` and remain unmerged.
The compact record below is retained for this batch's independent review, overriding the normal
omission of completed work. Each implementation branch contains its own affected-doc updates.

| Item | PR | Completion note |
|---|---|---|
| W7 | [#20](https://github.com/lucasrhode95/lrh-outlook-connector/pull/20) | Explicit text/HTML draft creation/edit, mandatory server read-back, existing-draft-only send; synthetic checks complete, live validation pending. |
| W9 | [#21](https://github.com/lucasrhode95/lrh-outlook-connector/pull/21) | Supported inbox-rule MCP tools; state-bound proposal/confirmation, one write/read-back; unsupported rules read-only. |
| H17 | [#22](https://github.com/lucasrhode95/lrh-outlook-connector/pull/22) | Default continuation after chunk errors, per-message outcomes/counts, unknown read-back failures and explicit not-sent results. |
| H20 | [#23](https://github.com/lucasrhode95/lrh-outlook-connector/pull/23) | Explicit export ids remain authoritative; other sources keep scope filters and copies still merge. |
| H21 | [#24](https://github.com/lucasrhode95/lrh-outlook-connector/pull/24) | Explicit-only 100 limit, existing 1,000-per-conversation cap, truncation notes and compact bulk results. |
| H23 | [#25](https://github.com/lucasrhode95/lrh-outlook-connector/pull/25) | Search normalizes naive/aware dates to UTC in the service; redundant MCP search normalization removed. |
| H25 | [#26](https://github.com/lucasrhode95/lrh-outlook-connector/pull/26) | UI automatically follows all body offsets; no total size cap; obsolete continuation scheduling stops on selection change. |
| H27 | [#27](https://github.com/lucasrhode95/lrh-outlook-connector/pull/27) | Deleted/Junk window exclusions from the existing per-folder count batch, reported once across pages. |
| H28 | [#27](https://github.com/lucasrhode95/lrh-outlook-connector/pull/27) | Coverage/docs explicitly say server_total includes meeting mail; no second counting mechanism. |
| H29 | [#27](https://github.com/lucasrhode95/lrh-outlook-connector/pull/27) | Neutral NotFound wording for messages, attachments, MIME and export gaps; mailbox GONE text removed. |
| H33 | [#27](https://github.com/lucasrhode95/lrh-outlook-connector/pull/27) | Batch-item retries share single-read transient statuses (429/502/503/504). |

All implementation PRs passed lint, formatting, type checks and their full Python test suites;
H25 also passed five behavioral JavaScript tests. No real mailbox was changed. W7's no-field-update
OWS send shape and draft/HTML behavior still require explicitly authorized live validation; W9's
folder read-back relies on OWS's reported name. H17/H21 touch the same mutation module but have no
branch dependency; preserve both behaviors when resolving merge conflicts.

## Current priority

1. **W7 → W8:** draft-first text/HTML sending, then signatures.
2. **W9:** inbox-rule MCP tools; the API contracts are already proven live.
3. **Correctness and reliability:** H17, H19–H29 and H33.
4. **Performance and cleanup:** H30–H38.

Status wording used below:

- **Decided, not built** — behavior is settled; implementation remains.
- **Pending** — a concrete problem and next fix are known.
- **Decision needed** — there is still an owner choice to make.
- **Later** — useful, but not part of the current build sequence.

# Send and mailbox features

## W7 — Draft-first text and HTML sending

**Status:** Implemented on this branch; synthetic validation complete, live validation pending.

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

**Next:** on explicit owner request, live-test plain-text/HTML creation, partial edit, replies and existing-draft send. The no-field-update `UpdateItem` send shape needs live confirmation; earlier evidence used a subject update. Deferred HTML/client research remains unchanged.

## W8 — Signatures

**Status:** Postponed until W7 is complete.

**First investigate server-side signature retrieval:** after W7's draft-first flow is working, create an empty or minimal new-message draft through Outlook and read it back through Graph. Determine whether Outlook inserts the user's configured signature into a server-created draft and, if so, whether the connector can reliably extract and reuse the signature HTML plus any associated inline images/CIDs. Prefer this zero-setup approach if it works reliably, because it would avoid requiring the user to export and provide an Outlook `.htm` signature.

The existing import design below remains the fallback if Outlook does not expose the configured signature through draft creation/read-back.

**Fallback scope:** exactly one signature for the connector: import it, remove it, enable it or disable it. No multiple-signature selection, routing rules or signature editor.

**Fallback build:**

- MCP tools: `import_signature`, `delete_signature`, `set_signature_active(true|false)`.
- Import an Outlook-style `.htm` signature and optional `<name>_files/` image folder.
- Store the one signature under the connector data directory; active by default after import.
- Send signature images as inline attachments referenced by `cid:`.
- Placement matches Outlook: new message = body then signature; reply = new text, signature, then Outlook's quoted history untouched.
- MCP only for now; no UI work required.

## W9 — Inbox rules

**Status:** Implemented on this branch; proven OWS contracts, synthetic tool validation complete.

**Current state:** the Outlook Web rule contracts have been captured and replayed successfully through `Ows.call_request`. Reading, creating, editing (including clearing a condition), enabling/disabling, reordering and deleting a throwaway rule all worked live. Graph folder ids are accepted by the write request.

**First-version scope:**

- `list_rules`
- `create_rule`
- `update_rule`, including enable/disable
- `reorder_rules`
- `delete_rule`

Support only the conditions/actions already proven and used by the mailbox: From, Sent to, Subject contains, Subject-or-body contains, Move to folder, and Stop processing more rules. Unsupported rules are listed read-only and are never rewritten.

**Safeguard:** every rule write is proposed first, requires user confirmation, is sent once, and is read back. Rules persist and affect future mail, so writes must not be retried automatically.

**Implementation notes:** stateless account/state-bound proposals, one confirmed write and fresh read-back. Enable/disable is a separate update; reordering refuses collections containing unsupported rules to avoid rewriting them. Folder read-back uses the reported name (OWS returns a mailbox path rather than a Graph id). No new live mailbox writes were performed.

# Correctness and reliability

## H17 — Partial mutation failures are reported incorrectly

**Status:** Implemented on this branch; synthetic validation complete.

**Problem:** mutations are sent in chunks. If a later chunk fails, earlier chunks may already have changed the mailbox while the tool call raises only the later error.

**Decision:** add `continue_on_error` to `set_read_state`, `set_flag`, `move_messages` and `delete_messages`, defaulting to `true`.

- A clearly failed chunk returns `failed` results for its messages.
- By default, later chunks are still attempted.
- With `continue_on_error=false`, later chunks are not sent and are returned as `failed: not sent`.
- A failed read-back after an unclear write outcome returns `unknown`.
- The call always returns per-message results and `counts`, so already-completed work is visible.

## H19 — Inline images can disappear from exports without an error

**Status:** Pending.

**Problem:** inline images are exported only when the HTML body references their `cid:`. If the body or content-id lookup fails, an image can currently be treated as unreferenced and silently skipped.

**Current evidence:** Graph can return each file attachment's `contentId` directly in the attachment-list request, so the separate per-inline-image content-id lookup is unnecessary.

**Next:** return `contentId` from the normal attachment listing, remove the extra lookup, and use a fail-open rule: if the connector cannot determine whether an inline image is referenced, include it rather than silently dropping it.

## H20 — Explicit message ids can be filtered out during export

**Status:** Recommended decision implemented on this branch; synthetic validation complete.

**Problem:** `export_messages(message_ids=[...])` reads the requested messages, then applies reach/scope filtering at the end. An explicitly supplied id from a hidden/out-of-reach folder can therefore disappear from the export without the selection behaving like `get_message(id)`.

**Recommended decision:** treat explicit message ids as authoritative, like `get_message` does. Do not apply the hidden-folder filter to that selection source; only merge/label it with the other selected messages.

**Alternative:** keep the filter but preserve and report its exclusion counts. This is stricter, but adds complexity for ids the connector itself never normally exposes.

## H21 — Long conversations cannot be marked read/unread cleanly

**Status:** Implemented on this branch; synthetic validation complete.

**Problem:** `set_read_state(conversation_ids=[...])` expands a conversation to messages and then applies the 100-message limit intended for explicit `message_ids`. A long conversation can therefore be rejected even though the caller supplied only one conversation id. Conversations beyond the 1,000-message listing cap also need an explicit truncation note.

**Next:** apply the 100 limit only to explicit `message_ids`. Expand conversation ids up to the existing conversation-listing limit, send changes in normal 20-item write chunks, and report when a conversation was truncated. Keep the response compact for very large conversations by relying on `counts` for ordinary `done`/`unchanged` results if needed.

## H22 — A filter flag alone can turn an export into a whole-mailbox range

**Status:** Pending.

**Root cause:** export has three additive selection sources: conversation ids, an optional mailbox range, and explicit message ids. `ExportRequest.by_range` is supposed to say whether the mailbox-range source is active, but it currently returns true not only for `since`, `until` or `folder`, but also when `include_sent_items=false` or `include_meeting_mail=false`.

`Exports._select()` first adds messages selected by conversation id, then calls `_select_range()` whenever `by_range` is true, then adds explicit message ids. These sources are merged; the explicit selection is not being ignored. The bug is that a filter flag accidentally activates an additional range selection. If no real range selector was supplied, `_select_range()` calls `list_messages(folder=None, since=None, until=None, ...)`, which is effectively an unbounded reachable-mailbox listing subject to the include filters.

**Example:** `export_messages(conversation_ids=["budget-thread"], include_meeting_mail=false)` should export that conversation. Today, `include_meeting_mail=false` also makes `by_range=true`, so the exporter additionally lists the whole reachable mailbox with meeting mail excluded and merges those messages with the requested conversation. On a large mailbox this can even hit the 2,000-message export cap before anything is exported.

**Next:** only `since`, `until` and `folder` activate the mailbox-range selection source. `include_meeting_mail` and `include_sent_items` remain filters for a range when one is actually requested; they must never create a range on their own or broaden an explicit conversation/message selection.

## H23 — Search dates without a timezone can crash

**Status:** Implemented on this branch; synthetic validation complete.

**Problem:** the web API can pass a naive `since`/`until` datetime into service code, which is then compared with timezone-aware Graph dates.

**Next:** normalize dates to UTC once at the service entry point and remove duplicate surface-specific normalization.

## H25 — The UI silently truncates very long message bodies

**Status:** Implemented on this branch; automatic full-body loading validated.

**Problem:** the UI reader asks for a bounded body and ignores `next_offset`, so a message over the current limit can stop mid-body without telling the user.

**Decision for this batch:** automatically follow continuation offsets until the chosen body is complete, without a new total length cap. Switching messages stops scheduling obsolete continuations using the existing reader request counter; no new in-flight cancellation infrastructure.

## H26 — Oversized attachments are classified as unexpected failures

**Status:** Pending.

**Problem:** the connector's 150 MB download guard is a known local limit, but exports currently classify it like an unexpected error.

**Why 150 MB:** keep the existing 150 MB limit as a connector policy. The code currently defines `MAX_DOWNLOAD_BYTES = 150 * 1024 * 1024` and streams attachment downloads to disk through that guard. This is not a Microsoft Graph download requirement; 150 MB is instead a conservative local ceiling that also lines up with upper Outlook/Exchange attachment/message limits in some Microsoft contexts. Most mail providers impose much tighter practical limits, so a normal file attachment reaching this guard should be extremely rare. Retaining it protects against unexpectedly huge downloads/disk usage without materially constraining ordinary email export.

**Next:** classify hitting the guard explicitly as a known connector limit rather than an unexpected failure: the attachment is larger than the connector's 150 MB download limit and should be downloaded directly from Outlook.

## H27 — Per-folder listing does not report excluded-folder counts

**Status:** Implemented in the combined H27/H28/H29/H33 branch; synthetic validation complete.

**Problem:** when the mailbox uses the per-folder listing strategy, Junk and Deleted Items are never read. That is efficient, but `coverage.excluded` therefore lacks the `deleted_or_junk` count even though those messages are outside the result.

**Next:** include the left-out folders in the existing count batch and report their counts for the requested window. No extra network round trip should be necessary.

## H28 — `server_total` includes meeting mail that the listing may hide

**Status:** Implemented in the combined H27/H28/H29/H33 branch; synthetic validation complete.

**Problem:** with `include_meeting_mail=false`, `server_total` still includes invitations, cancellations and RSVPs, so the total can be larger than the messages the listing can return.

**Constraint:** Graph's folder count endpoint cannot filter out meeting-message types with the available query shapes.

**Next:** make the coverage note explicit that `server_total` includes meeting mail even when the listing hides it. Do not add a second expensive counting path.

## H29 — NotFound wording assumes deletion

**Status:** Implemented in the combined H27/H28/H29/H33 branch; synthetic validation complete.

**Problem:** a bad or inaccessible message id currently produces wording equivalent to "not on the server (deleted on the server)", even though the id may simply be wrong or the item may have moved out of reach.

**Next:** use neutral wording: the message was not found; it may have been deleted, moved out of reach, or the id may be wrong. Remove the mailbox-level `GONE` wording that implies deletion.

## H33 — `$batch` retries fewer transient statuses than single requests

**Status:** Implemented in the combined H27/H28/H29/H33 branch; synthetic validation complete.

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

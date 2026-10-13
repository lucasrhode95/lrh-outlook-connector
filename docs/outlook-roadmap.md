# Outlook Connector Roadmap

Outstanding decisions, implementation, fixes and validation only. Completed behavior belongs in
the requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-13**.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §4 Repository layout](architecture.md#4-repository-layout)
- API evidence: [API research](outlook-api-research.md)
- Cross-product search routes and workflows: [Microsoft 365 search findings](m365-search-findings.md)

# Later

## X10 — Meeting/calendar search

**Status:** Partial — bounded Search v2 event discovery and OWS event hydration are implemented in
the Outlook connector (2026-10-12). Synthetic regression coverage and a live connector MCP check
passed for the current signed-in account. Fresh-account `m365.search` consent has not been verified.
Standalone OWA Search v2 discovery and OWS event hydration succeeded; the current Graph profile
still lacks the tested `Calendars.*` permission.

**Goal:** search Exchange calendar events (subject, time, organizer, attendees, agenda and Teams join link) and hand a Teams meeting's chat id to `lrh-teams-connector` when useful.

**Ownership:** Outlook owns the Exchange calendar event, including schedule/recurrence, agenda,
organizer, invitees, event attachments, and the virtual-meeting join link. A related Teams meeting's
chat, recordings, recap, meeting notes, and transcripts belong to `lrh-teams-connector`. An event
using another virtual-meeting platform remains an Outlook event without Teams resources.

**Required cross-product search scope:** calendar-event discovery, meeting-recap discovery,
transcript discovery/retrieval, and transcript-text keyword search are separate capabilities. Link
the event to Teams resources through verified identifiers; do not mark broad search complete with
only mail and files.

**Graph profiles (updated 2026-10-12):** `outlook.graph` has no `Calendars.*` scope, but the shared
cache's existing `teams.graph` profile has calendar-read permission. A read-only probe using that
profile returned HTTP 200 for `/me/calendarView`, Graph event Search, and event-body hydration.
No new interactive sign-in was needed. This account's result does not establish fresh-account
grants; see [research §3.18](outlook-api-research.md#318-calendar-range-listing-and-description-search-through-the-teams-client-graph-profile-2026-10-12).

**Search API documentation (2026-10-10):** Graph also documents `event` search, scoped to the
signed-in user's primary calendar, with a 25-result page limit. This does not resolve the current
`Calendars.Read` blocker or establish other-calendar coverage; see [research §3.8](outlook-api-research.md#38-microsoft-365-search-api-coverage-review-2026-10-10).

**Capture/source review (2026-10-10):** Teams recap returns `MeetingInsights`, while transcript
routes retrieve separate artifacts. Neither observation proves global transcript-text search or
resolves the calendar-scope blocker. See [research §3.9](outlook-api-research.md#39-microsoft-365-search-capture-and-source-route-map-2026-10-10)
and the [cross-product search findings](m365-search-findings.md). Transcript discovery and retrieval
are Teams-owned work and remain in the [Teams roadmap](../../lrh-teams-connector/docs/teams-connector-roadmap.md).

**Live Recaps UI check (2026-10-10):** both user-provided known recaps opened in the Teams Meeting
Recaps app with a loaded Transcript tab; in-page transcript search returned matches in one. The
download control was disabled by the app's permission boundary for one target. This confirms the signed-in web app can search a known loaded transcript. The app-owned request route, authentication, and connector-accessible transcript export remain unproven. The supplied OWA source bundle exposes a Substrate
Search scope-catalog entry and a Teams MCPS base configuration, but no concrete recap or transcript
route builder; see research §3.9.

**Standalone replay:** the existing One Outlook Web refresh token silently acquired the Search
profile. OWA-host Search v2 returned Event/Exchange results with advancing offsets and recurring
occurrences; OWS `GetItem` retrieved a selected event's agenda, attendees, and attachment. Search
reported `HasAttachments=false` for that event, and recurring-result totals did not describe the
number of returned occurrences. See [research §3.14](outlook-api-research.md#314-owa-search-v2-calendar-and-teams-message-replay-2026-10-10).

**Live connector MCP check (2026-10-12):** the current account's `m365.search` profile completed
calendar search and event hydration. A five-event page continued at offset 5; one hydrated result
included agenda text, attendees, and a Teams join URL. Date-bounded searches also produced empty
pages with `more_results_available=true`, so an empty page is not proof that the full search is
empty. The event fields expose no Teams chat id. A separate five-event hydration sample succeeded
for all five; one event returned one attachment metadata record. Event attachment content download
is not exposed by the MCP surface. Details are in
[research §3.17](outlook-api-research.md#317-live-outlook-connector-mcp-calendar-check-2026-10-12).

**Next:** establish and verify an event-to-Teams-chat reference through a supported Teams-side
lookup before offering automatic cross-product handoff. Verify fresh-account `m365.search` consent
if deploying with a new account. Search v2 totals remain intentionally omitted, event attachment
download is not exposed, and Graph `Calendars.Read` remains a separate alternative to reassess
rather than an automatic fallback. Transcript and recap discovery remain in Teams and are not
completed by calendar search.

**Calendar usability (2026-10-12, implemented):** `list_calendar_events(since, until, contains, include_body, limit, cursor)` lists the default calendar's occurrences that overlap a window, in time order and in UTC, through Graph's calendar view under the shared `teams.graph` profile (`config.CALENDAR_PROFILE`), with next links followed. Outlook Web's `GetCalendarView` answers a first page only when that Graph read fails for a service reason; it has no descriptions, so `contains` and `include_body` fail on that path. A live check on the signed-in account returned the same 16 events from both routes for a week, hydrated a listed id with `get_calendar_event`, and matched description text (research [§3.19](outlook-api-research.md#319-list_calendar_events-live-check-2026-10-12)). Still open: (a) confirm the Teams-profile calendar scopes on a fresh account before relying on them; (b) decide whether a purpose-named calendar profile in `lrh-m365-auth` should replace reusing `teams.graph` (that repository is outside this change); (c) topic search across the calendar's history: Graph event search returned 12 hits for a known description term, but its date and body-only semantics are unverified, so Search v2 remains the tool for that; (d) `get_calendar_event` still uses OWS only, and Graph-first hydration needs its own check; (e) shared and other people's calendars are not listed.

## Review follow-ups: open after the live checks (2026-10-13)

The live checks ([research §3.20](outlook-api-research.md#320-review-gap-live-checks-2026-10-13)) and the unfinished-markup save test ([research §3.21](outlook-api-research.md#321-unfinished-markup-on-save-2026-10-13)) are recorded. These items remain open:

- **Decision: one message-listing tool.** Merging `list_messages` and `search_messages` needs one response shape. A list returns flat messages in date order; search returns conversation groups in rank order. Options: one tool whose result carries a `mode` field, or one flat shape for both (search would lose conversation grouping and `message_count`). Routing is simple: an empty query lists, and a query searches. Folder, date and scope apply to both. Renaming the tools would break existing agent configurations.
- **Decision: `search_calendar_events` date bounds.** Date terms in the query returned no result set, so `since` and `until` only filter the pages fetched. Keep the bounds with the documented limit, or remove them so window questions go to `list_calendar_events`.
- **Enhancement: calendar fallback paging.** The Outlook Web fallback pages by offset over a window it re-reads. If events change between calls, a page can skip or repeat an event. Not reproduced on an unchanged window. Candidate: a keyset cursor on (start, id), tested with a change between pages.
- **Enhancement: passive HTML is a deny-list.** It relies on Python's HTML parser, so markup that tokenizes differently in a browser could pass. The one save test showed Outlook dropping an event handler, but that is one account and one client. Candidates: an allow-list, or wider sanitization evidence checked against browser parsing.
- **Limit: one conversation lists at most 1,000 messages.** `notes` reports the cap. Whether Graph pages past it is unverified.

## Outlook mail-search comparison (2026-10-12)

Three one-hit `search_messages` calls took 3.409 seconds cold, then 0.760 and 0.784 seconds warm.
They returned one conversation and one message each and were incomplete at the one-hit limit; these
end-to-end times include connector/tool overhead and do not isolate Graph latency. An equivalent
same-query comparison with Graph Search API has not been measured, so the results do not establish a
faster replacement or justify changing the current Graph mail `$search` route. If a route change is
considered later, compare identical controlled terms, result limits, hydration, request counts, and
latency. Full measurements and their limits are in [API research §3.16](outlook-api-research.md#316-outlook-mail-search-timing-and-route-comparison-2026-10-10).
## Small reliability follow-ups — completed 2026-10-11

Both synthetic, offline reproductions now have bounded fixes and regression coverage; no live
mailbox operation was made:

- `remote/graph.py` collection rejects repeated continuation URLs, including a cycle that returns
  to a previously visited page. A repeated empty page has no item-count progress, so `max_items`
  does not terminate it. Regression coverage is in `tests/unit/test_graph_collection.py`.
- `service/localfiles.py` can raise `FileNotFoundError` if another process removes an old output
  between `is_file()` and `stat()`. Cleanup now tolerates disappearance while preserving other
  filesystem errors; coverage is in `tests/unit/test_localfiles.py`.

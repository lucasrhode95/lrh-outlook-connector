# Outlook Connector Roadmap

Outstanding decisions, implementation, fixes and validation only. Completed behavior belongs in
the requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-12**.

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

**Graph limitation:** the Graph profile has no `Calendars.*` scope, so Graph `/me/calendarView` is
not available through the current Graph sign-in. This does not block the proven OWA route.

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
empty. The event fields expose no Teams chat id. Details are in
[research §3.17](outlook-api-research.md#317-live-outlook-connector-mcp-calendar-check-2026-10-12).

**Next:** establish and verify an event-to-Teams-chat reference through a supported Teams-side
lookup before offering automatic cross-product handoff. Verify fresh-account `m365.search` consent
if deploying with a new account. Search v2 totals remain intentionally omitted, event attachment
download is not exposed, and Graph `Calendars.Read` remains a separate alternative to reassess
rather than an automatic fallback. Transcript and recap discovery remain in Teams and are not
completed by calendar search.

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

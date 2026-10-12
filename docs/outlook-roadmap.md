# Outlook Connector Roadmap

Outstanding decisions, implementation, fixes and validation only. Completed behavior belongs in
the requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-12**.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §4 Repository layout](architecture.md#4-repository-layout)
- API evidence: [API research](outlook-api-research.md)
- Search investigation and checkpoint: [Microsoft 365 search plan](m365-search-investigation-plan.md)

# Later

## X11 — SharePoint file search and download

**Status:** Implemented in the separate read-only `lrh-sharepoint-connector`. Outlook retains
email-attachment handling and the historical file-access evidence recorded here. The SharePoint
connector's own roadmap records its shipped search, selected metadata hydration, bounded download,
and deployment-specific live checks. General files shared as links remain discoverable through
SharePoint; email attachments stay in Outlook and Teams message attachments stay in Teams.

**Evidence:** the existing Graph sign-in can search for and begin downloading a known SharePoint
file; unrestricted file search and known-person search also worked. Limits and unproven cases
are recorded once in [research §3.7](outlook-api-research.md#37-file-and-cross-category-search-feasibility-2026-10-07).

**Search API documentation (2026-10-10):** Graph documents `driveItem` search over content the
signed-in user can read. Content-only matching and exhaustive scope coverage still need known-hit
probes; see [research §3.8](outlook-api-research.md#38-microsoft-365-search-api-coverage-review-2026-10-10).

**Capture/source review (2026-10-10):** the personal SharePoint web client submitted a separate
browser search request with `EntityType=File`; suggestions use another route. This does not change
the Graph implementation decision or establish a standalone private API. Initial route evidence
and the next bounded questions are in [research §3.9](outlook-api-research.md#39-microsoft-365-search-capture-and-source-route-map-2026-10-10).

**Known-hit validation (2026-10-10):** direct Graph `/search/query` returned two hits for each
controlled subject-only and body-only mail term; the connector's `$search` grouped each into one
conversation at a one-result limit, with incomplete coverage. After the user confirmed OneDrive
upload, bounded Graph `driveItem` searches returned one hit each for filename-only and content-only
terms. This establishes two narrow known hits for the tested account and location; broad recall
and the effective permission basis remain unverified. Details are in
[research §3.12](outlook-api-research.md#312-controlled-microsoft-search-known-hit-checks-2026-10-10).

**Scope boundary:** the connector uses the returned drive/item reference for hydration and download;
it does not fetch arbitrary caller-supplied URLs. Its access grants and known-hit results are
deployment-specific, and broad sharing-link resolution remains a separate design question.

Keep each category's coverage and permissions explicit if a future broad discovery surface
orchestrates these providers. Graph calendar authentication remains a separate limitation (X10);
the standalone OWA calendar-search path below is proven.

## X10 — Meeting/calendar search

**Status:** Partial — bounded Search v2 event discovery and OWS event hydration are implemented in
the Outlook connector (2026-10-12). The adapter has synthetic regression coverage; a live call
through the connector and fresh-account `m365.search` consent have not been verified. Standalone
OWA Search v2 discovery and OWS event hydration succeeded; the current Graph profile still lacks
the tested `Calendars.*` permission.

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
and the [Microsoft 365 search plan](m365-search-investigation-plan.md).

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

**Next:** verify the wired path with the connector's shared-auth profile when safe authorization is
available; establish a verified event-to-Teams reference before offering automatic cross-product
handoff. Search v2 totals remain intentionally omitted, event attachment download is not exposed,
and Graph `Calendars.Read` remains a separate alternative to reassess rather than an automatic
fallback. Transcript and recap discovery remain in Teams and are not completed by calendar search.

## Microsoft 365 search follow-up (2026-10-10)

- The two known transcript-bearing Teams chat histories each include one embedded transcript URI
  reference with sanitized route shape `/drives/{drive}/items/{item}/{segment}/{segment}/media/transcripts/{transcript}/content`.
  This does not establish an HTTP request, standalone authorization, or transcript retrieval. Next:
  inspect a sanitized DevTools/Resource Saver capture from the loaded transcript viewer and map the
  two extra path segments.
- Secondary investigation work: review existing search/list/get paths for redundant calls,
  over-fetch, paging/retry fragility, and latency opportunities. Keep code simplification distinct
  from Graph migration and measure candidates before recommending implementation.
# Microsoft 365 search investigation — 2026-10-10

The cross-connector search investigation is in progress. Dated evidence, measurements, design requirements, open questions, and next steps are recorded in [the investigation log](m365-search-investigation-log-2026-10-10.md) and the [product-separated endpoint and workflow ledger](m365-search-findings-2026-10-10.md). The ledger separates Outlook, Teams, SharePoint/OneDrive, and Microsoft Graph, and links shared mechanisms. No product code changed in this research update.

## Small reliability follow-ups — completed 2026-10-11

Both synthetic, offline reproductions now have bounded fixes and regression coverage; no live
mailbox operation was made:

- `remote/graph.py` collection rejects repeated continuation URLs, including a cycle that returns
  to a previously visited page. A repeated empty page has no item-count progress, so `max_items`
  does not terminate it. Regression coverage is in `tests/unit/test_graph_collection.py`.
- `service/localfiles.py` can raise `FileNotFoundError` if another process removes an old output
  between `is_file()` and `stat()`. Cleanup now tolerates disappearance while preserving other
  filesystem errors; coverage is in `tests/unit/test_localfiles.py`.

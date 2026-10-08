# Outlook Connector Roadmap

Outstanding decisions, implementation, fixes and validation only. Completed behavior belongs in
the requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-08**.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §4 Repository layout](architecture.md#4-repository-layout)
- API evidence: [API research](outlook-api-research.md)

# Human checks

## W8 — Recipient-side signature rendering

**Status:** Open; visual verification needs a human (2026-10-08).

**Evidence:** Public-tool validation sent one synthetic signed message. Read-back retained the
signature text and inline image reference, but the strict `data-signature-name` check did not pass.
The Outlook recipient-side rendering was not visually inspected.

**Next:** Have a human inspect the synthetic signed self-send in Outlook (the copy is in Deleted
Items after cleanup) and confirm that its signature text and inline image render as expected. Do not
inspect unrelated mailbox content.

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

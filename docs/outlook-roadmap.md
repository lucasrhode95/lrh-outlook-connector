# Outlook Connector Roadmap

Outstanding decisions, implementation, fixes and validation only. Completed behavior belongs in
the requirements and architecture; completed live evidence belongs in the research record.

Snapshot **2026-10-08**.

- Product requirements: [Requirements v4](outlook-requirements-v4.md)
- Build/module map: [architecture §11](architecture.md)
- API evidence: [API research](outlook-api-research.md)

## Current priority

1. **Performance and cleanup:** H37–H38.
2. **W12:** README rewrite and architecture doc cleanup, last, once everything above is finished.

Status wording:

- **Decided, not built** — behavior is settled; implementation remains.
- **Pending** — a concrete problem and next fix are known.
- **Decision needed** — an owner choice remains.
- **Later** — useful work outside the current build sequence.

# Correctness and reliability

# Performance

# Service and code cleanup

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

**docs/outlook-roadmap.md (final review, as the very last step):**

- Keep only what is still open: pending implementation or decisions, live tests not yet run, and
  checks that need a human (e.g. recipient-side rendering). Remove everything finished, including
  W12 itself, the "Current priority" entries that no longer apply and status-wording entries no
  longer used.
- If nothing open remains, delete the file and remove every link to it (README, `CLAUDE.md` and
  `AGENTS.md`, which must stay identical, and the other docs). "Later" items such as X10 and X11 count
  as open: keep them unless the owner drops them.

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
